"""Web kazıma (scraping) modülü – modeli dış veriyle besler.

Kaynaklar
---------
1. TCMB gösterge kurları (günlük XML arşivi)          -> dis_veri/kurlar.csv        (T14 kur testi)
2. Resmi tatiller (Nager.Date açık API)               -> dis_veri/tatiller.csv      (T04 / H05)
3. OFAC SDN yaptırım listesi (CSV)                    -> dis_veri/yaptirim.csv      (G08)
4. Reddit topluluk tartışmaları (r/Accounting, r/Audit, r/FraudPrevention, r/forensicaccounting)
                                                      -> dis_veri/topluluk.csv      (test önceliklendirme)
5. Big Four / ACFE kamuya açık forensic sayfaları     -> dis_veri/topluluk.csv

Her kaynak bağımsızdır: erişilemezse dis_veri/yedek/ altındaki çevrimdışı yedek kullanılır ve
dis_veri/kazima_log.csv dosyasına "YEDEK" olarak yazılır. Yedek kur verisi SENTETİKTİR; gerçek
denetimde mutlaka canlı kazıma yapın.

Etik / hukuki: robots.txt kontrol edilir, istekler arasında bekleme yapılır, kişisel veri toplanmaz
(Reddit'ten yalnızca başlık, gövde metni, puan ve bağlantı alınır; kullanıcı adı alınmaz).
Reddit'te üretim kullanımı için OAuth API (praw) ve Reddit Data API şartları önerilir.

Kullanım:  python src/kazima.py [--yil 2025] [--cevrimdisi]
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import math
import random
import re
import time
import urllib.robotparser
import xml.etree.ElementTree as ET
from urllib.parse import urlparse

from ortak import DIS_VERI, YEDEK, unvan_normalize
from testler import TUM_TESTLER

UA = "Mozilla/5.0 (compatible; KontrolDenetimArastirma/1.0; +internal-audit-research)"
BEKLEME = 0.35  # saniye – kibar kazıma
_LOG: list[dict] = []

# Kamuya açık forensic / hile içerikleri. URL'ler zamanla değişebilir: kaynaklar.json ile güncelleyin.
VARSAYILAN_SAYFALAR = [
    ("PwC", "https://www.pwc.com/gx/en/services/forensics/economic-crime-survey.html"),
    ("Deloitte", "https://www2.deloitte.com/us/en/pages/advisory/solutions/forensic-data-analytics.html"),
    ("EY", "https://www.ey.com/en_gl/services/forensic-integrity-services"),
    ("KPMG", "https://kpmg.com/xx/en/what-we-do/services/advisory/forensic.html"),
    ("ACFE", "https://legacy.acfe.com/fraud-tree.aspx"),
]
SUBREDDITLER = ["Accounting", "Audit", "FraudPrevention", "forensicaccounting", "InternalAudit"]
REDDIT_SORGULARI = ["fraud red flags", "ghost vendor", "duplicate payments", "lapping", "benford",
                    "split purchase", "data analytics audit"]


# ----------------------------------------------------------------------------- HTTP
def _oturum():
    import requests
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "tr,en;q=0.8"})
    return s


_ROBOTS: dict[str, urllib.robotparser.RobotFileParser | None] = {}


def _robots_izin(url: str, oturum) -> bool:
    kok = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if kok not in _ROBOTS:
        rp = urllib.robotparser.RobotFileParser()
        try:
            r = oturum.get(kok + "/robots.txt", timeout=8)
            rp.parse(r.text.splitlines() if r.status_code == 200 else [])
            _ROBOTS[kok] = rp
        except Exception:
            _ROBOTS[kok] = None  # robots okunamadı -> temkinli: izin verildi say ama logla
    rp = _ROBOTS[kok]
    return True if rp is None else rp.can_fetch(UA, url)


def _getir(oturum, url: str, **kw):
    if not _robots_izin(url, oturum):
        raise PermissionError(f"robots.txt izin vermiyor: {url}")
    time.sleep(BEKLEME)
    r = oturum.get(url, timeout=kw.pop("timeout", 20), **kw)
    r.raise_for_status()
    return r


def _log(kaynak, url, durum, satir, not_=""):
    _LOG.append(dict(zaman=dt.datetime.now().isoformat(timespec="seconds"), kaynak=kaynak, url=url,
                     durum=durum, satir=satir, aciklama=not_))
    print(f"[{durum:6}] {kaynak:22} {satir:>6} satır  {not_}")


def _yaz(ad, satirlar, alanlar):
    with open(DIS_VERI / ad, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=alanlar)
        w.writeheader()
        w.writerows(satirlar)


def _oku(yol):
    with open(yol, encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ------------------------------------------------------------------------ 1) TCMB
def tcmb_kurlar(yil: int, oturum, tatiller: set[dt.date]):
    """TCMB arşivi: https://www.tcmb.gov.tr/kurlar/YYYYMM/DDMMYYYY.xml  (ForexBuying)."""
    satirlar, ardisik_hata = [], 0
    gun = dt.date(yil - 1, 12, 20)  # yıl başı hafta sonu/tatil için önceki iş günü kuru gerekir
    while gun <= dt.date(yil, 12, 31):
        if ardisik_hata >= 8:
            raise RuntimeError("TCMB art arda 8 kez yanıt vermedi")
        if gun.weekday() < 5 and gun not in tatiller:
            url = f"https://www.tcmb.gov.tr/kurlar/{gun:%Y%m}/{gun:%d%m%Y}.xml"
            try:
                kok = ET.fromstring(_getir(oturum, url).content)
                kur = {}
                for c in kok.findall("Currency"):
                    kod = c.get("CurrencyCode")
                    if kod in ("USD", "EUR", "GBP"):
                        birim = float(c.findtext("Unit") or 1)
                        kur[kod] = round(float(c.findtext("ForexBuying")) / birim, 4)
                if len(kur) == 3:
                    satirlar.append(dict(Tarih=gun.isoformat(), **kur))
                ardisik_hata = 0
            except Exception:
                ardisik_hata += 1  # TCMB yayımlamadığı günler (ör. idari izin) 404 döner
        gun += dt.timedelta(days=1)
    if len(satirlar) < 50:
        raise RuntimeError("TCMB'den yeterli veri alınamadı")
    return satirlar


def yedek_kurlar(yil: int, tatiller: set[dt.date]):
    """SENTETİK kur serisi (yalnızca çevrimdışı demo için; gerçek TCMB verisi DEĞİLDİR)."""
    rnd = random.Random(42)
    usd, eur, gbp = 35.30, 36.70, 44.20
    satirlar = []
    gun = dt.date(yil - 1, 12, 20)
    while gun <= dt.date(yil, 12, 31):
        if gun.weekday() < 5 and gun not in tatiller:
            usd *= math.exp(0.00075 + rnd.gauss(0, 0.0025))
            eur *= math.exp(0.00095 + rnd.gauss(0, 0.0045))
            gbp *= math.exp(0.00080 + rnd.gauss(0, 0.0045))
            satirlar.append(dict(Tarih=gun.isoformat(), USD=round(usd, 4), EUR=round(eur, 4), GBP=round(gbp, 4)))
        gun += dt.timedelta(days=1)
    return satirlar


# --------------------------------------------------------------------- 2) Tatiller
def nager_tatiller(yil: int, oturum):
    satirlar = []
    for y in (yil - 1, yil, yil + 1):
        for h in _getir(oturum, f"https://date.nager.at/api/v3/PublicHolidays/{y}/TR").json():
            satirlar.append(dict(Tarih=h["date"], Ad=h.get("localName") or h["name"], Tur="Resmi tatil"))
    return satirlar


def yedek_tatiller():
    return _oku(YEDEK / "tatiller.csv")


# --------------------------------------------------------------------- 3) Yaptırım
def ofac_sdn(oturum):
    r = _getir(oturum, "https://www.treasury.gov/ofac/downloads/sdn.csv", timeout=60)
    satirlar = []
    for row in csv.reader(io.StringIO(r.text)):
        if len(row) > 3 and row[2].strip().strip('"') in ("-0-", ""):  # SDN_Type boş = kuruluş
            ad = row[1].strip()
            satirlar.append(dict(Liste="OFAC SDN", Ad=ad, Program=row[3].strip(), Norm=unvan_normalize(ad)))
    if not satirlar:
        raise RuntimeError("SDN listesi boş")
    return satirlar


def yedek_yaptirim():
    return _oku(YEDEK / "yaptirim.csv")


# ------------------------------------------------- 4-5) Reddit + Big Four metin madenciliği
def _test_eslestir(metin: str) -> str:
    m = metin.lower()
    return ";".join(t["kod"] for t in TUM_TESTLER if any(k.lower() in m for k in t["anahtar"]))


def reddit_kazi(oturum, limit=50):
    satirlar, gorulen, hata = [], set(), 0
    for sub in SUBREDDITLER:
        for q in REDDIT_SORGULARI:
            url = f"https://www.reddit.com/r/{sub}/search.json"
            if hata >= 5 and not satirlar:
                raise RuntimeError("Reddit art arda 5 kez yanıt vermedi")
            try:
                veri = _getir(oturum, url, params=dict(q=q, restrict_sr=1, sort="top", t="all", limit=limit)).json()
            except Exception:
                hata += 1
                continue
            for c in veri.get("data", {}).get("children", []):
                d = c["data"]
                if d["id"] in gorulen:
                    continue
                gorulen.add(d["id"])
                metin = f"{d.get('title', '')} {d.get('selftext', '')}"
                satirlar.append(dict(Kaynak=f"Reddit r/{sub}", Baslik=d.get("title", "")[:250],
                                     Puan=d.get("score", 0), Yorum=d.get("num_comments", 0),
                                     URL="https://www.reddit.com" + d.get("permalink", ""),
                                     Testler=_test_eslestir(metin)))
    if not satirlar:
        raise RuntimeError("Reddit'ten veri alınamadı")
    return satirlar


def sayfa_kazi(oturum, sayfalar):
    from bs4 import BeautifulSoup
    satirlar = []
    for kurum, url in sayfalar:
        try:
            soup = BeautifulSoup(_getir(oturum, url).text, "lxml")
            for etiket in soup(["script", "style", "nav", "footer"]):
                etiket.decompose()
            baslik = (soup.title.string or kurum).strip() if soup.title else kurum
            paragraflar = [p.get_text(" ", strip=True) for p in soup.find_all(["p", "li", "h2", "h3"])]
            for p in paragraflar:
                eslesen = _test_eslestir(p)
                if eslesen:
                    satirlar.append(dict(Kaynak=kurum, Baslik=f"{baslik}: {p[:200]}", Puan="", Yorum="",
                                         URL=url, Testler=eslesen))
            _log(f"Sayfa {kurum}", url, "CANLI", len(paragraflar))
        except Exception as e:
            _log(f"Sayfa {kurum}", url, "HATA", 0, str(e)[:120])
    return satirlar


# ---------------------------------------------------------------------- Orkestrasyon
def calistir(yil: int = 2025, cevrimdisi: bool = False):
    oturum = None if cevrimdisi else _oturum()

    def dene(ad, url, canli, yedek, dosya, alanlar):
        if not cevrimdisi:
            try:
                s = canli()
                _yaz(dosya, s, alanlar)
                _log(ad, url, "CANLI", len(s))
                return s
            except Exception as e:
                hata = str(e)[:120]
        else:
            hata = "çevrimdışı mod"
        s = yedek()
        _yaz(dosya, s, alanlar)
        _log(ad, url, "YEDEK", len(s), hata)
        return s

    tat = dene("Resmi tatiller", "date.nager.at", lambda: nager_tatiller(yil, oturum), yedek_tatiller,
               "tatiller.csv", ["Tarih", "Ad", "Tur"])
    tatil_set = {dt.date.fromisoformat(t["Tarih"]) for t in tat}
    dene("TCMB kurları", "tcmb.gov.tr/kurlar", lambda: tcmb_kurlar(yil, oturum, tatil_set),
         lambda: yedek_kurlar(yil, tatil_set), "kurlar.csv", ["Tarih", "USD", "EUR", "GBP"])
    dene("OFAC SDN", "treasury.gov/ofac", lambda: ofac_sdn(oturum), yedek_yaptirim, "yaptirim.csv",
         ["Liste", "Ad", "Program", "Norm"])

    alanlar = ["Kaynak", "Baslik", "Puan", "Yorum", "URL", "Testler"]
    topluluk = []
    if not cevrimdisi:
        try:
            r = reddit_kazi(oturum)
            topluluk += r
            _log("Reddit", "reddit.com/r/…/search.json", "CANLI", len(r))
        except Exception as e:
            _log("Reddit", "reddit.com", "HATA", 0, str(e)[:120])
        sayfalar = VARSAYILAN_SAYFALAR
        ozel = DIS_VERI / "kaynaklar.json"
        if ozel.exists():
            sayfalar = [tuple(x) for x in json.loads(ozel.read_text(encoding="utf-8"))]
        topluluk += sayfa_kazi(oturum, sayfalar)
    else:
        _log("Reddit + Big Four", "-", "YOK", 0, "çevrimdışı: topluluk sinyali üretilmedi (uydurma veri yazılmaz)")
    _yaz("topluluk.csv", topluluk, alanlar)
    _yaz("kazima_log.csv", _LOG, ["zaman", "kaynak", "url", "durum", "satir", "aciklama"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--yil", type=int, default=2025)
    ap.add_argument("--cevrimdisi", action="store_true", help="İnternete çıkmadan yedek verileri kullan")
    a = ap.parse_args()
    calistir(a.yil, a.cevrimdisi)
