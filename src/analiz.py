"""İleri seviye Python analitiği – Excel'in tek başına zorlandığı matematiksel/istatistiksel testler.

Çıktılar cikti/py_*.csv olarak yazılır ve excel_model.py tarafından "Python_Bulgular" sayfasına aktarılır.

  1. Benford ailesi   : ilk hane, ikinci hane, ilk iki hane, son iki hane (ki-kare, MAD, KS) + toplam testi
  2. Isolation Forest : 9 özellikli çok değişkenli anomali skoru (scikit-learn)
  3. Varlık ağı       : ortak IBAN / adres / telefon / VKN üzerinden tedarikçi-personel bağlı bileşenleri
  4. Bulanık eşleşme  : tedarikçi–tedarikçi ve tedarikçi–yaptırım listesi ünvan benzerliği
  5. Kullanıcı davranışı: onaylayan bazında mesai dışı, limit altı yoğunlaşma, saat dağılımı ki-kare
  6. Lapping zinciri  : ödeyen→fatura sahibi yönlü graf üzerinde zincir tespiti
  7. Makbuz boşluk testi
"""
from __future__ import annotations

import json
import math
from collections import defaultdict
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
from scipy import stats

from ortak import CIKTI, DIS_VERI, VERI, fatura_numara, unvan_normalize

ONAY_LIMITI = 50_000


def yukle():
    od = pd.read_csv(VERI / "odemeler.csv", parse_dates=["OdemeTarihi", "FaturaTarihi"], dtype={"FaturaNo": str})
    th = pd.read_csv(VERI / "tahsilatlar.csv", parse_dates=["Tarih", "BankayaYatirma"])
    td = pd.read_csv(VERI / "tedarikciler.csv", parse_dates=["KayitTarihi"], dtype={"VKN": str})
    pr = pd.read_csv(VERI / "personel.csv", dtype={"TCKN": str})
    ku = pd.read_csv(VERI / "kullanicilar.csv")
    yp = pd.read_csv(DIS_VERI / "yaptirim.csv")
    return od, th, td, pr, ku, yp


# ------------------------------------------------------------------ 1) Benford
def _nigrini(mad, tur):
    esik = {"ilk": (0.006, 0.012, 0.015), "ikinci": (0.008, 0.010, 0.012), "ilkiki": (0.0012, 0.0018, 0.0022)}
    if tur not in esik:
        return "-"
    a, b, c = esik[tur]
    return "Yakın uyum" if mad <= a else "Kabul edilebilir" if mad <= b else "Sınırda" if mad <= c else "UYUMSUZ"


def benford(tutar: pd.Series):
    x = tutar[tutar >= 10].astype(float)
    s = x.map(lambda v: f"{v:.2f}".replace(".", "").lstrip("0"))
    sonuc, detay = [], []

    def test(ad, tur, gozlenen, beklenen_p, etiketler):
        n = gozlenen.sum()
        gp = gozlenen / n
        chi2 = float(((gozlenen - n * beklenen_p) ** 2 / (n * beklenen_p)).sum())
        df = len(etiketler) - 1
        mad = float(np.abs(gp - beklenen_p).mean())
        ks = float(np.abs(np.cumsum(gp) - np.cumsum(beklenen_p)).max())
        sonuc.append(dict(Test=ad, n=int(n), KiKare=round(chi2, 2), sd=df, p_degeri=float(stats.chi2.sf(chi2, df)),
                          MAD=round(mad, 5), Nigrini=_nigrini(mad, tur), KS=round(ks, 4),
                          KS_kritik_05=round(1.36 / math.sqrt(n), 4)))
        z = (np.abs(gp - beklenen_p) - 1 / (2 * n)) / np.sqrt(beklenen_p * (1 - beklenen_p) / n)
        for e, g, b, zz in zip(etiketler, gp, beklenen_p, z):
            detay.append(dict(Test=ad, Hane=e, Gozlenen=round(g, 5), Beklenen=round(b, 5), Z=round(float(zz), 2),
                              Anlamli=bool(zz > 1.96)))

    d1 = np.arange(1, 10)
    test("İlk hane", "ilk", s.str[0].astype(int).value_counts().reindex(d1, fill_value=0).values,
         np.log10(1 + 1 / d1), d1)
    d2 = np.arange(10)
    p2 = np.array([sum(math.log10(1 + 1 / (10 * a + b)) for a in range(1, 10)) for b in d2])
    test("İkinci hane", "ikinci", s.str[1].astype(int).value_counts().reindex(d2, fill_value=0).values, p2, d2)
    d12 = np.arange(10, 100)
    test("İlk iki hane", "ilkiki", s.str[:2].astype(int).value_counts().reindex(d12, fill_value=0).values,
         np.log10(1 + 1 / d12), d12)
    son = (x.astype(int) % 100)
    test("Son iki hane (tam kısım)", "son", son.value_counts().reindex(range(100), fill_value=0).values,
         np.full(100, 0.01), range(100))
    # Toplam (summation) testi: her ilk-iki-hane grubunun toplamı eşit olmalı (Nigrini Ch.7)
    grp = x.groupby(s.str[:2].astype(int)).sum().reindex(d12, fill_value=0)
    pay = grp / grp.sum()
    toplam = pd.DataFrame(dict(IlkIkiHane=d12, ToplamPayi=pay.values.round(5), Beklenen=round(1 / 90, 5),
                               Kat=(pay.values * 90).round(2))).sort_values("Kat", ascending=False).head(10)
    return pd.DataFrame(sonuc), pd.DataFrame(detay), toplam


# --------------------------------------------------------- 2) Isolation Forest
def ozellikler(od, td, ku):
    df = od.copy()
    df["ln_tl"] = np.log(df["TutarTL"].clip(lower=1))
    df["saat"] = df["Saat"].str[:2].astype(int) + df["Saat"].str[3:].astype(int) / 60
    df["haftagunu"] = df["OdemeTarihi"].dt.weekday
    kayit = td.set_index("TedarikciID")["KayitTarihi"]
    df["kayittan_gun"] = (df["OdemeTarihi"] - df["TedarikciID"].map(kayit)).dt.days
    g = df.groupby("TedarikciID")["ln_tl"]
    df["ted_z"] = ((df["ln_tl"] - g.transform("mean")) / g.transform("std").replace(0, np.nan)).fillna(0)
    df["yuvarlaklik"] = df["BrutTutar"].map(lambda v: len(str(int(round(v)))) - len(str(int(round(v))).rstrip("0")))
    df["ayni_tutar"] = df.groupby(["TedarikciID", "BrutTutar"])["OdemeID"].transform("count")
    df["fatura_odeme_gun"] = (df["OdemeTarihi"] - df["FaturaTarihi"]).dt.days
    lim = ku.set_index("KullaniciID")["OnayLimitiTL"]
    df["limit_orani"] = (df["TutarTL"] / df["Onaylayan"].map(lim).replace(0, np.nan)).fillna(5).clip(upper=5)
    return df, ["ln_tl", "saat", "haftagunu", "kayittan_gun", "ted_z", "yuvarlaklik", "ayni_tutar",
                "fatura_odeme_gun", "limit_orani"]


def isolation_forest(df, kolonlar):
    from sklearn.ensemble import IsolationForest
    from sklearn.preprocessing import RobustScaler
    X = RobustScaler().fit_transform(df[kolonlar].astype(float))
    model = IsolationForest(n_estimators=400, contamination=0.04, random_state=7).fit(X)
    df = df.copy()
    df["IF_skor"] = (-model.score_samples(X)).round(4)
    df["IF_aykiri"] = model.predict(X) == -1
    Xz = pd.DataFrame(X, columns=kolonlar)
    df["Etken_ozellikler"] = [", ".join(Xz.iloc[i].abs().sort_values(ascending=False).index[:3])
                              for i in range(len(Xz))]
    # Robust Mahalanobis (MCD yerine medyan/IQR ölçekli Öklid) – açıklanabilir ek skor
    df["Robust_uzaklik"] = np.sqrt((X ** 2).sum(axis=1)).round(3)
    cikti = df.sort_values("IF_skor", ascending=False)[
        ["OdemeID", "TedarikciID", "OdemeTarihi", "Saat", "TutarTL", "Onaylayan", "IF_skor", "IF_aykiri",
         "Robust_uzaklik", "Etken_ozellikler"]]
    cikti["Sira"] = range(1, len(cikti) + 1)
    return cikti


# ---------------------------------------------------------------- 3) Varlık ağı
def varlik_agi(td, pr):
    ebeveyn = {}

    def bul(a):
        ebeveyn.setdefault(a, a)
        while ebeveyn[a] != a:
            ebeveyn[a] = ebeveyn[ebeveyn[a]]
            a = ebeveyn[a]
        return a

    def birlestir(a, b):
        ebeveyn[bul(a)] = bul(b)

    baglar = defaultdict(set)
    kayitlar = [("T:" + r.TedarikciID, {"IBAN": r.IBAN, "Adres": r.Adres, "Telefon": str(r.Telefon), "VKN": r.VKN})
                for r in td.itertuples()]
    kayitlar += [("P:" + r.PersonelID, {"IBAN": r.IBAN, "Adres": r.Adres, "Telefon": str(r.Telefon)})
                 for r in pr.itertuples()]
    indeks = defaultdict(list)
    for kim, ozel in kayitlar:
        bul(kim)
        for tur, deger in ozel.items():
            indeks[(tur, deger)].append(kim)
    for (tur, _), kimler in indeks.items():
        for k in kimler[1:]:
            birlestir(kimler[0], k)
            baglar[kimler[0]].add(tur)
            baglar[k].add(tur)
    kumeler = defaultdict(list)
    for kim, _ in kayitlar:
        kumeler[bul(kim)].append(kim)
    satirlar = []
    for i, uyeler in enumerate(sorted((u for u in kumeler.values() if len(u) > 1), key=len, reverse=True), 1):
        turler = sorted({t for u in uyeler for t in baglar[u]})
        t_var, p_var = any(u.startswith("T:") for u in uyeler), any(u.startswith("P:") for u in uyeler)
        karma = t_var and p_var
        satirlar.append(dict(Kume=f"K{i:02d}", Uyeler=", ".join(sorted(uyeler)), UyeSayisi=len(uyeler),
                             BagTurleri=", ".join(turler),
                             Risk="YÜKSEK – personel/tedarikçi bağı" if karma else
                             ("Orta – tedarikçiler arası" if t_var else "Orta – personeller arası (hayalet personel?)")))
    return pd.DataFrame(satirlar)


# ------------------------------------------------------------ 4) Bulanık eşleşme
def bulanik(td, yp, esik=0.86):
    satirlar = []
    norm = [(r.TedarikciID, r.Unvan, unvan_normalize(r.Unvan)) for r in td.itertuples()]
    for i in range(len(norm)):
        for j in range(i + 1, len(norm)):
            b = SequenceMatcher(None, norm[i][2], norm[j][2]).ratio()
            if b >= esik:
                satirlar.append(dict(Tur="Tedarikçi–Tedarikçi", A=f"{norm[i][0]} {norm[i][1]}",
                                     B=f"{norm[j][0]} {norm[j][1]}", Benzerlik=round(b, 3)))
    for tid, ad, n in norm:
        for y in yp.itertuples():
            b = SequenceMatcher(None, n, str(y.Norm)).ratio()
            # token kesişimi: ilk iki anlamlı kelime ortaksa da yakala
            ortak = len(set(n.split()[:3]) & set(str(y.Norm).split()[:3]))
            if b >= 0.80 or ortak >= 2:
                satirlar.append(dict(Tur=f"Yaptırım ({y.Liste})", A=f"{tid} {ad}", B=y.Ad, Benzerlik=round(b, 3)))
    return pd.DataFrame(satirlar, columns=["Tur", "A", "B", "Benzerlik"])


# -------------------------------------------------------- 5) Kullanıcı davranışı
def kullanici(od, ku):
    df = od.copy()
    df["saat"] = df["Saat"].str[:2].astype(int)
    df["mesai_disi"] = (df["saat"] < 8) | (df["saat"] >= 19)
    df["limit_alti"] = (df["TutarTL"] >= ONAY_LIMITI * 0.9) & (df["TutarTL"] < ONAY_LIMITI)
    df["sod"] = df["Olusturan"] == df["Onaylayan"]
    genel = df["saat"].value_counts(normalize=True).sort_index()
    satirlar = []
    for u, g in df.groupby("Onaylayan"):
        gozlenen = g["saat"].value_counts().reindex(genel.index, fill_value=0)
        beklenen = genel * len(g)
        maske = beklenen >= 1
        chi2 = float(((gozlenen[maske] - beklenen[maske]) ** 2 / beklenen[maske]).sum())
        p = float(stats.chi2.sf(chi2, max(maske.sum() - 1, 1)))
        satirlar.append(dict(Onaylayan=u, Islem=len(g), ToplamTL=round(g["TutarTL"].sum(), 2),
                             MesaiDisiOran=round(g["mesai_disi"].mean(), 3), LimitAltiOran=round(g["limit_alti"].mean(), 3),
                             SoD_Adet=int(g["sod"].sum()), SaatDagilimi_p=round(p, 4),
                             Yorum="İncele" if (p < 0.01 or g["sod"].sum() > 3 or g["limit_alti"].mean() > 0.08) else ""))
    return pd.DataFrame(satirlar).sort_values("SoD_Adet", ascending=False)


# ---------------------------------------------------------- 6) Lapping zinciri
def lapping(th):
    kenar = th[th["OdeyenMusteriID"] != th["FaturaMusteriID"]].sort_values("Tarih")
    if kenar.empty:
        return pd.DataFrame(columns=["Zincir", "Uzunluk", "Musteriler", "Tutar", "Kasiyerler", "Baslangic", "Bitis"])
    sonraki = dict(zip(kenar["FaturaMusteriID"], kenar["OdeyenMusteriID"]))
    baslar = set(sonraki) - set(sonraki.values())
    satirlar = []
    for i, b in enumerate(sorted(baslar), 1):
        yol, a = [b], b
        while a in sonraki and sonraki[a] not in yol:
            a = sonraki[a]
            yol.append(a)
        alt = kenar[kenar["FaturaMusteriID"].isin(yol[:-1])]
        satirlar.append(dict(Zincir=f"L{i:02d}", Uzunluk=len(yol) - 1, Musteriler=" → ".join(yol),
                             Tutar=round(alt["TahsilatTutari"].sum(), 2), Kasiyerler=", ".join(sorted(alt["Kasiyer"].unique())),
                             Baslangic=alt["Tarih"].min().date(), Bitis=alt["Tarih"].max().date()))
    return pd.DataFrame(satirlar)


# ------------------------------------------------------------- 7) Boşluk testi
def bosluk(th):
    m = np.sort(th["MakbuzNo"].unique())
    fark = np.diff(m)
    satirlar = [dict(Onceki=int(a), Sonraki=int(b), EksikAdet=int(b - a - 1)) for a, b, f in zip(m[:-1], m[1:], fark) if f > 1]
    return pd.DataFrame(satirlar, columns=["Onceki", "Sonraki", "EksikAdet"])


def calistir():
    od, th, td, pr, ku, yp = yukle()
    b_ozet, b_detay, b_toplam = benford(od["TutarTL"])
    df, kol = ozellikler(od, td, ku)
    iso = isolation_forest(df, kol)
    ag = varlik_agi(td, pr)
    fz = bulanik(td, yp)
    kul = kullanici(od, ku)
    lap = lapping(th)
    bos = bosluk(th)
    for ad, d in [("py_benford_ozet", b_ozet), ("py_benford_detay", b_detay), ("py_benford_toplam", b_toplam),
                  ("py_isolation", iso), ("py_ag", ag), ("py_bulanik", fz), ("py_kullanici", kul),
                  ("py_lapping", lap), ("py_bosluk", bos)]:
        d.to_csv(CIKTI / f"{ad}.csv", index=False, encoding="utf-8")
    ozet = dict(odeme=len(od), tahsilat=len(th), IF_aykiri=int(iso["IF_aykiri"].sum()), ag_kume=len(ag),
                bulanik=len(fz), lapping_zincir=len(lap), makbuz_bosluk=len(bos),
                benford_ilk_hane=b_ozet.iloc[0].to_dict())
    (CIKTI / "py_ozet.json").write_text(json.dumps(ozet, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(ozet, ensure_ascii=False, indent=1, default=str))
    print(b_ozet.to_string())
    print(ag.to_string())
    print(fz.to_string())
    print(lap.to_string())
    print(kul.to_string())


if __name__ == "__main__":
    calistir()
