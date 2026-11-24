"""Ortak yardımcılar: dizinler, kimlik/IBAN doğrulama algoritmaları, metin normalizasyonu.

Bu algoritmalar Excel modelindeki formüllerin (VKN, TCKN, IBAN mod-97) birebir Python karşılığıdır;
böylece Excel sonuçları Python ile çapraz doğrulanabilir.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent
VERI = KOK / "veri"
DIS_VERI = KOK / "dis_veri"
YEDEK = DIS_VERI / "yedek"
CIKTI = KOK / "cikti"

for _d in (VERI, DIS_VERI, YEDEK, CIKTI):
    _d.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------------------------------- VKN
def vkn_kontrol_hanesi(ilk9: str) -> int:
    """Vergi Kimlik No (10 hane) kontrol hanesi – GİB algoritması."""
    toplam = 0
    for j in range(1, 10):
        t = (int(ilk9[j - 1]) + 10 - j) % 10
        v = (t * 2 ** (10 - j)) % 9
        if t == 9:
            v = 9
        toplam += v
    return (10 - toplam % 10) % 10


def vkn_gecerli(vkn: str) -> bool:
    vkn = str(vkn)
    return len(vkn) == 10 and vkn.isdigit() and vkn_kontrol_hanesi(vkn[:9]) == int(vkn[9])


# -------------------------------------------------------------------------- TCKN
def tckn_tamamla(ilk9: str) -> str:
    d = [int(c) for c in ilk9]
    d10 = ((d[0] + d[2] + d[4] + d[6] + d[8]) * 7 - (d[1] + d[3] + d[5] + d[7])) % 10
    d11 = (sum(d) + d10) % 10
    return ilk9 + str(d10) + str(d11)


def tckn_gecerli(tckn: str) -> bool:
    tckn = str(tckn)
    if len(tckn) != 11 or not tckn.isdigit() or tckn[0] == "0":
        return False
    return tckn_tamamla(tckn[:9]) == tckn


# -------------------------------------------------------------------------- IBAN
def _iban_mod97(iban: str) -> int:
    yeniden = iban[4:] + iban[:4]
    sayisal = "".join(str(int(c, 36)) for c in yeniden)
    return int(sayisal) % 97


def iban_olustur(banka_kodu: str, hesap: str) -> str:
    """TR IBAN: TR + 2 kontrol + 5 banka + 1 rezerv(0) + 16 hesap."""
    bban = f"{banka_kodu:0>5}0{hesap:0>16}"
    kontrol = 98 - _iban_mod97("TR00" + bban)
    return f"TR{kontrol:02d}{bban}"


def iban_gecerli(iban: str) -> bool:
    iban = str(iban).replace(" ", "").upper()
    return bool(re.fullmatch(r"TR\d{24}", iban)) and _iban_mod97(iban) == 1


# ---------------------------------------------------------------- Normalizasyon
_SIRKET_EKLERI = [
    "LTD. ŞTİ.", "LTD.ŞTİ.", "LTD ŞTİ", "LIMITED SIRKETI", "A.Ş.", "A.S.", " AŞ", " AS ", "SAN. VE TİC.",
    "SAN.", "TİC.", "LTD", "ŞTİ", "INC", "LLC", "FZE", "CO.", "CORP", "GMBH",
]

_TR_HARF = str.maketrans("ÇĞİIÖŞÜçğıiöşü", "CGIIOSUcgiiosu")


def unvan_normalize(ad: str) -> str:
    """Ünvanı karşılaştırılabilir çekirdeğe indirger (büyük harf, Türkçe karakter ve şirket eki temizliği)."""
    s = str(ad).upper()
    for ek in _SIRKET_EKLERI:
        s = s.replace(ek, " ")
    s = s.translate(_TR_HARF)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def fatura_normalize(no: str) -> str:
    return re.sub(r"[-\s/.]", "", str(no).upper())


def fatura_numara(no: str) -> int:
    n = fatura_normalize(no)[3:]
    return int(n) if n.isdigit() else 0
