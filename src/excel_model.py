"""Excel hile denetim modeli oluşturucu (XlsxWriter).

Tüm test sonuçları canlı Excel formülleridir: veri sayfalarındaki satırları kendi ERP dökümünüzle
değiştirip yeniden hesaplattığınızda bayraklar, skorlar ve pano güncellenir.

* Klasik katman  : Excel 2010+ ve LibreOffice'te çalışan formüller (COUNTIFS, SUMIFS, SUMPRODUCT,
                   AGGREGATE, FREQUENCY, FORECAST.LINEAR, CORREL, SUBTOTAL, CHIDIST, ...)
* Dinamik katman : "Dinamik_365" sayfası – Microsoft 365 dinamik dizi fonksiyonları
                   (TAKE, DROP, VSTACK, HSTACK, TOCOL, TOROW, SORTBY, GROUPBY, PIVOTBY, LAMBDA, MAP,
                   BYROW, SCAN, LET, FILTER, UNIQUE, XLOOKUP, REDUCE) ve Ad Yöneticisi'nde tanımlı
                   LAMBDA fonksiyonları (VKN_GECERLI, TCKN_GECERLI, IBAN_GECERLI, ZSKOR, BENFORD_P).
"""
from __future__ import annotations

import csv
import datetime as dt
import numbers
import re

import numpy as np

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as harf

from ortak import CIKTI, DIS_VERI, VERI
from testler import (ISTATISTIK_TESTLERI, ODEME_TESTLERI, ORGANIZASYON_TESTLERI, PERSONEL_TESTLERI,
                     TAHSILAT_TESTLERI, TUM_TESTLER)

DOSYA = CIKTI / "Hile_Denetim_Modeli.xlsx"

# ------------------------------------------------------------------ formül ön-ekleri
_XLFN = ["TAKE", "DROP", "VSTACK", "HSTACK", "TOCOL", "TOROW", "SORTBY", "GROUPBY", "PIVOTBY", "LAMBDA",
         "MAP", "BYROW", "BYCOL", "SCAN", "REDUCE", "LET", "UNIQUE", "SEQUENCE", "XLOOKUP", "CHOOSECOLS",
         "CHOOSEROWS", "FORECAST.LINEAR", "AGGREGATE", "STDEV.S", "TEXTJOIN", "MAXIFS", "QUARTILE.INC",
         "PERCENTILE.INC", "CHISQ.DIST.RT", "IFS"]
_XLWS = ["FILTER", "SORT"]


def fx(f: str) -> str:
    """Excel 2010 sonrası fonksiyonlara dosya biçiminin gerektirdiği _xlfn./_xlws. ön-eklerini,
    LET/LAMBDA parametrelerine (§ad) _xlpm. ön-ekini ekler."""
    for ad in _XLWS:
        f = re.sub(rf"(?<![\w.]){ad}\(", f"_xlfn._xlws.{ad}(", f)
    for ad in _XLFN:
        f = re.sub(rf"(?<![\w.]){re.escape(ad)}\(", f"_xlfn.{ad}(", f)
    return f.replace("§", "_xlpm.")


# ------------------------------------------------------------------ veri
def oku(yol):
    with open(yol, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def tarih(s):
    return dt.datetime.fromisoformat(s) if s else None


OD = oku(VERI / "odemeler.csv")
TH = oku(VERI / "tahsilatlar.csv")
TD = oku(VERI / "tedarikciler.csv")
PR = oku(VERI / "personel.csv")
KU = oku(VERI / "kullanicilar.csv")
KUR = oku(DIS_VERI / "kurlar.csv")
TAT = oku(DIS_VERI / "tatiller.csv")
YAP = oku(DIS_VERI / "yaptirim.csv")
TOP = oku(DIS_VERI / "topluluk.csv")
LOG = oku(DIS_VERI / "kazima_log.csv")
NO, NH, NT, NP, NK = len(OD), len(TH), len(TD), len(PR), len(KU)
CANLI = {r["kaynak"]: r["durum"] for r in LOG}

wb = xlsxwriter.Workbook(DOSYA, {"default_format_properties": {"font_name": "Arial", "font_size": 10},
                                 "strings_to_numbers": False})
wb.set_properties({"title": "Hile Denetim ve Anomali Analitiği Modeli", "subject": "İç denetim / suistimal denetimi",
                   "comments": "Kontrol projesi – src/excel_model.py ile üretilmiştir"})
wb.set_calc_mode("auto")

# ------------------------------------------------------------------ biçimler
LACI, ACIK, GRI = "#1F3864", "#D9E1F2", "#F2F2F2"


def F(**k):
    return wb.add_format(k)


f_baslik = F(bold=True, font_size=16, font_color=LACI)
f_alt = F(italic=True, font_color="#595959")
f_bolum = F(bold=True, font_size=12, font_color="white", bg_color=LACI)
f_h = F(bold=True, font_color="white", bg_color=LACI, text_wrap=True, valign="vcenter", border=1)
f_h2 = F(bold=True, bg_color=ACIK, text_wrap=True, valign="vcenter", border=1)
f_metin = F(text_wrap=True, valign="top")
f_mavi = F(font_color="#0000FF")
f_mavi_tarih = F(font_color="#0000FF", num_format="dd.mm.yyyy")
f_mavi_saat = F(font_color="#0000FF", num_format="hh:mm")
f_mavi_tl = F(font_color="#0000FF", num_format="#,##0.00")
f_mavi_kur = F(font_color="#0000FF", num_format="0.0000")
f_mavi_yuzde = F(font_color="#0000FF", num_format="0%")
f_girdi = F(font_color="#0000FF", bg_color="#FFFF00", border=1)
f_girdi_tl = F(font_color="#0000FF", bg_color="#FFFF00", border=1, num_format="#,##0")
f_girdi_yuzde = F(font_color="#0000FF", bg_color="#FFFF00", border=1, num_format="0.0%")
f_girdi_saat = F(font_color="#0000FF", bg_color="#FFFF00", border=1, num_format="hh:mm")
f_girdi_ondalik = F(font_color="#0000FF", bg_color="#FFFF00", border=1, num_format="0.000")
f_yesil = F(font_color="#008000")
f_yesil_tl = F(font_color="#008000", num_format="#,##0.00")
f_yesil_tarih = F(font_color="#008000", num_format="dd.mm.yyyy")
f_tl = F(num_format="#,##0.00;(#,##0.00);-")
f_tl0 = F(num_format="#,##0;(#,##0);-")
f_tarih = F(num_format="dd.mm.yyyy")
f_yuzde = F(num_format="0.0%;-0.0%;-")
f_yuzde2 = F(num_format="0.00%")
f_ondalik = F(num_format="0.000")
f_ondalik4 = F(num_format="0.0000")
f_ondalik5 = F(num_format="0.00000")
f_tam = F(num_format="0")
f_bayrak = F(num_format="0;-0;-", align="center")
f_skor = F(num_format="0", bold=True, align="center")
f_kpi_ad = F(bold=True, font_color="#595959", text_wrap=True, valign="vcenter", bg_color=GRI, border=1)
f_kpi = F(bold=True, font_size=14, font_color=LACI, num_format="#,##0", align="center", valign="vcenter", border=1)
f_kpi_tl = F(bold=True, font_size=14, font_color=LACI, num_format="#,##0", align="center", valign="vcenter", border=1)
f_kpi_txt = F(bold=True, font_size=12, font_color=LACI, align="center", valign="vcenter", border=1, text_wrap=True)
f_kod = F(font_name="Consolas", font_size=9, font_color="#404040", text_wrap=True, valign="top", bg_color=GRI)
f_kirmizi = F(bg_color="#F8CBAD", font_color="#9C0006")
f_turuncu = F(bg_color="#FFE699", font_color="#7F6000")
f_yesil_bg = F(bg_color="#C6EFCE", font_color="#006100")
f_bilimsel = F(num_format="0.00E+00")
f_not = F(italic=True, font_color="#7F7F7F", text_wrap=True, valign="top")


def ad(ad_, ref):
    wb.define_name(ad_, "=" + ref)


def aralik(sayfa, kol, n, bas=2):
    s = f"'{sayfa}'" if " " in sayfa else sayfa
    return f"{s}!${kol}${bas}:${kol}${bas + n - 1}"


# ================================================================== VERİ SAYFALARI
def veri_sayfasi(isim, satirlar, kolonlar, aciklama):
    """kolonlar: [(alan, başlık, tür, genişlik)] tür: s|n|tl|d|t|kur|pct"""
    ws = wb.add_worksheet(isim)
    ws.set_tab_color("#808080")
    for j, (_, b, _, w) in enumerate(kolonlar):
        ws.write(0, j, b, f_h)
        ws.set_column(j, j, w)
    for i, r in enumerate(satirlar, 1):
        for j, (alan, _, tur, _) in enumerate(kolonlar):
            v = r.get(alan, "")
            if v in ("", None):
                continue
            if tur == "s":
                ws.write_string(i, j, str(v), f_mavi)
            elif tur == "d":
                ws.write_datetime(i, j, tarih(v), f_mavi_tarih)
            elif tur == "t":
                h, m = v.split(":")
                ws.write_number(i, j, (int(h) * 60 + int(m)) / 1440, f_mavi_saat)
            elif tur == "tl":
                ws.write_number(i, j, float(v), f_mavi_tl)
            elif tur == "kur":
                ws.write_number(i, j, float(v), f_mavi_kur)
            elif tur == "pct":
                ws.write_number(i, j, float(v), f_mavi_yuzde)
            else:
                ws.write_number(i, j, float(v), f_mavi)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, max(len(satirlar), 1), len(kolonlar) - 1)
    ws.write_comment(0, 0, aciklama, {"x_scale": 3, "y_scale": 2})
    return ws


# sayfaları ekleme sırası = sekme sırası: önce analiz sayfaları oluşturulur
ws_kil = wb.add_worksheet("Kilavuz")
ws_pan = wb.add_worksheet("Risk_Panosu")
ws_par = wb.add_worksheet("Parametreler")
ws_to = wb.add_worksheet("T_Odeme")
ws_th = wb.add_worksheet("T_Tahsilat")
ws_tg = wb.add_worksheet("T_Org")
ws_tp = wb.add_worksheet("T_Personel")
ws_bf = wb.add_worksheet("Benford")
ws_is = wb.add_worksheet("Istatistik")
ws_dn = wb.add_worksheet("Dinamik_365")
ws_py = wb.add_worksheet("Python_Bulgular")
ws_pif = wb.add_worksheet("Py_IsolationForest")
ws_lib = wb.add_worksheet("Test_Kutuphanesi")

veri_sayfasi("Odeme", OD, [
    ("OdemeID", "Ödeme ID", "s", 10), ("OdemeTarihi", "Ödeme Tarihi", "d", 11), ("Saat", "Saat", "t", 7),
    ("TedarikciID", "Tedarikçi ID", "s", 10), ("FaturaNo", "Fatura No", "s", 13), ("FaturaTarihi", "Fatura Tarihi", "d", 11),
    ("Miktar", "Miktar", "n", 8), ("BirimFiyat", "Birim Fiyat", "tl", 12), ("Matrah", "Matrah", "tl", 13),
    ("KDVOrani", "KDV Oranı", "pct", 8), ("KDV", "KDV", "tl", 11), ("BrutTutar", "Brüt Tutar (döviz)", "tl", 14),
    ("ParaBirimi", "Para Birimi", "s", 7), ("Kur", "Kullanılan Kur", "kur", 10), ("TutarTL", "Tutar (TL)", "tl", 14),
    ("Olusturan", "Oluşturan", "s", 9), ("Onaylayan", "Onaylayan", "s", 9), ("OdemeYontemi", "Yöntem", "s", 11),
    ("MasrafMerkezi", "Masraf Merkezi", "s", 15), ("Aciklama", "Açıklama", "s", 22)],
    "Ödeme (borç) hareketleri. Kendi ERP dökümünüzü aynı sütun sırasıyla yapıştırın; satır sayısı değişirse "
    "src/excel_model.py ile modeli yeniden üretin (adlandırılmış aralıklar satır sayısına göre kurulur).")
veri_sayfasi("Tahsilat", TH, [
    ("TahsilatID", "Tahsilat ID", "s", 10), ("MakbuzNo", "Makbuz No", "n", 10), ("Tarih", "Tarih", "d", 11),
    ("OdeyenMusteriID", "Ödeyen Müşteri", "s", 10), ("FaturaNo", "Kapatılan Fatura", "s", 12),
    ("FaturaMusteriID", "Fatura Müşterisi", "s", 10), ("FaturaTutari", "Fatura Tutarı (TL)", "tl", 14),
    ("TahsilatTutari", "Tahsilat Tutarı (TL)", "tl", 14), ("Yontem", "Yöntem", "s", 11),
    ("BankayaYatirma", "Bankaya Yatırma", "d", 12), ("Kasiyer", "Kasiyer", "s", 8)],
    "Tahsilat (alacak) hareketleri.")
veri_sayfasi("Tedarikci", TD, [
    ("TedarikciID", "Tedarikçi ID", "s", 10), ("Unvan", "Ünvan", "s", 34), ("VKN", "VKN", "s", 12), ("IBAN", "IBAN", "s", 28),
    ("Adres", "Adres", "s", 34), ("Telefon", "Telefon", "s", 12), ("KayitTarihi", "Kayıt Tarihi", "d", 11),
    ("Olusturan", "Kaydı Açan", "s", 9), ("Durum", "Durum", "s", 8), ("Kategori", "Kategori", "s", 12),
    ("ParaBirimi", "Para Birimi", "s", 7), ("SozlesmeLimitiTL", "Sözleşme Limiti (TL)", "tl", 15)],
    "Tedarikçi ana verisi (vendor master).")
veri_sayfasi("Personel", PR, [
    ("PersonelID", "Personel ID", "s", 10), ("AdSoyad", "Ad Soyad", "s", 18), ("TCKN", "TCKN", "s", 13),
    ("IBAN", "IBAN", "s", 28), ("Adres", "Adres", "s", 34), ("Telefon", "Telefon", "s", 12),
    ("Departman", "Departman", "s", 11), ("IseGiris", "İşe Giriş", "d", 11)], "Personel ana verisi.")
veri_sayfasi("Kullanici", KU, [
    ("KullaniciID", "Kullanıcı ID", "s", 10), ("AdSoyad", "Ad Soyad", "s", 18), ("Rol", "Rol", "s", 22),
    ("OnayLimitiTL", "Onay Limiti (TL)", "tl", 15)], "ERP kullanıcıları ve onay yetki matrisi.")
ws_kur = veri_sayfasi("Kur", KUR, [("Tarih", "Tarih", "d", 11), ("USD", "USD", "kur", 9), ("EUR", "EUR", "kur", 9),
                                   ("GBP", "GBP", "kur", 9)],
                      "TCMB döviz alış kurları – src/kazima.py ile kazınır. Durum: " + CANLI.get("TCMB kurları", "?") +
                      " (YEDEK = sentetik demo serisi, gerçek TCMB verisi değildir).")
veri_sayfasi("Tatil", TAT, [("Tarih", "Tarih", "d", 11), ("Ad", "Tatil", "s", 36), ("Tur", "Tür", "s", 12)],
             "Resmi tatiller – Nager.Date API'den kazınır. Durum: " + CANLI.get("Resmi tatiller", "?"))
ws_yap = veri_sayfasi("Yaptirim", YAP, [("Liste", "Liste", "s", 22), ("Ad", "Ad", "s", 36), ("Program", "Program", "s", 12),
                                        ("Norm", "Normalize Ad", "s", 34)],
                      "Yaptırım listesi – OFAC SDN'den kazınır. Durum: " + CANLI.get("OFAC SDN", "?") +
                      " (YEDEK = kurgusal örnek isimler, gerçek liste değildir).")
NY = len(YAP)
ws_yap.write(0, 4, "Eşleşme Anahtarı (ilk 2 kelime)", f_h)
ws_yap.set_column(4, 4, 28)
for i in range(2, NY + 2):
    ws_yap.write_formula(f"E{i}", fx(f'=IFERROR(LEFT(D{i},FIND(" ",D{i},FIND(" ",D{i})+1)-1),D{i})'))
veri_sayfasi("Topluluk", TOP, [("Kaynak", "Kaynak", "s", 20), ("Baslik", "Başlık / Paragraf", "s", 70), ("Puan", "Puan", "n", 7),
                               ("Yorum", "Yorum", "n", 7), ("URL", "URL", "s", 40), ("Testler", "Eşleşen Testler", "s", 18)],
             "Reddit (r/Accounting, r/Audit, r/FraudPrevention, r/forensicaccounting, r/InternalAudit) ve Big Four/ACFE "
             "kamuya açık sayfalarından kazınan metinler; testler.py anahtar kelimeleriyle eşleştirilir. "
             "Çevrimdışı çalıştırmada boş kalır (uydurma topluluk verisi yazılmaz).")
veri_sayfasi("Kazima_Log", LOG, [("zaman", "Zaman", "s", 20), ("kaynak", "Kaynak", "s", 22), ("url", "URL", "s", 32),
                                 ("durum", "Durum", "s", 8), ("satir", "Satır", "n", 7), ("aciklama", "Açıklama", "s", 60)],
             "Web kazıma çalıştırma günlüğü.")

# ------------------------------------------------------------------ adlandırılmış aralıklar
for k, col in dict(O_ID="A", O_Tarih="B", O_Saat="C", O_Ted="D", O_Fatura="E", O_FTarih="F", O_Miktar="G",
                   O_BF="H", O_Matrah="I", O_KDVOran="J", O_KDV="K", O_Brut="L", O_PB="M", O_Kur="N", O_TL="O",
                   O_Olusturan="P", O_Onaylayan="Q").items():
    ad(k, aralik("Odeme", col, NO))
for k, col in dict(H_ID="A", H_Makbuz="B", H_Tarih="C", H_Musteri="D", H_Fatura="E", H_FMusteri="F",
                   H_FTutar="G", H_Tutar="H", H_Yontem="I", H_Yatirma="J", H_Kasiyer="K").items():
    ad(k, aralik("Tahsilat", col, NH))
for k, col in dict(Ted_ID="A", Ted_Unvan="B", Ted_VKN="C", Ted_IBAN="D", Ted_Adres="E", Ted_Tel="F",
                   Ted_Kayit="G", Ted_Olusturan="H", Ted_Durum="I", Ted_Sozlesme="L").items():
    ad(k, aralik("Tedarikci", col, NT))
for k, col in dict(P_ID="A", P_Ad="B", P_TCKN="C", P_IBAN="D", P_Adres="E", P_Tel="F").items():
    ad(k, aralik("Personel", col, NP))
ad("K_ID", aralik("Kullanici", "A", NK))
ad("K_Limit", aralik("Kullanici", "D", NK))
ad("Kur_Tarih", aralik("Kur", "A", len(KUR)))
ad("Kur_Tablo", f"Kur!$B$2:$D${len(KUR) + 1}")
ad("Kur_Baslik", "Kur!$B$1:$D$1")
ad("Tatil_Tarih", aralik("Tatil", "A", len(TAT)))
ad("Y_Norm", aralik("Yaptirim", "D", NY))
ad("Y_Anahtar", aralik("Yaptirim", "E", NY))

# ================================================================== PARAMETRELER
P = ws_par
P.set_column("A:A", 34); P.set_column("B:B", 14); P.set_column("C:C", 90)
P.write("A1", "Parametreler ve Test Ağırlıkları", f_baslik)
P.write("A2", "Sarı zeminli mavi hücreler kullanıcı girdisidir. Tüm testler bu hücrelere adla bağlıdır "
              "(Formüller › Ad Yöneticisi).", f_alt)
P.write_row("A3", ["Parametre", "Değer", "Açıklama / kaynak"], f_h)
PARAM = [
    ("Onay_Limiti", "Onay limiti (TL)", 50_000, f_girdi_tl, "Tek imza ile onaylanabilecek üst tutar – kurum yetki matrisinden (varsayım)."),
    ("Yakinlik", "Limit altı yakınlık oranı", 0.10, f_girdi_yuzde, "Limitin %X altına kadar olan tutarlar 'limit altı' sayılır (T07)."),
    ("Bolme_Gun", "Bölünmüş ödeme pencere (± gün)", 5, f_girdi, "Aynı tedarikçiye bu pencere içindeki limit altı ödemeler toplanır (T08)."),
    ("Bolme_AltOran", "Bölünmüş ödeme alt oranı", 0.75, f_girdi_yuzde, "Yalnızca limitin bu oranından büyük parçalar dikkate alınır (gürültüyü azaltır)."),
    ("Yuvarlak_Birim", "Yuvarlak tutar birimi", 1_000, f_girdi_tl, "MOD(tutar; birim)=0 ise yuvarlak (T06, H08)."),
    ("Mesai_Bas", "Mesai başlangıcı", 8 / 24, f_girdi_saat, "Bu saatten önceki kayıtlar mesai dışı (T05)."),
    ("Mesai_Bit", "Mesai bitişi", 19 / 24, f_girdi_saat, "Bu saat ve sonrası mesai dışı (T05)."),
    ("Hesap_Tol", "Hesaplama toleransı (TL)", 1, f_girdi_tl, "Yeniden hesaplama testlerinde kuruş yuvarlama payı (T11–T14, H06)."),
    ("Kur_Tol", "Kur sapma toleransı", 0.005, f_girdi_yuzde, "Kullanılan kur / TCMB kuru − 1 mutlak değeri (T14)."),
    ("Z_Esik", "Z skoru eşiği", 3, f_girdi_ondalik, "Tedarikçi içi leave-one-out log-tutar Z skoru (T19)."),
    ("IQR_K", "IQR çarpanı (uç çit)", 3, f_girdi_ondalik, "Q3 + k×IQR üstü aşırı aykırı (T20). Tukey: 1,5 iç / 3 dış çit."),
    ("Tekrar_Esik", "Aynı tutar tekrar eşiği", 3, f_girdi, "Aynı tedarikçiye aynı brüt tutar en az N kez (T21)."),
    ("Yeni_Gun", "Yeni tedarikçi penceresi (gün)", 30, f_girdi, "Kayıttan sonraki N gün (T22)."),
    ("Yeni_Tutar", "Yeni tedarikçi tutar eşiği (TL)", 100_000, f_girdi_tl, "T22 için tutar eşiği."),
    ("Ardisik_Esik", "Ardışık fatura oranı eşiği", 0.50, f_girdi_yuzde, "n / (maks no − min no + 1) ≥ eşik → kabuk şirket örüntüsü (G10/T23)."),
    ("Ardisik_MinN", "Ardışık fatura min. ödeme adedi", 6, f_girdi, "Küçük örneklerde yanlış pozitifi önler."),
    ("RSF_Esik", "RSF eşiği", 8, f_girdi_ondalik, "En büyük / ikinci en büyük ödeme (Nigrini). T25."),
    ("Benford_MinN", "Benford min. örnek (tedarikçi)", 30, f_girdi, "Tedarikçi bazlı Benford için asgari işlem sayısı."),
    ("Benford_pEsik", "Benford ki-kare p eşiği", 0.001, f_girdi_ondalik, "Tedarikçi ilk hane ki-kare p-değeri bu eşiğin altındaysa uyumsuz (G11/T24)."),
    ("Yuvarlak_Oran", "Yuvarlak tutar oranı eşiği", 0.30, f_girdi_yuzde, "Tedarikçi ödemelerinin bu oranından fazlası yuvarlaksa (G12)."),
    ("Nakit_Gun", "Nakit yatırma süresi (gün)", 2, f_girdi, "Nakit tahsilatın bankaya yatırılması için azami gün (H02)."),
    ("Nakit_Esik", "Nakit raporlama eşiği (TL)", 30_000, f_girdi_tl, "VARSAYIM – güncel mevzuata/kurum politikasına göre güncelleyin (H03)."),
    ("Nakit_Yakinlik", "Nakit eşik yakınlık oranı", 0.10, f_girdi_yuzde, "Eşiğin %X altı yapılandırma şüphesi (H03)."),
    ("Tahmin_Sapma", "Aylık tahmin sapma eşiği", 0.30, f_girdi_yuzde, "|Gerçekleşen/FORECAST.LINEAR − 1| (S04)."),
    ("Yuksek_Esik", "Yüksek risk skoru eşiği", 8, f_girdi, "Ağırlıklı skor ≥ eşik → Yüksek."),
    ("Orta_Esik", "Orta risk skoru eşiği", 4, f_girdi, "Ağırlıklı skor ≥ eşik → Orta."),
]
for i, (n, a, v, fm, acik) in enumerate(PARAM, 4):
    P.write(f"A{i}", a)
    P.write(f"B{i}", v, fm)
    P.write(f"C{i}", acik, f_not)
    ad(n, f"Parametreler!$B${i}")
r0 = 4 + len(PARAM) + 2
P.merge_range(r0 - 1, 0, r0 - 1, 2, "Test ağırlıkları (risk skoru = Σ bayrak × ağırlık)", f_bolum)
P.write_row(r0, 0, ["Kod – Test", "Ağırlık", "Kategori / ACFE şeması"], f_h)
AG_SATIR = {}
for i, t in enumerate(ODEME_TESTLERI + TAHSILAT_TESTLERI, r0 + 1):
    P.write(i, 0, f"{t['kod']} – {t['ad']}")
    P.write(i, 1, t["agirlik"], f_girdi)
    P.write(i, 2, f"{t['kategori']} · {t['acfe']}", f_not)
    AG_SATIR[t["kod"]] = i + 1
rv = r0 + len(ODEME_TESTLERI) + len(TAHSILAT_TESTLERI) + 3
P.write(rv - 1, 0, "Ağırlık vektörleri (otomatik – SUMPRODUCT için yatay; düzenlemeyin)", f_h2)
P.write(rv, 0, "Ödeme testleri T01–T26", f_not)
for j, t in enumerate(ODEME_TESTLERI):
    P.write_formula(rv, 1 + j, f"=B{AG_SATIR[t['kod']]}")
P.write(rv + 1, 0, "Tahsilat testleri H01–H08", f_not)
for j, t in enumerate(TAHSILAT_TESTLERI):
    P.write_formula(rv + 1, 1 + j, f"=B{AG_SATIR[t['kod']]}")
ad("W_Odeme", f"Parametreler!$B${rv + 1}:${harf(len(ODEME_TESTLERI))}${rv + 1}")
ad("W_Tahsilat", f"Parametreler!$B${rv + 2}:${harf(len(TAHSILAT_TESTLERI))}${rv + 2}")
P.freeze_panes(3, 0)


# ================================================================== KOLON ŞABLONU MOTORU
def kolon_sayfasi(ws, spec, n, ad_onek, baslik_not=None):
    """spec: [dict(k, h, f, fmt, w, ad)] – f içinde @ satır no, [[k]] başka kolonun harfi."""
    harita = {s["k"]: harf(j) for j, s in enumerate(spec)}

    def coz(f, r):
        f = re.sub(r"\[\[(\w+)\]\]", lambda m: harita[m.group(1)], f)
        return fx(f.replace("@", str(r)))

    for j, s in enumerate(spec):
        ws.write(0, j, s["h"], f_h)
        ws.set_column(j, j, s.get("w", 9), None, {"hidden": s.get("gizli", False)})
        if s.get("aciklama"):
            ws.write_comment(0, j, s["aciklama"], {"x_scale": 2.5, "y_scale": 1.6})
        for r in range(2, n + 2):
            ws.write_formula(r - 1, j, coz(s["f"], r), s.get("fmt"))
        if s.get("ad"):
            ad(f"{ad_onek}_{s['ad']}", f"{ws.name}!${harita[s['k']]}$2:${harita[s['k']]}${n + 1}")
    ws.set_row(0, 42)
    ws.freeze_panes(1, 2)
    ws.autofilter(0, 0, n, len(spec) - 1)
    return harita


def bayrak_bicim(ws, ilk, son, n):
    ws.conditional_format(1, ilk, n, son, {"type": "cell", "criteria": "==", "value": 1, "format": f_kirmizi})


def seviye_bicim(ws, kol, n):
    ws.conditional_format(1, kol, n, kol, {"type": "cell", "criteria": "==", "value": '"Yüksek"', "format": f_kirmizi})
    ws.conditional_format(1, kol, n, kol, {"type": "cell", "criteria": "==", "value": '"Orta"', "format": f_turuncu})


# ================================================================== T_ODEME
TO_TEST_F = {
    "T01": "=--(COUNTIFS(O_Ted,[[ted]]@,O_Fatura,Odeme!$E@,O_Brut,Odeme!$L@)>1)",
    "T02": "=--(COUNTIFS(TO_Ted,[[ted]]@,TO_FatNum,[[fnum]]@,O_Brut,Odeme!$L@)>COUNTIFS(O_Ted,[[ted]]@,O_Fatura,Odeme!$E@,O_Brut,Odeme!$L@))",
    "T03": "=--([[hg]]@>5)",
    "T04": "=--(COUNTIF(Tatil_Tarih,[[tar]]@)>0)",
    "T05": "=--(OR(Odeme!$C@<Mesai_Bas,Odeme!$C@>=Mesai_Bit))",
    "T06": "=--(MOD(Odeme!$L@,Yuvarlak_Birim)=0)",
    "T07": "=--(AND([[tl]]@>=Onay_Limiti*(1-Yakinlik),[[tl]]@<Onay_Limiti))",
    "T08": ('=--(AND([[tl]]@<Onay_Limiti,[[tl]]@>=Onay_Limiti*Bolme_AltOran,'
            'COUNTIFS(O_Ted,[[ted]]@,O_Tarih,">="&([[tar]]@-Bolme_Gun),O_Tarih,"<="&([[tar]]@+Bolme_Gun),O_TL,">="&Onay_Limiti*Bolme_AltOran,O_TL,"<"&Onay_Limiti)>=2,'
            'SUMIFS(O_TL,O_Ted,[[ted]]@,O_Tarih,">="&([[tar]]@-Bolme_Gun),O_Tarih,"<="&([[tar]]@+Bolme_Gun),O_TL,">="&Onay_Limiti*Bolme_AltOran,O_TL,"<"&Onay_Limiti)>=Onay_Limiti))'),
    "T09": "=--(Odeme!$P@=Odeme!$Q@)",
    "T10": "=--([[tl]]@>INDEX(K_Limit,MATCH(Odeme!$Q@,K_ID,0)))",
    "T11": "=--(ABS(Odeme!$G@*Odeme!$H@-Odeme!$I@)>Hesap_Tol)",
    "T12": "=--(ABS(Odeme!$I@*Odeme!$J@-Odeme!$K@)>Hesap_Tol)",
    "T13": "=--(ABS(Odeme!$I@+Odeme!$K@-Odeme!$L@)>Hesap_Tol)",
    "T14": "=--(OR(ABS(Odeme!$N@/[[refkur]]@-1)>Kur_Tol,ABS(Odeme!$L@*Odeme!$N@-[[tl]]@)>Hesap_Tol))",
    "T15": "=--(Odeme!$F@>[[tar]]@)",
    "T16": "=--([[tar]]@<INDEX(Ted_Kayit,[[tsatir]]@))",
    "T17": '=--(INDEX(Ted_Durum,[[tsatir]]@)="Pasif")',
    "T18": "=--(INDEX(TOrg_Skor,[[tsatir]]@)>0)",
    "T19": "=--(ABS([[z]]@)>Z_Esik)",
    "T20": "=--([[tl]]@>IQR_Ust)",
    "T21": "=--(COUNTIFS(O_Ted,[[ted]]@,O_Brut,Odeme!$L@)>=Tekrar_Esik)",
    "T22": "=--(AND([[tar]]@>=INDEX(Ted_Kayit,[[tsatir]]@),[[tar]]@-INDEX(Ted_Kayit,[[tsatir]]@)<=Yeni_Gun,[[tl]]@>=Yeni_Tutar))",
    "T23": "=--(INDEX(TOrg_G10,[[tsatir]]@)=1)",
    "T24": "=--(INDEX(TOrg_G11,[[tsatir]]@)=1)",
    "T25": "=--(AND(INDEX(TOrg_RSF,[[tsatir]]@)>RSF_Esik,[[tl]]@=INDEX(TOrg_Max,[[tsatir]]@)))",
    "T26": "=--([[kum]]@>INDEX(Ted_Sozlesme,[[tsatir]]@))",
}
spec_to = [
    dict(k="id", h="Ödeme ID", f="=Odeme!A@", fmt=f_yesil, w=10, ad="ID"),
    dict(k="ted", h="Tedarikçi", f="=Odeme!D@", fmt=f_yesil, w=9, ad="Ted"),
    dict(k="tl", h="Tutar (TL)", f="=Odeme!O@", fmt=f_yesil_tl, w=13, ad="TL"),
    dict(k="tar", h="Tarih", f="=Odeme!B@", fmt=f_yesil_tarih, w=10, ad="Tarih"),
    dict(k="fnorm", h="Fatura No (normalize)", f='=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER(Odeme!E@),"-","")," ",""),"/",""),".","")',
         w=12, aciklama="Büyük harf; '-', boşluk, '/', '.' temizlenir – bulanık mükerrer anahtarı."),
    dict(k="fnum", h="Fatura Sayısal", f="=IFERROR(VALUE(MID([[fnorm]]@,4,20)),0)", fmt=f_tam, w=9, ad="FatNum",
         aciklama="Ön ek (3 harf) sonrası sayısal kısım – baştaki sıfırlar düşer."),
    dict(k="ilk", h="İlk Hane", f='=IF([[tl]]@>=10,VALUE(LEFT(INT([[tl]]@),1)),"")', fmt=f_tam, w=6, ad="Ilk"),
    dict(k="ilkiki", h="İlk İki Hane", f='=IF([[tl]]@>=10,VALUE(LEFT(INT([[tl]]@),2)),"")', fmt=f_tam, w=6, ad="IlkIki"),
    dict(k="son", h="Son İki Hane", f="=MOD(INT([[tl]]@),100)", fmt=f_tam, w=6, ad="Son"),
    dict(k="ln", h="ln(Tutar)", f="=LN(MAX([[tl]]@,1))", fmt=f_ondalik, w=7, ad="Ln"),
    dict(k="ay", h="Ay", f="=MONTH([[tar]]@)", fmt=f_tam, w=5, ad="Ay"),
    dict(k="refkur", h="TCMB Kuru", f='=IF(Odeme!$M@="TRY",1,INDEX(Kur_Tablo,MATCH([[tar]]@,Kur_Tarih,1),MATCH(Odeme!$M@,Kur_Baslik,0)))',
         fmt=f_ondalik4, w=9, aciklama="Ödeme tarihindeki (veya önceki son iş günündeki) kazınmış TCMB kuru."),
    dict(k="tsatir", h="Ted. Satır", f="=MATCH([[ted]]@,Ted_ID,0)", fmt=f_tam, w=6,
         aciklama="Tedarikçinin Tedarikci/T_Org sayfasındaki sıra numarası (INDEX için)."),
    dict(k="tn", h="Ted. n", f="=INDEX(TOrg_N,[[tsatir]]@)", fmt=f_tam, w=5, gizli=True),
    dict(k="loom", h="LOO ort. ln", f="=IF([[tn]]@>=5,([[tn]]@*INDEX(TOrg_OrtLn,[[tsatir]]@)-[[ln]]@)/([[tn]]@-1),0)",
         fmt=f_ondalik, w=7, gizli=True, aciklama="Bu kayıt hariç (leave-one-out) tedarikçi ortalaması."),
    dict(k="loov", h="LOO varyans",
         f="=IF([[tn]]@>=5,(([[tn]]@-1)*INDEX(TOrg_SdLn,[[tsatir]]@)^2+[[tn]]@*INDEX(TOrg_OrtLn,[[tsatir]]@)^2-[[ln]]@^2-([[tn]]@-1)*[[loom]]@^2)/([[tn]]@-2),0)",
         fmt=f_ondalik, w=7, gizli=True, aciklama="Bu kayıt hariç varyans: (Σx² − x² − (n−1)·m₋ᵢ²)/(n−2)."),
    dict(k="z", h="Log-Z (LOO, ted. içi)", f="=IF([[loov]]@>0,([[ln]]@-[[loom]]@)/SQRT([[loov]]@),0)",
         fmt=f_ondalik, w=7, ad="Z",
         aciklama="Studentize (leave-one-out) Z: kayıt kendi ortalama/sapmasını şişiremez; aykırı değer maskelemesini önler."),
    dict(k="hg", h="Hafta Günü (1=Pzt)", f="=WEEKDAY([[tar]]@,2)", fmt=f_tam, w=6, ad="HaftaGunu"),
    dict(k="kum", h="Kümülatif Ted. TL", f='=SUMIFS(O_TL,O_Ted,[[ted]]@,O_Tarih,"<="&[[tar]]@)', fmt=f_tl0, w=12),
]
TEST_ILK = len(spec_to)
for t in ODEME_TESTLERI:
    spec_to.append(dict(k=t["kod"], h=t["kod"], f=TO_TEST_F[t["kod"]], fmt=f_bayrak, w=4.5,
                        ad=t["kod"], aciklama=f"{t['ad']}\n{t['mantik']}"))
TEST_SON = len(spec_to) - 1
spec_to += [
    dict(k="adet", h="Test Sayısı", f=f"=SUM([[T01]]@:[[T26]]@)", fmt=f_tam, w=6, ad="Adet"),
    dict(k="skor", h="Risk Skoru", f="=SUMPRODUCT([[T01]]@:[[T26]]@,W_Odeme)", fmt=f_skor, w=7, ad="Skor"),
    dict(k="sev", h="Risk Seviyesi", f='=IF([[skor]]@>=Yuksek_Esik,"Yüksek",IF([[skor]]@>=Orta_Esik,"Orta","Düşük"))', w=8, ad="Seviye"),
    dict(k="sira", h="Sıralama Anahtarı", f="=[[skor]]@+ROW()/10000000", fmt=f_ondalik, w=8, ad="SiraAnahtar", gizli=True),
    dict(k="pyif", h="Python IF Skoru", f='=IFERROR(INDEX(PY_IF_Skor,MATCH([[id]]@,PY_IF_ID,0)),"")', fmt=f_ondalik, w=8, ad="PyIF",
         aciklama="Isolation Forest anomali skoru (Python). Kural tabanlı skorla karşılaştırma için."),
    dict(k="pyayk", h="Python IF Aykırı", f='=IFERROR(IF(INDEX(PY_IF_Ayk,MATCH([[id]]@,PY_IF_ID,0)),1,0),"")', fmt=f_bayrak, w=7, ad="PyAyk"),
]
H_TO = kolon_sayfasi(ws_to, spec_to, NO, "TO")
ad("TO_Bayraklar", f"T_Odeme!${H_TO['T01']}$2:${H_TO['T26']}${NO + 1}")
ad("TO_Kodlar", f"T_Odeme!${H_TO['T01']}$1:${H_TO['T26']}$1")
bayrak_bicim(ws_to, TEST_ILK, TEST_SON, NO)
bayrak_bicim(ws_to, TEST_SON + 6, TEST_SON + 6, NO)
seviye_bicim(ws_to, TEST_SON + 3, NO)
ws_to.conditional_format(1, TEST_SON + 2, NO, TEST_SON + 2, {"type": "data_bar", "bar_color": "#C00000", "bar_solid": True})
ws_to.set_tab_color("#C00000")

# ================================================================== T_TAHSILAT
TH_TEST_F = {
    "H01": "=--(Tahsilat!$D@<>Tahsilat!$F@)",
    "H02": '=--(AND(Tahsilat!$I@="Nakit",Tahsilat!$J@-Tahsilat!$C@>Nakit_Gun))',
    "H03": '=--(AND(Tahsilat!$I@="Nakit",[[tl]]@>=Nakit_Esik*(1-Nakit_Yakinlik),[[tl]]@<Nakit_Esik))',
    "H04": "=--(AND(Tahsilat!$B@>MIN(H_Makbuz),COUNTIF(H_Makbuz,Tahsilat!$B@-1)=0))",
    "H05": "=--(OR(WEEKDAY([[tar]]@,2)>5,COUNTIF(Tatil_Tarih,[[tar]]@)>0))",
    "H06": "=--([[tl]]@>Tahsilat!$G@+Hesap_Tol)",
    "H07": "=--(COUNTIFS(H_Musteri,Tahsilat!$D@,H_Fatura,Tahsilat!$E@,H_Tutar,[[tl]]@,H_Tarih,[[tar]]@)>1)",
    "H08": '=--(AND(Tahsilat!$I@="Nakit",MOD([[tl]]@,Yuvarlak_Birim)=0))',
}
spec_th = [
    dict(k="id", h="Tahsilat ID", f="=Tahsilat!A@", fmt=f_yesil, w=10, ad="ID"),
    dict(k="mus", h="Ödeyen Müşteri", f="=Tahsilat!D@", fmt=f_yesil, w=9, ad="Musteri"),
    dict(k="tl", h="Tutar (TL)", f="=Tahsilat!H@", fmt=f_yesil_tl, w=13, ad="TL"),
    dict(k="tar", h="Tarih", f="=Tahsilat!C@", fmt=f_yesil_tarih, w=10, ad="Tarih"),
    dict(k="ay", h="Ay", f="=MONTH([[tar]]@)", fmt=f_tam, w=5, ad="Ay"),
]
TH_ILK = len(spec_th)
for t in TAHSILAT_TESTLERI:
    spec_th.append(dict(k=t["kod"], h=t["kod"], f=TH_TEST_F[t["kod"]], fmt=f_bayrak, w=5, ad=t["kod"],
                        aciklama=f"{t['ad']}\n{t['mantik']}"))
TH_SON = len(spec_th) - 1
spec_th += [
    dict(k="adet", h="Test Sayısı", f="=SUM([[H01]]@:[[H08]]@)", fmt=f_tam, w=6, ad="Adet"),
    dict(k="skor", h="Risk Skoru", f="=SUMPRODUCT([[H01]]@:[[H08]]@,W_Tahsilat)", fmt=f_skor, w=7, ad="Skor"),
    dict(k="sev", h="Risk Seviyesi", f='=IF([[skor]]@>=Yuksek_Esik,"Yüksek",IF([[skor]]@>=Orta_Esik,"Orta","Düşük"))', w=8, ad="Seviye"),
    dict(k="sira", h="Sıralama Anahtarı", f="=[[skor]]@+ROW()/10000000", fmt=f_ondalik, w=8, ad="SiraAnahtar", gizli=True),
]
H_TH = kolon_sayfasi(ws_th, spec_th, NH, "TH")
bayrak_bicim(ws_th, TH_ILK, TH_SON, NH)
seviye_bicim(ws_th, TH_SON + 3, NH)
ws_th.set_tab_color("#C00000")

# ================================================================== T_ORG (tedarikçi doğrulama & örüntü)
J9 = "{1,2,3,4,5,6,7,8,9}"
VKN_T = f"MOD(MID(Tedarikci!$C@,{J9},1)+10-{J9},10)"
VKN_F = (f'=IF(AND(LEN(Tedarikci!$C@)=10,ISNUMBER(--Tedarikci!$C@)),'
         f'--(MOD(10-MOD(SUMPRODUCT(MOD({VKN_T}*2^(10-{J9}),9)+9*({VKN_T}=9)),10),10)<>VALUE(RIGHT(Tedarikci!$C@,1))),1)')


def iban_f(hucre):
    s = f'MID({hucre},5,22)&"2927"&MID({hucre},3,2)'
    return (f'=IFERROR(IF(AND(LEN({hucre})=26,LEFT({hucre},2)="TR"),'
            f'--(MOD(VALUE(MOD(VALUE(MOD(VALUE(MOD(VALUE(LEFT({s},9)),97)&MID({s},10,7)),97)&MID({s},17,7)),97)&MID({s},24,5)),97)<>1),1),1)')


_EKLER = ["LTD. ŞTI.", "LTD. ŞTİ.", "SAN. VE TIC.", "SAN. VE TİC.", "A.Ş.", "LTD.", "ŞTI.", "ŞTİ."]
NORM_F = "UPPER(Tedarikci!$B@)"
for _e in _EKLER:
    NORM_F = f'SUBSTITUTE({NORM_F},"{_e}","")'
NORM_F = f'=TRIM(SUBSTITUTE(SUBSTITUTE({NORM_F},"."," "),","," "))'
spec_tg = [
    dict(k="id", h="Tedarikçi ID", f="=Tedarikci!A@", fmt=f_yesil, w=9, ad="ID"),
    dict(k="unvan", h="Ünvan", f="=Tedarikci!B@", fmt=f_yesil, w=30),
    dict(k="norm", h="Normalize Ünvan", f=NORM_F, w=26),
    dict(k="n", h="Ödeme Adedi", f="=COUNTIF(O_Ted,[[id]]@)", fmt=f_tam, w=7, ad="N"),
    dict(k="top", h="Toplam (TL)", f="=SUMIF(O_Ted,[[id]]@,O_TL)", fmt=f_tl0, w=13, ad="Toplam"),
    dict(k="ortln", h="Ort. ln", f="=IF([[n]]@>0,AVERAGEIF(TO_Ted,[[id]]@,TO_Ln),0)", fmt=f_ondalik, w=7, ad="OrtLn"),
    dict(k="sdln", h="SS ln", f="=IF([[n]]@>=3,SQRT(SUMPRODUCT((TO_Ted=[[id]]@)*(TO_Ln-[[ortln]]@)^2)/([[n]]@-1)),0)",
         fmt=f_ondalik, w=7, ad="SdLn"),
    dict(k="mx", h="En Büyük (TL)", f="=IF([[n]]@>0,AGGREGATE(14,6,O_TL/(O_Ted=[[id]]@),1),0)", fmt=f_tl0, w=12, ad="Max",
         aciklama="AGGREGATE(14;6;…) – LARGE, hata değerlerini (0'a bölme) yok sayar: koşullu maksimum."),
    dict(k="mx2", h="2. En Büyük (TL)", f="=IF([[n]]@>=2,AGGREGATE(14,6,O_TL/(O_Ted=[[id]]@),2),0)", fmt=f_tl0, w=12),
    dict(k="rsf", h="RSF", f="=IF([[mx2]]@>0,[[mx]]@/[[mx2]]@,0)", fmt=f_ondalik, w=6, ad="RSF"),
    dict(k="fmin", h="Min Fatura No", f="=IF([[n]]@>0,AGGREGATE(15,6,TO_FatNum/(TO_Ted=[[id]]@),1),0)", fmt=f_tam, w=9),
    dict(k="fmax", h="Maks Fatura No", f="=IF([[n]]@>0,AGGREGATE(14,6,TO_FatNum/(TO_Ted=[[id]]@),1),0)", fmt=f_tam, w=9),
    dict(k="ard", h="Ardışıklık Oranı", f="=IF([[n]]@>0,[[n]]@/([[fmax]]@-[[fmin]]@+1),0)", fmt=f_yuzde, w=8, ad="Ardisik"),
    dict(k="yuv", h="Yuvarlak Oranı", f="=IF([[n]]@>0,COUNTIFS(TO_Ted,[[id]]@,TO_T06,1)/[[n]]@,0)", fmt=f_yuzde, w=8),
    dict(k="bn", h="Benford n", f='=COUNTIFS(TO_Ted,[[id]]@,TO_Ilk,">0")', fmt=f_tam, w=7, ad="BenN"),
]
for d in range(1, 10):
    spec_tg.append(dict(k=f"b{d}", h=f"Hane {d} %", f=f"=IF([[bn]]@>0,COUNTIFS(TO_Ted,[[id]]@,TO_Ilk,{d})/[[bn]]@,0)",
                        fmt=f_yuzde, w=6))
spec_tg += [
    dict(k="bchi", h="Benford Ki-kare", f="=[[bn]]@*SUMPRODUCT(([[b1]]@:[[b9]]@-BenfordExpRow)^2/BenfordExpRow)", fmt=f_ondalik, w=8),
    dict(k="bp", h="Benford p", f="=IF([[bn]]@>=2,CHIDIST([[bchi]]@,8),1)", fmt=f_ondalik5, w=8),
    dict(k="bmad", h="Benford MAD", f="=SUMPRODUCT(ABS([[b1]]@:[[b9]]@-BenfordExpRow))/9", fmt=f_ondalik4, w=8),
    dict(k="G01", h="G01 VKN geçersiz", f=VKN_F, fmt=f_bayrak, w=6, ad="G01"),
    dict(k="G02", h="G02 Mükerrer VKN", f="=--(COUNTIF(Ted_VKN,Tedarikci!$C@)>1)", fmt=f_bayrak, w=6, ad="G02"),
    dict(k="G03", h="G03 IBAN geçersiz", f=iban_f("Tedarikci!$D@"), fmt=f_bayrak, w=6, ad="G03"),
    dict(k="G04", h="G04 Ortak IBAN", f="=--(COUNTIF(Ted_IBAN,Tedarikci!$D@)>1)", fmt=f_bayrak, w=6, ad="G04"),
    dict(k="G05", h="G05 IBAN=Personel", f="=--(COUNTIF(P_IBAN,Tedarikci!$D@)>0)", fmt=f_bayrak, w=6, ad="G05"),
    dict(k="G06", h="G06 Adres=Personel", f="=--(COUNTIF(P_Adres,Tedarikci!$E@)>0)", fmt=f_bayrak, w=6, ad="G06"),
    dict(k="G07", h="G07 Tel=Personel", f="=--(COUNTIF(P_Tel,Tedarikci!$F@)>0)", fmt=f_bayrak, w=6, ad="G07"),
    dict(k="G08", h="G08 Yaptırım", f="=--(OR(COUNTIF(Y_Norm,[[norm]]@)>0,SUMPRODUCT(ISNUMBER(SEARCH(Y_Anahtar,[[norm]]@))*(LEN(Y_Anahtar)>=8))>0))",
         fmt=f_bayrak, w=6, ad="G08"),
    dict(k="G09", h="G09 Açan=Onaylayan", f="=--(COUNTIFS(O_Ted,[[id]]@,O_Onaylayan,Tedarikci!$H@)>0)", fmt=f_bayrak, w=6, ad="G09"),
    dict(k="G10", h="G10 Ardışık Fatura", f="=--(AND([[n]]@>=Ardisik_MinN,[[ard]]@>=Ardisik_Esik))", fmt=f_bayrak, w=6, ad="G10"),
    dict(k="G11", h="G11 Benford", f="=--(AND([[bn]]@>=Benford_MinN,[[bp]]@<Benford_pEsik))", fmt=f_bayrak, w=6, ad="G11"),
    dict(k="G12", h="G12 Yuvarlak", f="=--(AND([[n]]@>=5,[[yuv]]@>Yuvarlak_Oran))", fmt=f_bayrak, w=6, ad="G12"),
    dict(k="skor", h="Org Bayrak Sayısı", f="=SUM([[G01]]@:[[G12]]@)", fmt=f_skor, w=7, ad="Skor"),
    dict(k="soz", h="Sözleşme Kullanımı", f="=IF(Tedarikci!$L@>0,[[top]]@/Tedarikci!$L@,0)", fmt=f_yuzde, w=8),
    dict(k="odrisk", h="Ödeme Risk Puanı Toplamı", f="=SUMIF(TO_Ted,[[id]]@,TO_Skor)", fmt=f_tl0, w=9, ad="OdemeRisk"),
    dict(k="sira", h="Sıralama Anahtarı", f="=[[skor]]@*1000000+[[odrisk]]@+ROW()/10000", fmt=f_ondalik, w=8, ad="SiraAnahtar", gizli=True),
]
H_TG = kolon_sayfasi(ws_tg, spec_tg, NT, "TOrg")
ad("TOrg_BenfordMatris", f"T_Org!${H_TG['b1']}$2:${H_TG['b9']}${NT + 1}")
gi = [j for j, s in enumerate(spec_tg) if s["k"].startswith("G")]
bayrak_bicim(ws_tg, gi[0], gi[-1], NT)
ws_tg.conditional_format(1, gi[-1] + 1, NT, gi[-1] + 1, {"type": "cell", "criteria": ">", "value": 0, "format": f_kirmizi})
ws_tg.conditional_format(1, gi[-1] + 2, NT, gi[-1] + 2, {"type": "cell", "criteria": ">", "value": 1, "format": f_kirmizi})
ws_tg.set_tab_color("#C00000")

# ================================================================== T_PERSONEL
TCKN = "Personel!$C@"
TCKN_F = (f'=IF(AND(LEN({TCKN})=11,ISNUMBER(--{TCKN}),LEFT({TCKN},1)<>"0"),'
          f'--OR(MOD(SUMPRODUCT(--MID({TCKN},{{1,3,5,7,9}},1))*7-SUMPRODUCT(--MID({TCKN},{{2,4,6,8}},1)),10)<>--MID({TCKN},10,1),'
          f'MOD(SUMPRODUCT(--MID({TCKN},{{1,2,3,4,5,6,7,8,9,10}},1)),10)<>--MID({TCKN},11,1)),1)')
spec_tp = [
    dict(k="id", h="Personel ID", f="=Personel!A@", fmt=f_yesil, w=10, ad="ID"),
    dict(k="ad", h="Ad Soyad", f="=Personel!B@", fmt=f_yesil, w=18),
    dict(k="P01", h="P01 TCKN geçersiz", f=TCKN_F, fmt=f_bayrak, w=8, ad="P01"),
    dict(k="P02", h="P02 Ortak IBAN", f="=--(COUNTIF(P_IBAN,Personel!$D@)>1)", fmt=f_bayrak, w=8, ad="P02"),
    dict(k="P03", h="P03 IBAN geçersiz", f=iban_f("Personel!$D@"), fmt=f_bayrak, w=8, ad="P03"),
    dict(k="P04", h="P04 Tedarikçi bağı",
         f="=--(OR(COUNTIF(Ted_IBAN,Personel!$D@)>0,COUNTIF(Ted_Adres,Personel!$E@)>0,COUNTIF(Ted_Tel,Personel!$F@)>0))",
         fmt=f_bayrak, w=8, ad="P04"),
    dict(k="skor", h="Bayrak Sayısı", f="=SUM([[P01]]@:[[P04]]@)", fmt=f_skor, w=8, ad="Skor"),
]
H_TP = kolon_sayfasi(ws_tp, spec_tp, NP, "TP")
bayrak_bicim(ws_tp, 2, 5, NP)
ws_tp.conditional_format(1, 6, NP, 6, {"type": "cell", "criteria": ">", "value": 0, "format": f_kirmizi})
ws_tp.set_tab_color("#C00000")

# ================================================================== BENFORD
B = ws_bf
B.set_column("A:A", 8); B.set_column("B:H", 11); B.set_column("I:I", 3); B.set_column("J:R", 8)
B.write("A1", "Benford Kanunu – Dijital Analiz (Nigrini)", f_baslik)
B.write("A2", "Kaynak: T_Odeme ilk hane / ilk iki hane / son iki hane yardımcı sütunları (TL ≥ 10). "
              "Z: süreklilik düzeltmeli oran testi; MAD eşikleri Nigrini (2012).", f_alt)
B.write_row("J3", list(range(1, 10)), f_h2)
B.write("I4", "P(d)", f_not)
for j in range(9):
    B.write_formula(3, 9 + j, f"=LOG10(1+1/{harf(9 + j)}3)", f_ondalik4)
ad("BenfordExpRow", "Benford!$J$4:$R$4")
B.write("J5", "↑ Beklenen ilk hane oranları (T_Org tedarikçi ki-kare testinde kullanılır)", f_not)
B.merge_range("A4:H4", "1) İlk hane testi", f_bolum)
B.write_row("A5", ["Hane", "Gözlenen", "Gözlenen %", "Beklenen %", "Fark", "Z", "Ki-kare katkı", "Anlamlı (Z>1,96)"], f_h)
for i, d in enumerate(range(1, 10), 6):
    B.write(f"A{i}", d, f_h2)
    B.write_formula(f"B{i}", f"=COUNTIF(TO_Ilk,A{i})", f_tam)
    B.write_formula(f"C{i}", f"=B{i}/$B$15", f_yuzde)
    B.write_formula(f"D{i}", f"=LOG10(1+1/A{i})", f_yuzde)
    B.write_formula(f"E{i}", f"=C{i}-D{i}", f_yuzde)
    B.write_formula(f"F{i}", f"=MAX(0,(ABS(C{i}-D{i})-1/(2*$B$15)))/SQRT(D{i}*(1-D{i})/$B$15)", f_ondalik)
    B.write_formula(f"G{i}", f"=$B$15*(C{i}-D{i})^2/D{i}", f_ondalik)
    B.write_formula(f"H{i}", f'=IF(F{i}>1.96,"Evet","")')
B.conditional_format("H6:H14", {"type": "cell", "criteria": "==", "value": '"Evet"', "format": f_kirmizi})
for r, (et, fo, fm) in enumerate([
        ("n", "=SUM(B6:B14)", f_tam), ("Ki-kare", "=SUM(G6:G14)", f_ondalik), ("sd", 8, f_tam),
        ("p-değeri", "=CHIDIST(B16,B17)", f_ondalik5), ("MAD", "=SUMPRODUCT(ABS(E6:E14))/9", f_ondalik5),
        ("Nigrini sonucu", '=IF(B19<=0.006,"Yakın uyum",IF(B19<=0.012,"Kabul edilebilir",IF(B19<=0.015,"Sınırda","UYUMSUZ")))', None)], 15):
    B.write(f"A{r}", et, f_h2)
    (B.write_formula if isinstance(fo, str) else B.write)(f"B{r}", fo, fm)
B.write("C20", "Eşikler: ≤0,006 yakın · ≤0,012 kabul · ≤0,015 sınırda · >0,015 uyumsuz", f_not)
ch = wb.add_chart({"type": "column"})
ch.add_series({"name": "Gözlenen", "categories": "=Benford!$A$6:$A$14", "values": "=Benford!$C$6:$C$14", "fill": {"color": "#2F5597"}})
ln = wb.add_chart({"type": "line"})
ln.add_series({"name": "Benford beklenen", "categories": "=Benford!$A$6:$A$14", "values": "=Benford!$D$6:$D$14",
               "line": {"color": "#C00000", "width": 2.25}, "marker": {"type": "circle", "size": 5}})
ch.combine(ln)
ch.set_title({"name": "İlk hane: gözlenen vs Benford"}); ch.set_y_axis({"num_format": "0%", "major_gridlines": {"visible": True}})
ch.set_legend({"position": "bottom"}); ch.set_size({"width": 560, "height": 300})
B.insert_chart("J7", ch)

B.merge_range("A23:G23", "2) İlk iki hane testi (10–99)", f_bolum)
B.write_row("A24", ["Hane", "Gözlenen", "Gözlenen %", "Beklenen %", "Fark", "Z", "Anlamlı"], f_h)
for i, d in enumerate(range(10, 100), 25):
    B.write(f"A{i}", d, f_h2)
    B.write_formula(f"B{i}", f"=COUNTIF(TO_IlkIki,A{i})", f_tam)
    B.write_formula(f"C{i}", f"=B{i}/SUM($B$25:$B$114)", f_yuzde2)
    B.write_formula(f"D{i}", f"=LOG10(1+1/A{i})", f_yuzde2)
    B.write_formula(f"E{i}", f"=C{i}-D{i}", f_yuzde2)
    B.write_formula(f"F{i}", f"=MAX(0,(ABS(C{i}-D{i})-1/(2*SUM($B$25:$B$114))))/SQRT(D{i}*(1-D{i})/SUM($B$25:$B$114))", f_ondalik)
    B.write_formula(f"G{i}", f'=IF(F{i}>1.96,"Evet","")')
B.conditional_format("G25:G114", {"type": "cell", "criteria": "==", "value": '"Evet"', "format": f_kirmizi})
B.write("J24", "MAD (ilk iki)", f_h2); B.write_formula("K24", "=SUMPRODUCT(ABS(E25:E114))/90", f_ondalik5)
B.write("J25", "Sonuç", f_h2)
B.write_formula("K25", '=IF(K24<=0.0012,"Yakın uyum",IF(K24<=0.0018,"Kabul edilebilir",IF(K24<=0.0022,"Sınırda","UYUMSUZ")))')
B.write("J26", "Eşikler: ≤0,0012 · ≤0,0018 · ≤0,0022 (Nigrini)", f_not)
ch2 = wb.add_chart({"type": "column"})
ch2.add_series({"name": "Gözlenen", "categories": "=Benford!$A$25:$A$114", "values": "=Benford!$C$25:$C$114",
                "fill": {"color": "#2F5597"}, "gap": 30})
ln2 = wb.add_chart({"type": "line"})
ln2.add_series({"name": "Beklenen", "categories": "=Benford!$A$25:$A$114", "values": "=Benford!$D$25:$D$114",
                "line": {"color": "#C00000", "width": 1.75}})
ch2.combine(ln2)
ch2.set_title({"name": "İlk iki hane – çıkıntılar uydurma/limit odaklı tutarları gösterir"})
ch2.set_y_axis({"num_format": "0.0%"}); ch2.set_x_axis({"interval_unit": 10}); ch2.set_legend({"position": "bottom"})
ch2.set_size({"width": 760, "height": 320})
B.insert_chart("J28", ch2)

B.merge_range("A118:F118", "3) Son iki hane testi (00–99, tekdüze beklenir)", f_bolum)
B.write_row("A119", ["Son iki", "Gözlenen", "Gözlenen %", "Beklenen %", "Ki-kare katkı", "Sapma kat"], f_h)
for i, d in enumerate(range(100), 120):
    B.write(f"A{i}", d, F(num_format="00", bold=True, bg_color=ACIK))
    B.write_formula(f"B{i}", f"=COUNTIF(TO_Son,A{i})", f_tam)
    B.write_formula(f"C{i}", f"=B{i}/SUM($B$120:$B$219)", f_yuzde2)
    B.write(f"D{i}", 0.01, f_yuzde2)
    B.write_formula(f"E{i}", f"=SUM($B$120:$B$219)*(C{i}-D{i})^2/D{i}", f_ondalik)
    B.write_formula(f"F{i}", f"=C{i}/D{i}", f_ondalik)
B.conditional_format("F120:F219", {"type": "cell", "criteria": ">=", "value": 2, "format": f_kirmizi})
B.write("J118", "Ki-kare", f_h2); B.write_formula("K118", "=SUM(E120:E219)", f_ondalik)
B.write("J119", "p-değeri (sd=99)", f_h2); B.write_formula("K119", "=CHIDIST(K118,99)", f_ondalik5)
B.write("J120", "'00' payı", f_h2); B.write_formula("K120", "=C120", f_yuzde2)
B.write("J121", "Yorum: '00' ve '50' çıkıntıları tahmini / uydurma / yuvarlak tutarlara işaret eder.", f_not)
B.freeze_panes(0, 1)

# ================================================================== İSTATİSTİK
S = ws_is
S.set_column("A:A", 26); S.set_column("B:H", 14); S.set_column("J:J", 22); S.set_column("K:Q", 12)
S.write("A1", "İstatistiksel Analiz, Tahmin ve Organizasyon Analitiği", f_baslik)
S.merge_range("A3:B3", "Tanımlayıcı istatistikler – Ödeme TL", f_bolum)
TL = f"T_Odeme!${H_TO['tl']}$2:${H_TO['tl']}${NO + 1}"
desc = [
    ("Adet – SUBTOTAL(2)", f"=SUBTOTAL(2,{TL})", f_tam),
    ("Toplam – SUBTOTAL(9)", f"=SUBTOTAL(9,{TL})", f_tl0),
    ("Filtreli toplam – SUBTOTAL(109)", f"=SUBTOTAL(109,{TL})", f_tl0),
    ("Ortalama", "=AVERAGE(O_TL)", f_tl0), ("Medyan", "=MEDIAN(O_TL)", f_tl0),
    ("Standart sapma", "=STDEV.S(O_TL)", f_tl0), ("Değişkenlik katsayısı", "=B9/B7", f_yuzde),
    ("Çarpıklık (SKEW)", "=SKEW(O_TL)", f_ondalik), ("Basıklık (KURT)", "=KURT(O_TL)", f_ondalik),
    ("Q1", "=QUARTILE.INC(O_TL,1)", f_tl0), ("Q3", "=QUARTILE.INC(O_TL,3)", f_tl0), ("IQR", "=B14-B13", f_tl0),
    ("Uç çit (Q3 + k×IQR)", "=B14+IQR_K*B15", f_tl0),
    ("P95 – AGGREGATE(16)", "=AGGREGATE(16,6,O_TL,0.95)", f_tl0), ("P99 – AGGREGATE(16)", "=AGGREGATE(16,6,O_TL,0.99)", f_tl0),
    ("En büyük – AGGREGATE(14)", "=AGGREGATE(14,6,O_TL,1)", f_tl0),
    ("Log-normal μ (ort. ln)", "=AVERAGE(TO_Ln)", f_ondalik), ("Log-normal σ", "=STDEV.S(TO_Ln)", f_ondalik),
]
for i, (et, fo, fm) in enumerate(desc, 4):
    S.write(f"A{i}", et)
    S.write_formula(f"B{i}", fx(fo), fm)
ad("IQR_Ust", "Istatistik!$B$16")
S.write("A23", "SUBTOTAL(109) T_Odeme'de filtre uygulandığında yalnız görünen satırları toplar.", f_not)

# FREQUENCY histogram
S.merge_range("D3:F3", "Tutar dağılımı – FREQUENCY", f_bolum)
S.write_row("D4", ["Üst sınır (TL)", "Adet", "Pay"], f_h)
kutu = [1_000, 2_500, 5_000, 10_000, 25_000, 50_000, 100_000, 250_000, 500_000, 1_000_000]
for i, k in enumerate(kutu, 5):
    S.write(f"D{i}", k, f_girdi_tl)
S.write(f"D{5 + len(kutu)}", "üstü", f_h2)
ad("Ist_Kutu", f"Istatistik!$D$5:$D${4 + len(kutu)}")
S.write_array_formula(f"E5:E{5 + len(kutu)}", "{=FREQUENCY(O_TL,Ist_Kutu)}", f_tam)
for i in range(5, 6 + len(kutu)):
    S.write_formula(f"F{i}", f"=E{i}/SUM($E$5:$E${5 + len(kutu)})", f_yuzde)
chh = wb.add_chart({"type": "column"})
chh.add_series({"name": "Adet", "categories": f"=Istatistik!$D$5:$D${5 + len(kutu)}",
                "values": f"=Istatistik!$E$5:$E${5 + len(kutu)}", "fill": {"color": "#2F5597"}, "gap": 40,
                "data_labels": {"value": True}})
chh.set_title({"name": "Ödeme tutarı histogramı (log ölçekli kutular)"}); chh.set_legend({"none": True})
chh.set_x_axis({"num_format": "#,##0"}); chh.set_size({"width": 520, "height": 280})
S.insert_chart("H3", chh)

# Aylık analiz + FORECAST.LINEAR + CORREL
S.merge_range("A26:H26", "Aylık analiz – FORECAST.LINEAR ile beklenen değer ve sapma (S04)", f_bolum)
S.write_row("A27", ["Ay", "Ödeme (TL)", "Ödeme adedi", "Tahsilat (TL)", "Tahmin (önceki aylar)", "Sapma", "Bayrak",
                    "Riskli ödeme TL"], f_h)
for i, m in enumerate(range(1, 13), 28):
    S.write(f"A{i}", m, f_h2)
    S.write_formula(f"B{i}", f"=SUMIF(TO_Ay,A{i},O_TL)", f_tl0)
    S.write_formula(f"C{i}", f"=COUNTIF(TO_Ay,A{i})", f_tam)
    S.write_formula(f"D{i}", f"=SUMIF(TH_Ay,A{i},H_Tutar)", f_tl0)
    if m >= 4:
        S.write_formula(f"E{i}", fx(f"=FORECAST.LINEAR(A{i},$B$28:B{i - 1},$A$28:A{i - 1})"), f_tl0)
        S.write_formula(f"F{i}", f"=IF(E{i}>0,B{i}/E{i}-1,0)", f_yuzde)
        S.write_formula(f"G{i}", f'=IF(ABS(F{i})>Tahmin_Sapma,"SAPMA","")')
    else:
        S.write(f"E{i}", "–", f_not)
    S.write_formula(f"H{i}", f'=SUMIFS(O_TL,TO_Ay,A{i},TO_Skor,">="&Yuksek_Esik)', f_tl0)
S.conditional_format("G31:G39", {"type": "cell", "criteria": "==", "value": '"SAPMA"', "format": f_kirmizi})
ad("Ist_Ay", "Istatistik!$A$28:$A$39")
ad("Ist_AylikOdeme", "Istatistik!$B$28:$B$39")
ad("Ist_AylikTahsilat", "Istatistik!$D$28:$D$39")
S.write("A41", "Gelecek ay tahmini (13)", F(bold=True)); S.write_formula("B41", fx("=FORECAST.LINEAR(13,Ist_AylikOdeme,Ist_Ay)"), f_tl0)
S.write("A42", "CORREL(ödeme; tahsilat)", F(bold=True)); S.write_formula("B42", "=CORREL(Ist_AylikOdeme,Ist_AylikTahsilat)", f_ondalik)
S.write("A43", "CORREL(ted. adet; toplam)", F(bold=True)); S.write_formula("B43", "=CORREL(TOrg_N,TOrg_Toplam)", f_ondalik)
S.write("A44", "CORREL(TL; ln TL)", F(bold=True)); S.write_formula("B44", "=CORREL(O_TL,TO_Ln)", f_ondalik)
S.write("C42", "Ödeme ve tahsilat genelde birlikte hareket eder; ani kopuş (düşük/negatif korelasyon) nakit "
               "döngüsünde manipülasyonu veya dönem sonu kaydırmasını işaret edebilir.", f_not)
chm = wb.add_chart({"type": "column"})
chm.add_series({"name": "Gerçekleşen ödeme", "categories": "=Istatistik!$A$28:$A$39", "values": "=Istatistik!$B$28:$B$39",
                "fill": {"color": "#2F5597"}})
lnm = wb.add_chart({"type": "line"})
lnm.add_series({"name": "FORECAST.LINEAR", "categories": "=Istatistik!$A$28:$A$39", "values": "=Istatistik!$E$28:$E$39",
                "line": {"color": "#C00000", "dash_type": "dash", "width": 2}, "marker": {"type": "diamond"}})
lnm.add_series({"name": "Tahsilat", "categories": "=Istatistik!$A$28:$A$39", "values": "=Istatistik!$D$28:$D$39",
                "line": {"color": "#548235", "width": 2}})
chm.combine(lnm)
chm.set_title({"name": "Aylık ödeme: gerçekleşen vs tahmin"}); chm.set_y_axis({"num_format": "#,##0,,\" mn\""})
chm.set_legend({"position": "bottom"}); chm.set_size({"width": 620, "height": 300})
S.insert_chart("J26", chm)

# Haftanın günü & saat dağılımı
S.merge_range("A47:C47", "Gizli zaman örüntüsü – haftanın günü", f_bolum)
S.write_row("A48", ["Gün", "Adet", "TL"], f_h)
for i, (g, adg) in enumerate(zip(range(1, 8), ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz"]), 49):
    S.write(f"A{i}", f"{g} – {adg}", f_h2)
    S.write_formula(f"B{i}", f"=COUNTIF(TO_HaftaGunu,{g})", f_tam)
    S.write_formula(f"C{i}", f"=SUMIF(TO_HaftaGunu,{g},O_TL)", f_tl0)
S.merge_range("E47:G47", "Saat dağılımı", f_bolum)
S.write_row("E48", ["Saat", "Adet", "Mesai dışı?"], f_h)
for i, h in enumerate(range(24), 49):
    S.write(f"E{i}", h, f_h2)
    S.write_formula(f"F{i}", f'=COUNTIFS(O_Saat,">="&E{i}/24,O_Saat,"<"&(E{i}+1)/24)', f_tam)
    S.write_formula(f"G{i}", f'=IF(OR(E{i}/24<Mesai_Bas,E{i}/24>=Mesai_Bit),"Evet","")')
S.conditional_format("F49:F72", {"type": "data_bar", "bar_color": "#2F5597", "bar_solid": True})
S.conditional_format("G49:G72", {"type": "cell", "criteria": "==", "value": '"Evet"', "format": f_turuncu})

# Kullanıcı / onay organizasyon analizi
S.merge_range("I47:Q47", "Organizasyon – onaylayan kullanıcı analizi (yetki, görevler ayrılığı, limit altı yoğunlaşma)", f_bolum)
S.write_row("I48", ["Kullanıcı", "Rol", "Onay limiti", "Onay adedi", "Onay TL", "SoD ihlali (T09)", "Yetki aşımı (T10)",
                    "Limit altı (T07)", "Limit altı payı"], f_h)
for i, u in enumerate(KU, 49):
    S.write(f"I{i}", u["KullaniciID"], f_h2)
    S.write_formula(f"J{i}", f"=INDEX(Kullanici!$C$2:$C${NK + 1},MATCH(I{i},K_ID,0))")
    S.write_formula(f"K{i}", f"=INDEX(K_Limit,MATCH(I{i},K_ID,0))", f_tl0)
    S.write_formula(f"L{i}", f"=COUNTIF(O_Onaylayan,I{i})", f_tam)
    S.write_formula(f"M{i}", f"=SUMIF(O_Onaylayan,I{i},O_TL)", f_tl0)
    S.write_formula(f"N{i}", f"=SUMPRODUCT((O_Onaylayan=I{i})*TO_T09)", f_tam)
    S.write_formula(f"O{i}", f"=SUMPRODUCT((O_Onaylayan=I{i})*TO_T10)", f_tam)
    S.write_formula(f"P{i}", f"=SUMPRODUCT((O_Onaylayan=I{i})*TO_T07)", f_tam)
    S.write_formula(f"Q{i}", f"=IF(L{i}>0,P{i}/L{i},0)", f_yuzde)
S.conditional_format(f"N49:P{48 + NK}", {"type": "cell", "criteria": ">", "value": 0, "format": f_kirmizi})
S.freeze_panes(1, 0)

# ================================================================== DİNAMİK_365
Dn = ws_dn
Dn.set_column("A:A", 4); Dn.set_column("B:B", 16); Dn.set_column("C:C", 34); Dn.set_column("D:D", 58)
Dn.set_column("E:E", 2); Dn.set_column("F:F", 12); Dn.set_column("G:G", 30); Dn.set_column("H:T", 12)
Dn.write("A1", "Excel 365 Dinamik Dizi Fonksiyonları ile Denetim Analitiği", f_baslik)
Dn.write("A2", "Bu sayfa Microsoft 365 / Excel 2024 gerektirir (GROUPBY ve PIVOTBY yalnızca Microsoft 365). "
               "Sonuçlar F sütunundan itibaren 'taşar' (spill). Ad Yöneticisi'nde tanımlı LAMBDA fonksiyonları: "
               "VKN_GECERLI, TCKN_GECERLI, IBAN_GECERLI, ZSKOR, BENFORD_P.", f_alt)
Dn.write("B3", "Seçili tedarikçi:", F(bold=True))
Dn.write("C3", "V061", f_girdi)
Dn.data_validation("C3", {"validate": "list", "source": "=Ted_ID"})
ad("Secili_Ted", "Dinamik_365!$C$3")
Dn.write("D3", "← Listeden tedarikçi seçin (AGGREGATE, XLOOKUP, SCAN blokları bu hücreyi kullanır)", f_not)
Dn.write_row("A5", ["#", "Fonksiyon", "Denetim amacı", "Formül"], f_h)
Dn.write("F5", "Canlı sonuç (taşan dizi) →", f_h)

wb.define_name("BENFORD_P", fx("=LAMBDA(§d,LOG10(1+1/§d))"))
wb.define_name("ZSKOR", fx("=LAMBDA(§x,§aralik,(§x-AVERAGE(§aralik))/STDEV.S(§aralik))"))
wb.define_name("VKN_GECERLI", fx(
    '=LAMBDA(§v,IFERROR(LET(§s,TEXT(§v,"0000000000"),§j,SEQUENCE(9),§t,MOD(--MID(§s,§j,1)+10-§j,10),'
    '§k,MOD(§t*2^(10-§j),9)+9*(§t=9),AND(LEN(§s)=10,MOD(10-MOD(SUM(§k),10),10)=--RIGHT(§s,1))),FALSE))'))
wb.define_name("TCKN_GECERLI", fx(
    '=LAMBDA(§n,IFERROR(LET(§d,--MID(§n,SEQUENCE(11),1),§o,SUM(CHOOSEROWS(§d,1,3,5,7,9)),§e,SUM(CHOOSEROWS(§d,2,4,6,8)),'
    'AND(LEN(§n)=11,INDEX(§d,1)<>0,MOD(§o*7-§e,10)=INDEX(§d,10),MOD(SUM(TAKE(§d,10)),10)=INDEX(§d,11))),FALSE))'))
wb.define_name("IBAN_GECERLI", fx(
    '=LAMBDA(§i,IFERROR(LET(§s,MID(§i,5,22)&"2927"&MID(§i,3,2),'
    '§m,REDUCE(0,SEQUENCE(LEN(§s)),LAMBDA(§a,§k,MOD(§a*10+MID(§s,§k,1),97))),'
    'AND(LEN(§i)=26,LEFT(§i,2)="TR",§m=1)),FALSE))'))

SIRALI = "SORTBY(HSTACK(O_ID,O_Ted,O_TL,TO_Skor),TO_Skor,-1)"
BLOKLAR = [
    ("TAKE", "En riskli 15 ödeme (kural tabanlı skor sırasıyla)", f"=TAKE({SIRALI},15)",
     ["Ödeme ID", "Tedarikçi", "TL", "Skor"], 16),
    ("DROP", "İkinci öncelik dilimi: 16.–30. sıradaki ödemeler (ilk 15 düşülür)", f"=DROP(TAKE({SIRALI},30),15)",
     ["Ödeme ID", "Tedarikçi", "TL", "Skor"], 16),
    ("VSTACK + MAP + FILTER", "Tedarikçi ve personel IBAN'larını alt alta birleştir; birden fazla varlıkta geçen IBAN'ları listele",
     '=LET(§v,VSTACK(HSTACK(Ted_ID,Ted_IBAN),HSTACK(P_ID,P_IBAN)),§i,CHOOSECOLS(§v,2),'
     '§adet,MAP(§i,LAMBDA(§x,SUM(--(§i=§x)))),FILTER(HSTACK(§v,§adet),§adet>1,"Ortak IBAN yok"))',
     ["Varlık", "IBAN", "Geçiş"], 10),
    ("HSTACK + FILTER", "Organizasyon testinde bayrak alan tedarikçilerin kartı",
     '=FILTER(HSTACK(Ted_ID,Ted_Unvan,TOrg_N,TOrg_Toplam,TOrg_Skor),TOrg_Skor>0,"Bayraklı tedarikçi yok")',
     ["ID", "Ünvan", "Adet", "Toplam TL", "Bayrak"], 14),
    ("TOCOL + UNIQUE", "Telefon ve adresleri tek sütuna indir; tedarikçi/personel arasında tekrarlananları bul",
     '=LET(§t,TOCOL(HSTACK(VSTACK(Ted_Tel,P_Tel),VSTACK(Ted_Adres,P_Adres)),3),§u,UNIQUE(§t),'
     '§adet,MAP(§u,LAMBDA(§x,SUM(--(§t=§x)))),FILTER(HSTACK(§u,§adet),§adet>1,"Tekrar yok"))',
     ["Telefon / Adres", "Geçiş"], 8),
    ("TOROW", "Bayraklı tedarikçi kimlikleri tek satırda (rapor başlığı / e-posta için)",
     '=TOROW(FILTER(Ted_ID,TOrg_Skor>0,"Yok"))', None, 3),
    ("SORTBY", "Tedarikçileri önce organizasyon bayrağı, sonra ödeme risk puanına göre sırala (ilk 20)",
     "=TAKE(SORTBY(HSTACK(Ted_ID,Ted_Unvan,TOrg_Toplam,TOrg_Skor,TOrg_OdemeRisk),TOrg_Skor,-1,TOrg_OdemeRisk,-1),20)",
     ["ID", "Ünvan", "Toplam TL", "Org bayrak", "Ödeme risk"], 21),
    ("GROUPBY", "Onaylayan bazında: toplam TL, SoD ihlali, yetki aşımı, limit altı adedi",
     "=GROUPBY(O_Onaylayan,HSTACK(O_TL,TO_T09,TO_T10,TO_T07),LAMBDA(§x,SUM(§x)),0,1)",
     ["Onaylayan", "TL", "SoD", "Yetki aşımı", "Limit altı"], 13),
    ("PIVOTBY", "Onaylayan × ay risk puanı ısı haritası (dönem sonu / kişi yoğunlaşması)",
     "=PIVOTBY(O_Onaylayan,TO_Ay,TO_Skor,LAMBDA(§x,SUM(§x)))", None, 13),
    ("SUBTOTAL", "Toplam ve filtreye duyarlı toplam (T_Odeme'de filtre uygulayın)",
     f"=HSTACK(SUBTOTAL(9,O_TL),SUBTOTAL(109,T_Odeme!${H_TO['tl']}$2:${H_TO['tl']}${NO + 1}))", ["SUBTOTAL(9)", "SUBTOTAL(109)"], 3),
    ("AGGREGATE", "Seçili tedarikçinin en büyük, 2. en büyük ödemesi ve RSF (hata yok sayan koşullu LARGE)",
     "=LET(§a,AGGREGATE(14,6,O_TL/(O_Ted=Secili_Ted),1),§b,AGGREGATE(14,6,O_TL/(O_Ted=Secili_Ted),2),HSTACK(§a,§b,§a/§b))",
     ["En büyük", "2. en büyük", "RSF"], 3),
    ("FREQUENCY", "Tutar kutularına göre dağılım (Istatistik!D5:D14 kutuları)", "=FREQUENCY(O_TL,Ist_Kutu)", ["Adet"], 13),
    ("FORECAST.LINEAR + MAP", "Her ay için yalnız önceki aylarla kurulan doğrusal tahmin (4.–12. ay)",
     "=HSTACK(SEQUENCE(9,1,4),MAP(SEQUENCE(9,1,4),LAMBDA(§m,FORECAST.LINEAR(§m,TAKE(Ist_AylikOdeme,§m-1),SEQUENCE(§m-1)))))",
     ["Ay", "Tahmin TL"], 10),
    ("CORREL", "Aylık ödeme–tahsilat korelasyonu ve tedarikçi işlem adedi–tutar korelasyonu",
     "=HSTACK(CORREL(Ist_AylikOdeme,Ist_AylikTahsilat),CORREL(TOrg_N,TOrg_Toplam))", ["Öd.–Tah.", "Adet–Tutar"], 3),
    ("LAMBDA", "Ad Yöneticisi'ndeki özel fonksiyonlar: Benford olasılıkları, seçili tedarikçinin VKN/IBAN'ı, P045 TCKN'si, en büyük ödemenin Z'si",
     "=HSTACK(MAP(SEQUENCE(1,9),LAMBDA(§d,BENFORD_P(§d))),VKN_GECERLI(XLOOKUP(Secili_Ted,Ted_ID,Ted_VKN)),"
     "IBAN_GECERLI(XLOOKUP(Secili_Ted,Ted_ID,Ted_IBAN)),TCKN_GECERLI(XLOOKUP(\"P045\",P_ID,P_TCKN)),ZSKOR(MAX(O_TL),O_TL))",
     ["P(1)", "P(2)", "P(3)", "P(4)", "P(5)", "P(6)", "P(7)", "P(8)", "P(9)", "VKN ok?", "IBAN ok?", "TCKN P045", "Z(maks)"], 3),
    ("MAP", "Tüm tedarikçilerde VKN ve IBAN doğrulaması – yalnız geçersizleri listele",
     '=LET(§ok,MAP(Ted_VKN,LAMBDA(§v,VKN_GECERLI(§v))),§ib,MAP(Ted_IBAN,LAMBDA(§v,IBAN_GECERLI(§v))),'
     'FILTER(HSTACK(Ted_ID,Ted_VKN,Ted_IBAN,§ok,§ib),NOT(§ok*§ib),"Hepsi geçerli"))',
     ["ID", "VKN", "IBAN", "VKN ok", "IBAN ok"], 6),
    ("BYROW", "Tedarikçi Benford MAD'i satır satır (n ≥ Benford_MinN) – en uyumsuz 10 tedarikçi; altta skor mutabakatı",
     "=LET(§mad,BYROW(TOrg_BenfordMatris,LAMBDA(§satir,SUM(ABS(§satir-BenfordExpRow))/9)),"
     "§f,FILTER(HSTACK(Ted_ID,TOrg_BenN,§mad),TOrg_BenN>=Benford_MinN),TAKE(SORTBY(§f,CHOOSECOLS(§f,3),-1),10))",
     ["ID", "n", "MAD"], 11),
    ("BYROW (mutabakat)", "BYROW ile yeniden hesaplanan ödeme skorları T_Odeme skorundan farklı olan satır sayısı (0 olmalı)",
     "=SUM(--(BYROW(TO_Bayraklar,LAMBDA(§satir,SUM(§satir*W_Odeme)))<>TO_Skor))", ["Fark adedi"], 3),
    ("SCAN", "Seçili tedarikçiye kümülatif ödeme ve sözleşme limiti aşım anı",
     '=LET(§f,O_Ted=Secili_Ted,§t,FILTER(O_Tarih,§f),§a,FILTER(O_TL,§f),§k,SCAN(0,§a,LAMBDA(§s,§x,§s+§x)),'
     '§lim,XLOOKUP(Secili_Ted,Ted_ID,Ted_Sozlesme),HSTACK(§t,§a,§k,IF(§k>§lim,"AŞIM","")))',
     ["Tarih", "TL", "Kümülatif", "Limit"], 45),
]
satir = 6
f_dn_tl = F(num_format="#,##0")
for no, (fn, amac, formul, basliklar, yukseklik) in enumerate(BLOKLAR, 1):
    Dn.write(satir, 0, no, f_h2)
    Dn.write(satir, 1, fn, F(bold=True, font_color=LACI, valign="top", text_wrap=True))
    Dn.write(satir, 2, amac, f_metin)
    Dn.write_string(satir, 3, formul.replace("§", ""), f_kod)
    if basliklar:
        Dn.write_row(satir, 5, basliklar, f_h2)
        hedef = satir + 1
    else:
        hedef = satir
    Dn.write_dynamic_array_formula(hedef, 5, hedef, 5, fx(formul), f_dn_tl if fn not in ("CORREL",) else f_ondalik)
    Dn.set_row(satir, 48)
    satir += yukseklik + 2
Dn.freeze_panes(5, 0)

# ================================================================== PYTHON BULGULARI
def csv_df(ad_):
    p = CIKTI / f"{ad_}.csv"
    return pd.read_csv(p) if p.exists() else pd.DataFrame()


iso = csv_df("py_isolation")
Q = ws_pif
Q.write_row(0, 0, list(iso.columns), f_h)
for i, r in enumerate(iso.itertuples(index=False), 1):
    for j, v in enumerate(r):
        if isinstance(v, (bool, np.bool_)):
            Q.write_boolean(i, j, bool(v))
        elif isinstance(v, numbers.Number) and not pd.isna(v):
            Q.write_number(i, j, float(v))
        else:
            Q.write_string(i, j, str(v))
Q.set_column(0, len(iso.columns), 13)
Q.freeze_panes(1, 1)
Q.autofilter(0, 0, len(iso), len(iso.columns) - 1)
ci = {c: harf(j) for j, c in enumerate(iso.columns)}
ad("PY_IF_ID", f"Py_IsolationForest!${ci['OdemeID']}$2:${ci['OdemeID']}${len(iso) + 1}")
ad("PY_IF_Skor", f"Py_IsolationForest!${ci['IF_skor']}$2:${ci['IF_skor']}${len(iso) + 1}")
ad("PY_IF_Ayk", f"Py_IsolationForest!${ci['IF_aykiri']}$2:${ci['IF_aykiri']}${len(iso) + 1}")
Q.write_comment(0, 0, "src/analiz.py çıktısı (değer). Isolation Forest, 9 özellik: ln tutar, saat, hafta günü, "
                      "kayıttan geçen gün, tedarikçi içi Z, yuvarlaklık, aynı tutar tekrarı, fatura→ödeme gün farkı, onay limit oranı.",
                {"x_scale": 3, "y_scale": 2})

Y = ws_py
Y.set_column("A:A", 30); Y.set_column("B:J", 16)
Y.write("A1", "Python İleri Analitik Bulguları (src/analiz.py)", f_baslik)
Y.write("A2", "Bu sayfadaki değerler Python ile hesaplanıp aktarılmıştır (formül değildir). Veriyi değiştirirseniz "
              "python src/calistir.py ile yeniden üretin.", f_alt)


def tablo_yaz(baslik, df, r, aciklama=""):
    Y.merge_range(r, 0, r, max(len(df.columns) - 1, 3), baslik, f_bolum)
    if aciklama:
        Y.write(r + 1, 0, aciklama, f_not)
        r += 1
    Y.write_row(r + 1, 0, list(df.columns), f_h)
    for i, row in enumerate(df.itertuples(index=False), r + 2):
        for j, v in enumerate(row):
            if isinstance(v, (bool, np.bool_)):
                Y.write_string(i, j, "Evet" if v else "Hayır")
            elif isinstance(v, numbers.Number) and not pd.isna(v):
                v = float(v)
                fm = f_bilimsel if 0 < abs(v) < 0.001 else (f_ondalik if abs(v) < 10 and v != int(v) else f_tl0)
                Y.write_number(i, j, v, fm)
            else:
                Y.write_string(i, j, "" if pd.isna(v) else str(v))
    if df.empty:
        Y.write(r + 2, 0, "Bulgu yok", f_not)
    return r + 3 + max(len(df), 1)


r = 3
r = tablo_yaz("Benford ailesi – ki-kare, MAD, Kolmogorov-Smirnov", csv_df("py_benford_ozet"), r,
              "KS > KS_kritik_05 ise %5 düzeyinde uyumsuzluk. Son iki hane için tekdüze dağılım beklenir.")
r = tablo_yaz("Benford toplam (summation) testi – en yüksek 10 ilk-iki-hane grubu", csv_df("py_benford_toplam"), r,
              "Kat > 2: o grupta olağandışı büyük tutar yoğunlaşması (Nigrini Ch.7).")
iso_top = iso.head(25)[["Sira", "OdemeID", "TedarikciID", "OdemeTarihi", "Saat", "TutarTL", "Onaylayan", "IF_skor",
                        "Robust_uzaklik", "Etken_ozellikler"]] if not iso.empty else iso
r = tablo_yaz("Isolation Forest – en anormal 25 ödeme", iso_top, r,
              "Kural tabanlı testlerin göremediği çok değişkenli aykırılıklar. Etken özellikler: ölçeklenmiş en büyük 3 sapma.")
r = tablo_yaz("Varlık ağı – ortak IBAN / adres / telefon / VKN kümeleri (union-find)", csv_df("py_ag"), r)
r = tablo_yaz("Bulanık ünvan eşleşmesi (tedarikçi–tedarikçi, tedarikçi–yaptırım)", csv_df("py_bulanik"), r)
r = tablo_yaz("Onaylayan kullanıcı davranışı (saat dağılımı ki-kare, SoD, limit altı yoğunlaşma)", csv_df("py_kullanici"), r)
r = tablo_yaz("Lapping zinciri (ödeyen → fatura sahibi yönlü graf)", csv_df("py_lapping"), r)
r = tablo_yaz("Makbuz numarası boşlukları", csv_df("py_bosluk"), r)

# ================================================================== TEST KÜTÜPHANESİ
L = ws_lib
L.write("A1", "Test Kütüphanesi – ACFE Fraud Tree, Nigrini, IIA GTAG, ISA 240 ve Big Four FDA pratiğiyle eşlenmiş", f_baslik)
L.write("A2", "Topluluk sinyali: Topluluk sayfasında (Reddit + Big Four/ACFE kazıma) testin anahtar kelimeleriyle eşleşen "
              "metin adedi ve Reddit puan toplamı – test önceliklendirmesi için.", f_alt)
kol = ["Kod", "Test", "Kategori", "ACFE / şema", "Mantık", "Excel fonksiyonları", "Ağırlık", "İsabet", "Referans",
       "Topluluk eşleşme", "Topluluk puanı", "Anahtar kelimeler"]
L.write_row(3, 0, kol, f_h)
for j, w in enumerate([6, 34, 12, 30, 48, 26, 7, 8, 26, 9, 9, 40]):
    L.set_column(j, j, w)
ISABET = {}
for t in ODEME_TESTLERI:
    ISABET[t["kod"]] = f"=SUM(TO_{t['kod']})"
for t in TAHSILAT_TESTLERI:
    ISABET[t["kod"]] = f"=SUM(TH_{t['kod']})"
for t in ORGANIZASYON_TESTLERI:
    ISABET[t["kod"]] = f"=SUM(TOrg_{t['kod']})"
for t in PERSONEL_TESTLERI:
    ISABET[t["kod"]] = f"=SUM(TP_{t['kod']})"
ISABET.update({"S01": "=Benford!B20", "S02": "=Benford!K25", "S03": "=Benford!K119",
               "S04": '=COUNTIF(Istatistik!G28:G39,"SAPMA")', "S05": "=Istatistik!B42"})
NTOP = max(len(TOP), 1)
for i, t in enumerate(TUM_TESTLER, 5):
    L.write(f"A{i}", t["kod"], f_h2)
    L.write_string(f"B{i}", t["ad"], f_metin)
    L.write(f"C{i}", t["kategori"])
    L.write_string(f"D{i}", t["acfe"], f_metin)
    L.write_string(f"E{i}", t["mantik"], f_metin)
    L.write_string(f"F{i}", t["fonksiyonlar"], f_metin)
    if t["kod"] in AG_SATIR:
        L.write_formula(f"G{i}", f"=Parametreler!B{AG_SATIR[t['kod']]}", f_yesil)
    else:
        L.write(f"G{i}", "–", f_not)
    if t["kod"] in ISABET:
        L.write_formula(f"H{i}", ISABET[t["kod"]], f_ondalik if t["kod"] == "S05" else None)
    elif t["kod"] == "S06":
        L.write_formula(f"H{i}", "=COUNTIF(PY_IF_Ayk,TRUE)")
    else:
        L.write(f"H{i}", "Python", f_not)
    L.write(f"I{i}", t["referans"], f_metin)
    L.write_formula(f"J{i}", f'=COUNTIF(Topluluk!$F$2:$F${NTOP + 1},"*{t["kod"]}*")', f_tam)
    L.write_formula(f"K{i}", f'=SUMIF(Topluluk!$F$2:$F${NTOP + 1},"*{t["kod"]}*",Topluluk!$C$2:$C${NTOP + 1})', f_tam)
    L.write(f"L{i}", ", ".join(t["anahtar"]), f_not)
L.freeze_panes(4, 2)
L.autofilter(3, 0, 4 + len(TUM_TESTLER), len(kol) - 1)
L.set_tab_color("#548235")

# ================================================================== RİSK PANOSU
R = ws_pan
R.set_column("A:A", 2); R.set_column("B:B", 6); R.set_column("C:C", 34); R.set_column("D:L", 13)
R.write("B1", "Suistimal Denetimi – Risk Panosu", f_baslik)
R.write("B2", "Tüm göstergeler canlı formüldür; Parametreler sayfasındaki eşik/ağırlık değişince güncellenir.", f_alt)
kpis = [
    ("Ödeme adedi", "=COUNTA(O_ID)", f_kpi), ("Ödeme toplamı (TL)", "=SUM(O_TL)", f_kpi_tl),
    ("Yüksek riskli ödeme", '=COUNTIF(TO_Seviye,"Yüksek")', f_kpi),
    ("Yüksek riskli TL", '=SUMIF(TO_Seviye,"Yüksek",O_TL)', f_kpi_tl),
    ("Orta riskli ödeme", '=COUNTIF(TO_Seviye,"Orta")', f_kpi),
    ("Bayraklı tahsilat", '=COUNTIF(TH_Skor,">0")', f_kpi),
    ("Bayraklı tedarikçi", '=COUNTIF(TOrg_Skor,">0")', f_kpi),
    ("Bayraklı personel", '=COUNTIF(TP_Skor,">0")', f_kpi),
    ("Benford (ilk hane)", "=Benford!B20", f_kpi_txt),
    ("Python IF aykırı", "=COUNTIF(PY_IF_Ayk,TRUE)", f_kpi),
]
for j, (et, fo, fm) in enumerate(kpis):
    rr, cc = (3, 3 + j) if j < 5 else (6, 3 + j - 5)
    R.write(rr, cc, et, f_kpi_ad)
    R.write_formula(rr + 1, cc, fo, fm)
    R.set_row(rr, 30); R.set_row(rr + 1, 30)
R.write("C4", "Genel görünüm", F(bold=True, font_color=LACI))

R.merge_range("B10:H10", "Test isabetleri – ödeme (T) ve tahsilat (H)", f_bolum)
R.write_row("B11", ["Kod", "Test", "Ağırlık", "İsabet", "Etkilenen TL", "İsabet %", "Kategori"], f_h)
rr = 12
for t in ODEME_TESTLERI + TAHSILAT_TESTLERI:
    on = "TO" if t["kod"].startswith("T") else "TH"
    tl = "O_TL" if on == "TO" else "H_Tutar"
    nn = NO if on == "TO" else NH
    R.write(f"B{rr}", t["kod"], f_h2)
    R.write(f"C{rr}", t["ad"])
    R.write_formula(f"D{rr}", f"=Parametreler!B{AG_SATIR[t['kod']]}", f_yesil)
    R.write_formula(f"E{rr}", f"=SUM({on}_{t['kod']})", f_tam)
    R.write_formula(f"F{rr}", f"=SUMPRODUCT({on}_{t['kod']},{tl})", f_tl0)
    R.write_formula(f"G{rr}", f"=E{rr}/{nn}", f_yuzde)
    R.write(f"H{rr}", t["kategori"])
    rr += 1
son_test = rr - 1
R.conditional_format(f"E12:E{son_test}", {"type": "data_bar", "bar_color": "#C00000", "bar_solid": True})
ch = wb.add_chart({"type": "bar"})
ch.add_series({"name": "İsabet", "categories": f"=Risk_Panosu!$B$12:$B${son_test}", "values": f"=Risk_Panosu!$E$12:$E${son_test}",
               "fill": {"color": "#C00000"}, "gap": 40})
ch.set_title({"name": "Test bazında isabet adedi"}); ch.set_legend({"none": True})
ch.set_y_axis({"reverse": True}); ch.set_size({"width": 480, "height": 660})
R.insert_chart("J10", ch)

rr = son_test + 3
R.merge_range(rr - 1, 1, rr - 1, 11, "En riskli 20 ödeme (kural tabanlı skor; eşitlikte satır sırası)", f_bolum)
R.write_row(rr, 1, ["Sıra", "Tedarikçi ünvanı", "Ödeme ID", "Tedarikçi", "Tarih", "Tutar (TL)", "Skor", "Seviye",
                    "Python IF", "Satır", "Tetiklenen testler"], f_h)
R.set_column("L:L", 40)
for k in range(1, 21):
    i = rr + k
    x = i + 1  # Excel satır no
    R.write(i, 1, k, f_h2)
    R.write_formula(i, 10, f"=MATCH(LARGE(TO_SiraAnahtar,B{x}),TO_SiraAnahtar,0)", f_tam)
    R.write_formula(i, 3, f"=INDEX(O_ID,K{x})")
    R.write_formula(i, 4, f"=INDEX(O_Ted,K{x})")
    R.write_formula(i, 2, f"=INDEX(Ted_Unvan,MATCH(E{x},Ted_ID,0))")
    R.write_formula(i, 5, f"=INDEX(O_Tarih,K{x})", f_tarih)
    R.write_formula(i, 6, f"=INDEX(O_TL,K{x})", f_tl0)
    R.write_formula(i, 7, f"=INDEX(TO_Skor,K{x})", f_skor)
    R.write_formula(i, 8, f"=INDEX(TO_Seviye,K{x})")
    R.write_formula(i, 9, f"=INDEX(TO_PyIF,K{x})", f_ondalik)
    R.write_array_formula(i, 11, i, 11, fx(f'{{=TEXTJOIN(", ",TRUE,IF(INDEX(TO_Bayraklar,K{x},0)=1,TO_Kodlar,""))}}'))
seviye_bicim(R, 8, rr + 21)
rr += 24
R.merge_range(rr - 1, 1, rr - 1, 11, "En riskli 10 tedarikçi (önce organizasyon bayrağı, sonra ödeme risk puanı)", f_bolum)
R.write_row(rr, 1, ["Sıra", "Ünvan", "Tedarikçi", "Org bayrak", "Ödeme risk", "Ödeme adedi", "Toplam TL",
                    "Sözleşme kul.", "Durum", "Satır", "Bayraklı organizasyon testleri"], f_h)
g_ilk, g_son = H_TG["G01"], H_TG["G12"]
for k in range(1, 11):
    i = rr + k
    x = i + 1
    R.write(i, 1, k, f_h2)
    R.write_formula(i, 10, f"=MATCH(LARGE(TOrg_SiraAnahtar,B{x}),TOrg_SiraAnahtar,0)", f_tam)
    R.write_formula(i, 3, f"=INDEX(Ted_ID,K{x})")
    R.write_formula(i, 2, f"=INDEX(Ted_Unvan,K{x})")
    R.write_formula(i, 4, f"=INDEX(TOrg_Skor,K{x})", f_skor)
    R.write_formula(i, 5, f"=INDEX(TOrg_OdemeRisk,K{x})", f_tl0)
    R.write_formula(i, 6, f"=INDEX(TOrg_N,K{x})", f_tam)
    R.write_formula(i, 7, f"=INDEX(TOrg_Toplam,K{x})", f_tl0)
    R.write_formula(i, 8, f"=INDEX(T_Org!${H_TG['soz']}$2:${H_TG['soz']}${NT + 1},K{x})", f_yuzde)
    R.write_formula(i, 9, f"=INDEX(Ted_Durum,K{x})")
    R.write_array_formula(i, 11, i, 11, fx(
        f'{{=TEXTJOIN(", ",TRUE,IF(INDEX(T_Org!${g_ilk}$2:${g_son}${NT + 1},K{x},0)=1,T_Org!${g_ilk}$1:${g_son}$1,""))}}'))
rr += 13
R.merge_range(rr - 1, 1, rr - 1, 11, "En riskli 10 tahsilat", f_bolum)
R.write_row(rr, 1, ["Sıra", "Müşteri", "Tahsilat ID", "Tarih", "Tutar (TL)", "Skor", "Seviye", "Kasiyer", "Yöntem",
                    "Satır", "Tetiklenen testler"], f_h)
h_ilk, h_son = H_TH["H01"], H_TH["H08"]
for k in range(1, 11):
    i = rr + k
    x = i + 1
    R.write(i, 1, k, f_h2)
    R.write_formula(i, 10, f"=MATCH(LARGE(TH_SiraAnahtar,B{x}),TH_SiraAnahtar,0)", f_tam)
    R.write_formula(i, 2, f"=INDEX(H_Musteri,K{x})")
    R.write_formula(i, 3, f"=INDEX(H_ID,K{x})")
    R.write_formula(i, 4, f"=INDEX(H_Tarih,K{x})", f_tarih)
    R.write_formula(i, 5, f"=INDEX(H_Tutar,K{x})", f_tl0)
    R.write_formula(i, 6, f"=INDEX(TH_Skor,K{x})", f_skor)
    R.write_formula(i, 7, f"=INDEX(TH_Seviye,K{x})")
    R.write_formula(i, 8, f"=INDEX(H_Kasiyer,K{x})")
    R.write_formula(i, 9, f"=INDEX(H_Yontem,K{x})")
    R.write_array_formula(i, 11, i, 11, fx(
        f'{{=TEXTJOIN(", ",TRUE,IF(INDEX(T_Tahsilat!${h_ilk}$2:${h_son}${NH + 1},K{x},0)=1,'
        f'T_Tahsilat!${h_ilk}$1:${h_son}$1,""))}}'))
seviye_bicim(R, 7, rr + 11)
R.set_column("K:K", 6)
R.set_tab_color(LACI)

# ================================================================== KILAVUZ
K = ws_kil
K.set_column("A:A", 2); K.set_column("B:B", 24); K.set_column("C:C", 110)
K.write("B1", "İç Denetim · Suistimal (Hile) Denetimi · Anomali Analitiği Modeli", f_baslik)
K.write("B2", "Ödemeler, tahsilatlar, organizasyon doğrulama, hesaplama, istatistiksel ve gizli örüntü testleri + web kazıma "
              "(TCMB, resmi tatiller, OFAC, Reddit, Big Four) + Python makine öğrenmesi.", f_alt)
canli_ozet = "; ".join(f"{r['kaynak']}: {r['durum']}" for r in LOG)
satirlar = [
    ("Amaç", "Muhasebe hareketlerinde hata ve hile göstergelerini (red flag) kural tabanlı, istatistiksel ve makine "
             "öğrenmesi yöntemleriyle tespit edip risk skoruna göre önceliklendirmek. Model denetçi muhakemesinin yerini "
             "almaz; bayraklar 'inceleme önceliği'dir, hile kanıtı değildir (ISA 240)."),
    ("Akış", "1) src/kazima.py → dış veriler (Kur, Tatil, Yaptirim, Topluluk)  2) veri/*.csv → ERP dökümü (bu dosyada "
             "sentetik + bilinçli yerleştirilmiş hileler)  3) src/analiz.py → Python ileri analitik  4) src/excel_model.py → "
             "bu çalışma kitabı  5) src/dogrula.py → yakalama oranı raporu. Tümü: python src/calistir.py"),
    ("Risk_Panosu", "KPI'lar, test isabetleri, en riskli 20 ödeme / 10 tedarikçi / 10 tahsilat ve tetiklenen testler (TEXTJOIN)."),
    ("Parametreler", "Tüm eşikler ve test ağırlıkları (sarı hücreler). Formüller bu hücrelere adla bağlıdır."),
    ("T_Odeme", "Her ödeme için 26 test (T01–T26): mükerrer, bulanık mükerrer, hafta sonu/tatil/mesai dışı, yuvarlak, "
                "limit altı, bölünmüş ödeme, görevler ayrılığı, yetki aşımı, matrah/KDV/brüt/kur yeniden hesaplama, kesim, "
                "kayıt öncesi ödeme, pasif tedarikçi, Z skoru, IQR, tekrar, yeni tedarikçi, kabuk şirket, Benford, RSF, "
                "sözleşme aşımı. Risk skoru = SUMPRODUCT(bayraklar; ağırlıklar)."),
    ("T_Tahsilat", "8 test (H01–H08): lapping, nakdin geç yatırılması, eşik altı nakit (structuring), makbuz boşluğu, "
                   "hafta sonu, fazla tahsilat, mükerrer, yuvarlak nakit."),
    ("T_Org / T_Personel", "Organizasyon doğrulama: VKN (GİB algoritması), TCKN, IBAN mod-97 (ISO 13616) formülle doğrulanır; "
                           "tedarikçi–personel IBAN/adres/telefon eşleşmesi, mükerrer VKN, yaptırım listesi, kaydı açan = "
                           "onaylayan, ardışık fatura (kabuk şirket), tedarikçi bazlı Benford ki-kare, RSF (AGGREGATE)."),
    ("Benford", "İlk hane, ilk iki hane, son iki hane testleri; Z, ki-kare, MAD ve Nigrini uyum sınıfı; grafikler."),
    ("Istatistik", "SUBTOTAL, AGGREGATE persentilleri, IQR, FREQUENCY histogramı, aylık FORECAST.LINEAR sapma testi, CORREL, "
                   "haftanın günü / saat örüntüsü, onaylayan kullanıcı analizi."),
    ("Dinamik_365", "Görseldeki 18 fonksiyonun tamamı (TAKE, DROP, VSTACK, HSTACK, TOCOL, TOROW, SORTBY, GROUPBY, PIVOTBY, "
                    "SUBTOTAL, AGGREGATE, FREQUENCY, FORECAST.LINEAR, CORREL, LAMBDA, MAP, BYROW, SCAN) + LET, FILTER, UNIQUE, "
                    "XLOOKUP, CHOOSECOLS, REDUCE ile denetim senaryoları. Microsoft 365 gerektirir."),
    ("Python_Bulgular", "Benford KS/toplam testi, Isolation Forest, varlık ağı (union-find), bulanık ünvan/yaptırım eşleşmesi, "
                        "kullanıcı saat dağılımı ki-kare, lapping zinciri, makbuz boşlukları."),
    ("Test_Kutuphanesi", "Her testin ACFE şeması, mantığı, kullanılan fonksiyonlar, referansı, isabet sayısı ve Reddit/Big Four "
                         "topluluk sinyali."),
    ("Veri sayfaları", "Odeme, Tahsilat, Tedarikci, Personel, Kullanici (ERP) · Kur, Tatil, Yaptirim, Topluluk, Kazima_Log (web)."),
    ("Renk kodu", "Mavi yazı = sabit girdi/veri · Siyah = formül · Yeşil = başka sayfaya bağlantı · Sarı zemin = "
                  "kullanıcının değiştirebileceği varsayım · Kırmızı zemin = tetiklenen bayrak / yüksek risk."),
    ("Kendi verinizle", "veri/*.csv dosyalarını aynı sütun başlıklarıyla ERP dökümünüzle değiştirin, veri_uret.py adımını "
                        "atlayın: python src/calistir.py --gercek-veri. Adlandırılmış aralıklar satır sayısına göre yeniden kurulur."),
    ("Web kazıma durumu", canli_ozet + ". YEDEK: kaynak erişilemedi, çevrimdışı yedek kullanıldı – yedek kur serisi SENTETİK, "
                          "yedek yaptırım listesi KURGUSALDIR; gerçek denetimde internet erişimli makinede kazımayı çalıştırın."),
    ("Uyarı", "Veriler sentetiktir; kişi/şirket adları kurgusaldır. Eşikler (onay limiti, nakit eşiği vb.) kurum politikasına "
              "ve güncel mevzuata göre güncellenmelidir."),
    ("Kaynakça", "ACFE Report to the Nations & Fraud Tree · Nigrini, M. (2012) Benford's Law, Wiley · IIA GTAG 13 & 16 · "
                 "ISA/ISA 240, 500, 520 · COSO 2013 · Liu, Ting & Zhou (2008) Isolation Forest · GİB VKN ve NVİ TCKN "
                 "algoritmaları · ISO 13616 IBAN."),
]
K.write_row("B4", ["Bölüm", "Açıklama"], f_h)
for i, (b, c) in enumerate(satirlar, 5):
    K.write(f"B{i}", b, F(bold=True, valign="top", bg_color=ACIK, border=1))
    K.write(f"C{i}", c, F(text_wrap=True, valign="top", border=1))
    K.set_row(i - 1, 15 * max(1, len(c) // 105 + 1))
K.activate()
ws_kil.set_first_sheet()
ws_kil.set_tab_color(LACI)

wb.close()
print(f"Yazıldı: {DOSYA}")
