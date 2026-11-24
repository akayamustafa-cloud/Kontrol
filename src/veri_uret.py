"""Sentetik denetim veri seti üretici – bilinçli olarak yerleştirilmiş (ekilmiş) hile senaryolarıyla.

Gerçek veriniz varsa bu adımı atlayın ve veri/*.csv dosyalarını aynı sütun yapısında kendi
ERP dökümünüzle değiştirin. Ekilen anomaliler veri/gercek_anomaliler.csv dosyasına yazılır ve
dogrula.py modelin bunları yakalayıp yakalamadığını ölçer (recall).
"""
from __future__ import annotations

import bisect
import csv
import datetime as dt
import math
import random

from ortak import DIS_VERI, VERI, iban_olustur, tckn_tamamla, vkn_kontrol_hanesi

YIL = 2025
rnd = random.Random(2025)
GT: list[dict] = []  # ground truth


def gt(tablo, kimlik, tip, aciklama):
    GT.append(dict(Tablo=tablo, ID=kimlik, Tip=tip, Aciklama=aciklama))


# ---------------------------------------------------------------- dış veri
with open(DIS_VERI / "kurlar.csv", encoding="utf-8") as f:
    KURLAR = [(dt.date.fromisoformat(r["Tarih"]), {k: float(r[k]) for k in ("USD", "EUR", "GBP")})
              for r in csv.DictReader(f)]
KUR_TARIH = [k[0] for k in KURLAR]
with open(DIS_VERI / "tatiller.csv", encoding="utf-8") as f:
    TATIL = {dt.date.fromisoformat(r["Tarih"]) for r in csv.DictReader(f)}


def kur(tarih: dt.date, pb: str) -> float:
    if pb == "TRY":
        return 1.0
    i = bisect.bisect_right(KUR_TARIH, tarih) - 1
    return KURLAR[i][1][pb]


IS_GUNLERI = [d for d in (dt.date(YIL, 1, 1) + dt.timedelta(n) for n in range(365))
              if d.weekday() < 5 and d not in TATIL]
HAFTA_SONU = [d for d in (dt.date(YIL, 1, 1) + dt.timedelta(n) for n in range(365)) if d.weekday() >= 5]
TATIL_ISGUNU = sorted(d for d in TATIL if d.year == YIL and d.weekday() < 5)

ADLAR = ["Ahmet", "Mehmet", "Ayşe", "Fatma", "Mustafa", "Emine", "Ali", "Hatice", "Hüseyin", "Zeynep",
         "İbrahim", "Elif", "Murat", "Merve", "Ömer", "Esra", "Yusuf", "Büşra", "Hasan", "Seda", "Kemal",
         "Gamze", "Serkan", "Derya", "Okan", "Pınar", "Cem", "Tuğba", "Onur", "Gizem"]
SOYADLAR = ["Yılmaz", "Kaya", "Demir", "Şahin", "Çelik", "Yıldız", "Yıldırım", "Öztürk", "Aydın", "Özdemir",
            "Arslan", "Doğan", "Kılıç", "Aslan", "Çetin", "Kara", "Koç", "Kurt", "Özkan", "Şimşek", "Polat",
            "Erdoğan", "Güneş", "Aksoy", "Bulut", "Tekin", "Ünal", "Acar", "Eren", "Tuncer"]
ILCELER = ["Kadıköy", "Üsküdar", "Beşiktaş", "Şişli", "Bakırköy", "Ataşehir", "Maltepe", "Kartal",
           "Pendik", "Beylikdüzü", "Esenyurt", "Sarıyer", "Bağcılar", "Ümraniye", "Çankaya", "Bornova"]
SOKAKLAR = ["Atatürk Cd.", "Cumhuriyet Cd.", "İnönü Sk.", "Bağdat Cd.", "Gazi Blv.", "Fatih Sk.",
            "Lale Sk.", "Menekşe Sk.", "Çınar Sk.", "Barış Cd.", "Mimar Sinan Cd.", "Yavuz Sk."]
BANKALAR = ["00010", "00012", "00015", "00032", "00046", "00062", "00064", "00067", "00111", "00134"]


def adres():
    return f"{rnd.choice(SOKAKLAR)} No:{rnd.randint(1, 180)} D:{rnd.randint(1, 24)} {rnd.choice(ILCELER)}"


def telefon():
    return f"05{rnd.randint(30, 59)}{rnd.randint(1000000, 9999999)}"


def iban():
    return iban_olustur(rnd.choice(BANKALAR), str(rnd.randint(10 ** 15, 10 ** 16 - 1)))


def vkn():
    ilk9 = str(rnd.randint(10 ** 8, 10 ** 9 - 1))
    return ilk9 + str(vkn_kontrol_hanesi(ilk9))


def tckn():
    return tckn_tamamla(str(rnd.randint(1, 9)) + "".join(str(rnd.randint(0, 9)) for _ in range(8)))


def bozuk(no: str) -> str:
    """Kontrol hanesini bozarak geçersiz kimlik/IBAN üretir."""
    son = int(no[-1])
    return no[:-1] + str((son + rnd.randint(1, 8)) % 10)


# ---------------------------------------------------------------- Kullanıcılar
KULLANICILAR = [
    ("U01", "Ayşe Kaya", "Muhasebe Uzmanı", 50_000), ("U02", "Mehmet Demir", "Muhasebe Uzmanı", 50_000),
    ("U03", "Can Yılmaz", "Satınalma Şefi", 100_000), ("U04", "Elif Şahin", "Finans Şefi", 100_000),
    ("U05", "Burak Çelik", "Finans Müdürü", 250_000), ("U06", "Zeynep Arslan", "Satınalma Müdürü", 250_000),
    ("U07", "Hakan Öztürk", "Mali İşler Direktörü", 1_000_000), ("U08", "Selin Aydın", "CFO", 5_000_000),
    ("U09", "Emre Koç", "Muhasebe Personeli", 0), ("U10", "Deniz Kurt", "Muhasebe Personeli", 0),
]
LIMIT = {k[0]: k[3] for k in KULLANICILAR}
OLUSTURANLAR = ["U09", "U10", "U01", "U02"]


def onaylayan_sec(tutar, olusturan):
    aday = [u for u, l in LIMIT.items() if l >= tutar and u != olusturan]
    en_dusuk = min(LIMIT[u] for u in aday)
    # çoğunlukla yeterli en düşük kademe, bazen bir üst kademe onaylar
    kademe = [u for u in aday if LIMIT[u] == en_dusuk]
    if rnd.random() < 0.25:
        kademe = aday
    return rnd.choice(kademe)


# ---------------------------------------------------------------- Personel
personel = []
for i in range(1, 81):
    personel.append(dict(PersonelID=f"P{i:03d}", AdSoyad=f"{rnd.choice(ADLAR)} {rnd.choice(SOYADLAR)}",
                         TCKN=tckn(), IBAN=iban(), Adres=adres(), Telefon=telefon(),
                         Departman=rnd.choice(["Muhasebe", "Satınalma", "Satış", "Üretim", "Lojistik", "İK", "BT"]),
                         IseGiris=(dt.date(2012, 1, 1) + dt.timedelta(rnd.randint(0, 4700))).isoformat()))
P = {p["PersonelID"]: p for p in personel}
P["P045"]["TCKN"] = bozuk(P["P045"]["TCKN"]); gt("Personel", "P045", "P01", "Geçersiz TCKN")
P["P061"]["TCKN"] = bozuk(P["P061"]["TCKN"]); gt("Personel", "P061", "P01", "Geçersiz TCKN")
P["P079"]["IBAN"] = P["P012"]["IBAN"]; gt("Personel", "P079", "P02", "Hayalet personel: P012 ile aynı IBAN")
gt("Personel", "P012", "P02", "Hayalet personel: P079 ile aynı IBAN")
P["P070"]["IBAN"] = bozuk(P["P070"]["IBAN"]); gt("Personel", "P070", "P03", "Geçersiz IBAN")

# ---------------------------------------------------------------- Tedarikçiler
ON = ["Anadolu", "Ege", "Marmara", "Yıldız", "Doğu", "Akdeniz", "Toros", "Kapadokya", "Boğaziçi", "Trakya",
      "Kuzey", "Güney", "Başkent", "Lider", "Özgür", "Atlas", "Zirve", "Mavi", "Yeşil", "Altın", "Gümüş",
      "Kartal", "Çınar", "Defne", "Sedef", "Pınar", "Uğur", "Erciyes", "Uludağ", "Fırat"]
SEKTOR = [("Lojistik", "Nakliye hizmeti"), ("Gıda", "Yemek hizmeti"), ("İnşaat", "Bakım onarım"),
          ("Bilişim", "Yazılım lisansı"), ("Kırtasiye", "Ofis malzemesi"), ("Enerji", "Elektrik/doğalgaz"),
          ("Ambalaj", "Ambalaj malzemesi"), ("Makine", "Yedek parça"), ("Temizlik", "Temizlik hizmeti"),
          ("Güvenlik", "Güvenlik hizmeti"), ("Kimya", "Hammadde"), ("Tekstil", "İş kıyafeti"),
          ("Danışmanlık", "Danışmanlık hizmeti"), ("Matbaa", "Baskı hizmeti"), ("Otomotiv", "Araç kiralama")]
YABANCI = [("Nordic Tooling AB", "EUR"), ("Pacific Components Ltd", "USD"), ("Brightline Software Inc", "USD"),
           ("Rhein Maschinen GmbH", "EUR"), ("Oakfield Analytics LLC", "USD"), ("Crown Logistics Plc", "GBP"),
           ("Sierra Polymers Inc", "USD"), ("Lyon Instruments SA", "EUR"), ("Harbor Freight Systems LLC", "USD"),
           ("Delta Optics Corp", "USD")]

tedarikciler, prefix_kullanilan = [], set()


def prefix_uret(ad):
    harf = [c for c in ad.upper().translate(str.maketrans("ÇĞİÖŞÜ", "CGIOSU")) if c.isalpha()]
    while True:
        p = "".join(rnd.sample(harf, 3)) if len(harf) >= 3 else "XYZ"
        if p not in prefix_kullanilan:
            prefix_kullanilan.add(p)
            return p


kullanilan_ad = {("Toros", "Ambalaj"), ("Sedef", "Temizlik"), ("Uludağ", "Makine")}
yabanci_idx = set(rnd.sample(range(1, 61), len(YABANCI)))
yi = iter(YABANCI)
for i in range(1, 61):
    if i in yabanci_idx:
        ad, pb = next(yi)
        kat, acik = "Yurt dışı", "İthal hizmet/malzeme"
    else:
        while True:
            kat, acik = rnd.choice(SEKTOR)
            on = rnd.choice(ON)
            if (on, kat) not in kullanilan_ad:
                kullanilan_ad.add((on, kat))
                break
        ad = f"{on} {kat} {rnd.choice(['A.Ş.', 'Ltd. Şti.', 'San. ve Tic. Ltd. Şti.'])}"
        pb = "TRY"
    tedarikciler.append(dict(TedarikciID=f"V{i:03d}", Unvan=ad, VKN=vkn(), IBAN=iban(), Adres=adres(),
                             Telefon=telefon(),
                             KayitTarihi=(dt.date(2015, 1, 1) + dt.timedelta(rnd.randint(0, 3400))).isoformat(),
                             Olusturan=rnd.choice(["U09", "U10"]), Durum="Aktif", Kategori=kat,
                             ParaBirimi=pb, SozlesmeLimitiTL=0, _acik=acik, _pfx=prefix_uret(ad),
                             _olcek=rnd.uniform(8.3, 10.6), _sigma=rnd.uniform(0.55, 1.25),
                             _agirlik=rnd.paretovariate(1.3), _fno=rnd.randint(1000, 90000)))
# "Kira" tedarikçisi: her ay sabit tutar (meşru tekrar – yanlış pozitif tartışması için)
kira = tedarikciler[29]
kira.update(Unvan="Başkent Gayrimenkul Kiralama A.Ş.", Kategori="Gayrimenkul", ParaBirimi="TRY", _acik="Ofis kirası")
yabanci_idx.discard(30)

CIFT = next(v for v in tedarikciler if v["ParaBirimi"] == "TRY" and v is not kira)
ORTAK_IBAN = next(v for v in tedarikciler[15:] if v["ParaBirimi"] == "TRY" and v is not kira)
PLANTED = [
    dict(TedarikciID="V061", Unvan="Mavi Ufuk Danışmanlık Ltd. Şti.", VKN=vkn(), IBAN=P["P017"]["IBAN"],
         Adres=adres(), Telefon=telefon(), KayitTarihi="2025-03-10", Olusturan="U03", Durum="Aktif",
         Kategori="Danışmanlık", ParaBirimi="TRY", _acik="Danışmanlık hizmeti", _pfx="MVD", _fno=101),
    dict(TedarikciID="V062", Unvan="Kuzey Yıldızı Organizasyon Ltd. Şti.", VKN=bozuk(vkn()), IBAN=iban(),
         Adres=P["P033"]["Adres"], Telefon=P["P033"]["Telefon"], KayitTarihi="2025-06-15", Olusturan="U06",
         Durum="Aktif", Kategori="Organizasyon", ParaBirimi="TRY", _acik="Etkinlik organizasyonu", _pfx="KYO",
         _fno=2041),
    dict(TedarikciID="V063", Unvan="Karadeniz Global Trading Ltd.", VKN=vkn(), IBAN=iban(), Adres=adres(),
         Telefon=telefon(), KayitTarihi="2025-02-03", Olusturan="U09", Durum="Aktif", Kategori="Dış ticaret",
         ParaBirimi="USD", _acik="İthal hammadde", _pfx="KGT", _fno=55000),
    dict(TedarikciID="V064", Unvan=CIFT["Unvan"].upper().replace("Ş", "S").replace(".", ""),
         VKN=CIFT["VKN"], IBAN=iban(), Adres=CIFT["Adres"], Telefon=telefon(),
         KayitTarihi="2025-04-22", Olusturan="U09", Durum="Aktif", Kategori=CIFT["Kategori"],
         ParaBirimi="TRY", _acik=CIFT["_acik"], _pfx=CIFT["_pfx"], _fno=0),
    dict(TedarikciID="V065", Unvan="Toros Ambalaj Ltd. Şti.", VKN=vkn(), IBAN=iban(), Adres=adres(),
         Telefon=telefon(), KayitTarihi="2016-05-11", Olusturan="U10", Durum="Pasif", Kategori="Ambalaj",
         ParaBirimi="TRY", _acik="Ambalaj malzemesi", _pfx="TRA", _fno=30500),
    dict(TedarikciID="V066", Unvan="Sedef Temizlik Hizmetleri Ltd. Şti.", VKN=vkn(), IBAN=ORTAK_IBAN["IBAN"],
         Adres=adres(), Telefon=telefon(), KayitTarihi="2025-01-20", Olusturan="U09", Durum="Aktif",
         Kategori="Temizlik", ParaBirimi="TRY", _acik="Temizlik hizmeti", _pfx="SDT", _fno=8000),
    dict(TedarikciID="V067", Unvan="Uludağ Makine San. ve Tic. A.Ş.", VKN=vkn(), IBAN=bozuk(iban()),
         Adres=adres(), Telefon=telefon(), KayitTarihi="2019-09-09", Olusturan="U09", Durum="Aktif",
         Kategori="Makine", ParaBirimi="TRY", _acik="Yedek parça", _pfx="UMS", _fno=14000),
]
for v in PLANTED:
    v.setdefault("SozlesmeLimitiTL", 0)
    tedarikciler.append(v)
gt("Tedarikci", "V061", "G05", "Hayali tedarikçi: IBAN personel P017 ile aynı")
gt("Tedarikci", "V061", "G09", "Kaydı açan U03 ödemeleri de onaylıyor")
gt("Tedarikci", "V061", "G10", "Ardışık fatura numaraları (kabuk şirket)")
gt("Tedarikci", "V061", "G11", "Benford dışı uydurma tutarlar")
gt("Tedarikci", "V061", "G12", "Yuvarlak tutar ağırlıklı")
gt("Tedarikci", "V062", "G01", "Geçersiz VKN")
gt("Tedarikci", "V062", "G06", "Adres personel P033 ile aynı")
gt("Tedarikci", "V062", "G07", "Telefon personel P033 ile aynı")
gt("Tedarikci", "V062", "G10", "Ardışık fatura numaraları")
gt("Tedarikci", "V063", "G08", "Yaptırım listesi ile ünvan benzerliği")
gt("Tedarikci", "V064", "G02", f"{CIFT['TedarikciID']} ile aynı VKN – çift kayıt")
gt("Tedarikci", CIFT["TedarikciID"], "G02", "V064 ile aynı VKN – çift kayıt")
gt("Tedarikci", "V066", "G04", f"{ORTAK_IBAN['TedarikciID']} ile ortak IBAN")
gt("Tedarikci", ORTAK_IBAN["TedarikciID"], "G04", "V066 ile ortak IBAN")
gt("Tedarikci", "V067", "G03", "Geçersiz IBAN")
T = {v["TedarikciID"]: v for v in tedarikciler}

# ---------------------------------------------------------------- Ödemeler
odemeler = []


def saat_normal():
    dk = rnd.randint(8 * 60 + 30, 18 * 60 + 30)
    return f"{dk // 60:02d}:{dk % 60:02d}"


def satir_olustur(v, tarih, tl_hedef, olusturan=None, onaylayan=None, kdv_oran=None, brut_sabit=None, saat=None):
    pb = v["ParaBirimi"]
    k = kur(tarih, pb)
    if kdv_oran is None:
        kdv_oran = 0.0 if pb != "TRY" else rnd.choices([0.20, 0.10, 0.01], [0.85, 0.12, 0.03])[0]
    if brut_sabit is not None:
        brut = round(brut_sabit, 2)
        matrah = round(brut / (1 + kdv_oran), 2)
        miktar, bf = 1, matrah
        kdv = round(brut - matrah, 2)
    else:
        matrah_hedef = tl_hedef / k / (1 + kdv_oran)
        miktar = rnd.choice([1, 1, 1, 2, 3, 5, 10, 12, 20, 25, 50, 100, 200, 250, 500])
        bf = round(matrah_hedef / miktar, 2) or 0.01
        matrah = round(miktar * bf, 2)
        kdv = round(matrah * kdv_oran, 2)
        brut = round(matrah + kdv, 2)
    tl = round(brut * k, 2)
    olusturan = olusturan or rnd.choice(OLUSTURANLAR)
    onaylayan = onaylayan or onaylayan_sec(tl, olusturan)
    v["_fno"] += 1 if v["TedarikciID"] in ("V061", "V062") else rnd.randint(15, 420)
    fat_tarih = max(tarih - dt.timedelta(days=rnd.randint(5, 55)), dt.date(YIL - 1, 11, 1))
    if v["TedarikciID"] in ("V061", "V062"):
        fat_tarih = tarih - dt.timedelta(days=rnd.randint(1, 4))
    return dict(OdemeID="", OdemeTarihi=tarih, Saat=saat or saat_normal(), TedarikciID=v["TedarikciID"],
                FaturaNo=f"{v['_pfx']}{v['_fno']:06d}", FaturaTarihi=fat_tarih, Miktar=miktar, BirimFiyat=bf,
                Matrah=matrah, KDVOrani=kdv_oran, KDV=kdv, BrutTutar=brut, ParaBirimi=pb, Kur=k, TutarTL=tl,
                Olusturan=olusturan, Onaylayan=onaylayan,
                OdemeYontemi=rnd.choices(["Havale/EFT", "Çek", "DBS"], [0.8, 0.12, 0.08])[0],
                MasrafMerkezi=rnd.choice(["MM-100 Yönetim", "MM-200 Üretim", "MM-300 Satış", "MM-400 Lojistik", "MM-500 BT"]),
                Aciklama=v["_acik"], _etiket=[])


normal = [v for v in tedarikciler if v["TedarikciID"] <= "V060" and v["TedarikciID"] != "V030"]
agirliklar = [v["_agirlik"] for v in normal]
gunler_agirlik = [1.35 if d.month == 12 else (1.1 if d.month in (6, 9) else 1.0) for d in IS_GUNLERI]
for _ in range(1330):
    v = rnd.choices(normal, agirliklar)[0]
    tarih = rnd.choices(IS_GUNLERI, gunler_agirlik)[0]
    tl = min(math.exp(rnd.gauss(v["_olcek"], v["_sigma"])), 900_000)
    tl = max(tl, 150)
    odemeler.append(satir_olustur(v, tarih, tl))
# kira: her ay 5. iş günü sabit 85.000 TL
for ay in range(1, 13):
    tarih = [d for d in IS_GUNLERI if d.month == ay][4]
    odemeler.append(satir_olustur(kira, tarih, 0, kdv_oran=0.20, brut_sabit=85_000.0))

# --- V061 hayali danışmanlık: yuvarlak, Benford dışı, U03 oluşturup onaylıyor, sözleşme aşımı
v61_tutarlar = [40_000, 45_000, 60_000, 75_000, 80_000, 90_000, 40_000, 45_000, 70_000, 85_000, 95_000,
                40_000, 65_000, 75_000, 80_000, 90_000, 45_000, 60_000, 70_000, 95_000, 40_000, 80_000,
                85_000, 75_000, 90_000, 60_000, 45_000, 70_000, 95_000, 85_000, 40_000, 65_000]
v61_gunler = sorted(rnd.sample([d for d in IS_GUNLERI if d >= dt.date(YIL, 3, 25)], len(v61_tutarlar)))
ilk = satir_olustur(T["V061"], dt.date(YIL, 3, 20), 0, "U03", "U03", 0.20, brut_sabit=120_000.0)
ilk["_etiket"] += ["T22", "T09", "T06"]
odemeler.append(ilk)
for d, t in zip(v61_gunler, v61_tutarlar):
    r = satir_olustur(T["V061"], d, 0, "U03", "U03", 0.20, brut_sabit=float(t))
    r["_etiket"] += ["T09", "T06", "T18", "T23"]
    odemeler.append(r)
# --- V062: ardışık fatura, kayıttan önce ödeme (Mayıs), kayıt 15.06
for d in sorted(rnd.sample([x for x in IS_GUNLERI if x.month in (5, 7, 8, 9, 10, 11)], 14)):
    r = satir_olustur(T["V062"], d, rnd.uniform(18_000, 48_000))
    r["_etiket"] += ["T18", "T23"] + (["T16"] if d < dt.date(YIL, 6, 15) else [])
    odemeler.append(r)
# --- V063 yaptırım şüpheli
for d in sorted(rnd.sample([x for x in IS_GUNLERI if x >= dt.date(YIL, 2, 10)], 5)):
    r = satir_olustur(T["V063"], d, rnd.uniform(150_000, 600_000))
    r["_etiket"] += ["T18"]
    odemeler.append(r)
# --- V064 çift kayıt (V005 ile aynı VKN) ve V066/V067
for vid, n in (("V064", 4), ("V066", 6), ("V067", 5)):
    for d in sorted(rnd.sample([x for x in IS_GUNLERI if x >= dt.date.fromisoformat(T[vid]["KayitTarihi"])], n)):
        r = satir_olustur(T[vid], d, math.exp(rnd.gauss(9.3, 0.9)))
        r["_etiket"] += ["T18"]
        odemeler.append(r)
# --- V065 pasif tedarikçiye ödeme
for d in sorted(rnd.sample(IS_GUNLERI, 4)):
    r = satir_olustur(T["V065"], d, math.exp(rnd.gauss(9.8, 0.6)))
    r["_etiket"] += ["T17"]
    odemeler.append(r)

tr_normal = [r for r in odemeler if r["ParaBirimi"] == "TRY" and not r["_etiket"]
             and r["TedarikciID"] not in ("V030",) and r["TedarikciID"] <= "V060"]
kullanildi: set[int] = set()


def sec_normal(filtre=lambda r: True):
    aday = [r for r in tr_normal if id(r) not in kullanildi and filtre(r)]
    r = rnd.choice(aday)
    kullanildi.add(id(r))
    return r


def kopya(r, **deg):
    y = {k: (list(v) if isinstance(v, list) else v) for k, v in r.items()}
    y.update(deg)
    return y


def is_gunu_sonra(d, n):
    i = bisect.bisect_left(IS_GUNLERI, d)
    return IS_GUNLERI[min(i + n, len(IS_GUNLERI) - 1)]


# T01 tam mükerrer
for _ in range(6):
    r = sec_normal()
    y = kopya(r, OdemeTarihi=is_gunu_sonra(r["OdemeTarihi"], rnd.randint(3, 20)), Saat=saat_normal(), _etiket=["T01"])
    r["_etiket"].append("T01")
    odemeler.append(y)
# T02 bulanık mükerrer (fatura no formatı değiştirilmiş)
for bicim in (lambda p, n: f"{p}-{n:06d}", lambda p, n: f"{p} {n}", lambda p, n: f"{p}.{n:06d}",
              lambda p, n: f"{p}/{n:06d}"):
    r = sec_normal()
    p, n = r["FaturaNo"][:3], int(r["FaturaNo"][3:])
    y = kopya(r, OdemeTarihi=is_gunu_sonra(r["OdemeTarihi"], rnd.randint(5, 30)), FaturaNo=bicim(p, n),
              Saat=saat_normal(), _etiket=["T02"])
    r["_etiket"].append("T02")
    odemeler.append(y)
# T03 hafta sonu / T04 resmi tatil
def tarih_tasi(r, yeni):
    r["FaturaTarihi"] = yeni - (r["OdemeTarihi"] - r["FaturaTarihi"])
    r["OdemeTarihi"] = yeni


for _ in range(5):
    r = sec_normal(); tarih_tasi(r, rnd.choice(HAFTA_SONU)); r["_etiket"].append("T03")
for d in TATIL_ISGUNU[:4]:
    r = sec_normal(); tarih_tasi(r, d); r["_etiket"].append("T04")
# T05 mesai dışı
for _ in range(8):
    r = sec_normal(); s = rnd.choice([22, 23, 0, 1, 2, 3]); r["Saat"] = f"{s:02d}:{rnd.randint(0, 59):02d}"
    r["_etiket"].append("T05")
# T06 yuvarlak tutar (normal tedarikçilerde)
for _ in range(4):
    r = sec_normal()
    yeni = satir_olustur(T[r["TedarikciID"]], r["OdemeTarihi"], 0, r["Olusturan"], None, 0.20,
                         brut_sabit=float(rnd.choice([10_000, 20_000, 25_000, 30_000])))
    yeni["FaturaNo"] = r["FaturaNo"]; r.update(yeni); r["_etiket"] = ["T06"]
# T07+T08 bölünmüş ödemeler: limit 50.000 TL altı 3'lü gruplar
for _ in range(4):
    v = T[rnd.choice([x["TedarikciID"] for x in normal if x["ParaBirimi"] == "TRY"])]
    d0 = rnd.choice(IS_GUNLERI[5:-10])
    onay = rnd.choice(["U01", "U02"])
    for k in range(3):
        r = satir_olustur(v, is_gunu_sonra(d0, rnd.randint(0, 3)), rnd.uniform(45_500, 49_900), "U09", onay)
        r["_etiket"] += ["T07", "T08"]
        odemeler.append(r)
# T07 tekil limit altı
for _ in range(5):
    r = sec_normal(lambda r: r["Onaylayan"] in ("U01", "U02"))
    yeni = satir_olustur(T[r["TedarikciID"]], r["OdemeTarihi"], rnd.uniform(46_000, 49_950), r["Olusturan"],
                         r["Onaylayan"])
    yeni["FaturaNo"] = r["FaturaNo"]; r.update(yeni); r["_etiket"] = ["T07"]
# T09 görevler ayrılığı
for _ in range(8):
    r = sec_normal(lambda r: r["TutarTL"] <= 50_000)
    u = rnd.choice(["U01", "U02"]); r["Olusturan"] = u; r["Onaylayan"] = u; r["_etiket"].append("T09")
# T10 onay yetkisi aşımı
for _ in range(6):
    r = sec_normal(lambda r: 60_000 <= r["TutarTL"] <= 240_000)
    r["Onaylayan"] = rnd.choice(["U01", "U02", "U03"] if r["TutarTL"] > 100_000 else ["U01", "U02"])
    if r["Onaylayan"] == r["Olusturan"]:
        r["Olusturan"] = "U09"
    r["_etiket"].append("T10")
# T11 matrah şişirme, T12 KDV hatası, T13 brüt hatası
for _ in range(6):
    r = sec_normal(lambda r: r["Miktar"] > 1)
    r["Matrah"] = round(r["Matrah"] + rnd.uniform(500, 5000), 2)
    r["KDV"] = round(r["Matrah"] * r["KDVOrani"], 2); r["BrutTutar"] = round(r["Matrah"] + r["KDV"], 2)
    r["TutarTL"] = r["BrutTutar"]; r["_etiket"].append("T11")
for _ in range(5):
    r = sec_normal(lambda r: r["KDVOrani"] > 0)
    r["KDV"] = round(r["KDV"] + rnd.choice([-1, 1]) * rnd.uniform(150, 2500), 2)
    r["BrutTutar"] = round(r["Matrah"] + r["KDV"], 2); r["TutarTL"] = r["BrutTutar"]; r["_etiket"].append("T12")
for _ in range(4):
    r = sec_normal()
    r["BrutTutar"] = round(r["BrutTutar"] + rnd.uniform(300, 3000), 2); r["TutarTL"] = r["BrutTutar"]
    r["_etiket"].append("T13")
# T14 kur manipülasyonu (döviz)
doviz = [r for r in odemeler if r["ParaBirimi"] != "TRY" and not r["_etiket"]]
for r in rnd.sample(doviz, 7)[:5]:
    r["Kur"] = round(r["Kur"] * rnd.uniform(1.025, 1.06), 4); r["TutarTL"] = round(r["BrutTutar"] * r["Kur"], 2)
    r["_etiket"].append("T14")
for r in [x for x in doviz if not x["_etiket"]][:2]:
    r["TutarTL"] = round(r["TutarTL"] * 1.04, 2); r["_etiket"].append("T14")
# T15 fatura tarihi > ödeme tarihi
for _ in range(4):
    r = sec_normal(); r["FaturaTarihi"] = r["OdemeTarihi"] + dt.timedelta(days=rnd.randint(3, 25)); r["_etiket"].append("T15")
# T19/T20/T25 tekil aşırı büyük ödeme
# (dar dağılımlı, yeterli işlem hacmi olan tedarikçilerde; tutar = tipik tutarın 30–50 katı)
adet = {}
for r in odemeler:
    adet[r["TedarikciID"]] = adet.get(r["TedarikciID"], 0) + 1
uygun = [v["TedarikciID"] for v in normal if v["_sigma"] < 0.9 and adet.get(v["TedarikciID"], 0) >= 10
         and v["ParaBirimi"] == "TRY"]
for vid in rnd.sample(uygun, 5):
    r = sec_normal(lambda r, vid=vid: r["TedarikciID"] == vid)
    yeni = satir_olustur(T[vid], r["OdemeTarihi"], math.exp(T[vid]["_olcek"]) * rnd.uniform(30, 50), r["Olusturan"])
    yeni["FaturaNo"] = r["FaturaNo"]; r.update(yeni); r["_etiket"] = ["T19", "T25"]

# ---------------------------------------------------------------- sıralama, ID, sözleşme limitleri
odemeler.sort(key=lambda r: (r["OdemeTarihi"], r["Saat"]))
for i, r in enumerate(odemeler, 1):
    r["OdemeID"] = f"OD{i:05d}"
toplam = {}
for r in odemeler:
    toplam[r["TedarikciID"]] = toplam.get(r["TedarikciID"], 0) + r["TutarTL"]
for v in tedarikciler:
    t = toplam.get(v["TedarikciID"], 0)
    v["SozlesmeLimitiTL"] = max(100_000, math.ceil(t * rnd.uniform(1.25, 1.6) / 50_000) * 50_000)
T["V061"]["SozlesmeLimitiTL"] = 1_200_000
T["V015"]["SozlesmeLimitiTL"] = max(50_000, math.floor(toplam.get("V015", 0) * 0.8 / 50_000) * 50_000)
for vid in ("V061", "V015"):
    kum = 0
    for r in odemeler:
        if r["TedarikciID"] == vid:
            kum += r["TutarTL"]
            if kum > T[vid]["SozlesmeLimitiTL"]:
                r["_etiket"].append("T26")
for r in odemeler:
    for e in dict.fromkeys(r["_etiket"]):
        gt("Odeme", r["OdemeID"], e, r["TedarikciID"])

# ---------------------------------------------------------------- Tahsilatlar
musteriler = [f"C{i:03d}" for i in range(1, 41)]
tahsilat = []
fat_sayac = 50000
for _ in range(690):
    m = rnd.choice(musteriler)
    d = rnd.choice(IS_GUNLERI)
    fat_sayac += rnd.randint(1, 3)
    tut = round(math.exp(rnd.gauss(9.9, 1.0)), 2)
    yon = rnd.choices(["Havale/EFT", "Nakit", "Çek", "Kredi Kartı"], [0.68, 0.17, 0.10, 0.05])[0]
    if yon == "Nakit" and tut > 25_000:
        yon = "Havale/EFT"  # büyük tutarlar bankadan gelir
    yat = d + dt.timedelta(days=(rnd.choice([0, 0, 1]) if yon == "Nakit" else 0))
    tahsilat.append(dict(TahsilatID="", MakbuzNo=0, Tarih=d, OdeyenMusteriID=m, FaturaNo=f"SF{fat_sayac:06d}",
                         FaturaMusteriID=m, FaturaTutari=tut, TahsilatTutari=tut, Yontem=yon, BankayaYatirma=yat,
                         Kasiyer=rnd.choice(["K1", "K2", "K3", "K4"]), _etiket=[]))
# H01 lapping zinciri (K3): C011'in parası C012'nin faturasını kapatır ...
zincir = ["C011", "C012", "C013", "C014", "C015", "C016", "C017"]
d0 = dt.date(YIL, 8, 4)
for k in range(6):
    tut = round(rnd.uniform(18_000, 26_000), 2)
    fat_sayac += 1
    tahsilat.append(dict(TahsilatID="", MakbuzNo=0, Tarih=is_gunu_sonra(d0, 4 * k), OdeyenMusteriID=zincir[k + 1],
                         FaturaNo=f"SF{fat_sayac:06d}", FaturaMusteriID=zincir[k], FaturaTutari=tut,
                         TahsilatTutari=tut, Yontem="Havale/EFT", BankayaYatirma=is_gunu_sonra(d0, 4 * k),
                         Kasiyer="K3", _etiket=["H01"]))
nakit = [r for r in tahsilat if r["Yontem"] == "Nakit" and not r["_etiket"]]
for r in rnd.sample(nakit, 6):  # H02 geç yatırma
    r["BankayaYatirma"] = r["Tarih"] + dt.timedelta(days=rnd.randint(5, 12)); r["Kasiyer"] = "K2"; r["_etiket"].append("H02")
for k in range(6):  # H03 eşik altı yapılandırma (C020)
    d = is_gunu_sonra(dt.date(YIL, 10, 6), k)
    tut = round(rnd.uniform(27_200, 29_950), 2); fat_sayac += 1
    tahsilat.append(dict(TahsilatID="", MakbuzNo=0, Tarih=d, OdeyenMusteriID="C020", FaturaNo=f"SF{fat_sayac:06d}",
                         FaturaMusteriID="C020", FaturaTutari=tut, TahsilatTutari=tut, Yontem="Nakit",
                         BankayaYatirma=d, Kasiyer="K4", _etiket=["H03"]))
for r in rnd.sample([x for x in tahsilat if not x["_etiket"]], 3):  # H05 hafta sonu
    r["Tarih"] = rnd.choice(HAFTA_SONU); r["BankayaYatirma"] = r["Tarih"] + dt.timedelta(days=(7 - r["Tarih"].weekday()) % 7 or 1)
    r["_etiket"].append("H05")
for r in rnd.sample([x for x in tahsilat if not x["_etiket"] and x["Yontem"] != "Nakit"], 4):  # H06 fazla tahsilat
    r["TahsilatTutari"] = round(r["FaturaTutari"] * rnd.uniform(1.08, 1.5), 2); r["_etiket"].append("H06")
for r in rnd.sample([x for x in tahsilat if not x["_etiket"]], 3):  # H07 mükerrer
    y = dict(r); y["_etiket"] = ["H07"]; r["_etiket"] = ["H07"]; tahsilat.append(y)
for r in rnd.sample([x for x in tahsilat if not x["_etiket"] and x["Yontem"] == "Nakit"], 5):  # H08 yuvarlak nakit
    r["TahsilatTutari"] = r["FaturaTutari"] = float(rnd.choice([5_000, 10_000, 15_000, 20_000])); r["_etiket"].append("H08")
tahsilat.sort(key=lambda r: (r["Tarih"], r["FaturaNo"]))
makbuz = 100_000
silinen = set(rnd.sample(range(80, len(tahsilat) - 20), 4))
for i, r in enumerate(tahsilat, 1):
    makbuz += 2 if i in silinen else 1  # H04: silinmiş (kayıt dışı) makbuzlar boşluk bırakır
    r["MakbuzNo"] = makbuz
    r["TahsilatID"] = f"TH{i:05d}"
    if i in silinen:
        r["_etiket"].append("H04")
for r in tahsilat:
    for e in dict.fromkeys(r["_etiket"]):
        gt("Tahsilat", r["TahsilatID"], e, r["OdeyenMusteriID"])


# ---------------------------------------------------------------- yazma
def yaz(ad, satirlar):
    alanlar = [k for k in satirlar[0] if not k.startswith("_")]
    with open(VERI / ad, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=alanlar, extrasaction="ignore")
        w.writeheader()
        for s in satirlar:
            w.writerow({k: (v.isoformat() if isinstance(v, dt.date) else v) for k, v in s.items() if k in alanlar})


yaz("kullanicilar.csv", [dict(KullaniciID=a, AdSoyad=b, Rol=c, OnayLimitiTL=d) for a, b, c, d in KULLANICILAR])
yaz("personel.csv", personel)
yaz("tedarikciler.csv", tedarikciler)
yaz("odemeler.csv", odemeler)
yaz("tahsilatlar.csv", tahsilat)
yaz("gercek_anomaliler.csv", GT)
print(f"Ödeme {len(odemeler)} | Tahsilat {len(tahsilat)} | Tedarikçi {len(tedarikciler)} | "
      f"Personel {len(personel)} | Ekilen anomali kaydı {len(GT)}")
