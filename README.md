# Kontrol: Hile Denetimi ve Anomali Analitiği Modeli

İç denetim, suistimal (hile) denetimi ve anomali analizi için uçtan uca bir model. Ödemeler, tahsilatlar ve organizasyon verisi (tedarikçi, personel, kullanıcı yetkileri) üzerinde **doğrulama, hesaplama, istatistiksel, matematiksel ve gizli örüntü testleri** çalıştırılır. Dış veriler **web kazıma** ile toplanır (TCMB kurları, resmi tatiller, OFAC yaptırım listesi, Reddit ve Big Four/ACFE metinleri). Sonuçlar formüllerle çalışan bir **Excel modelinde** ve bir **Python makine öğrenmesi katmanında** birleşir.

> **Ana çıktı:** [`cikti/Hile_Denetim_Modeli.xlsx`](cikti/Hile_Denetim_Modeli.xlsx). 90.720 canlı formül içerir, formül hatası yoktur.
> Doğrulama: [`cikti/dogrulama_raporu.md`](cikti/dogrulama_raporu.md)

---

## 1. Mimari

```
  ┌────────────── WEB KAZIMA (src/kazima.py) ───────────────┐
  │ TCMB kur XML · Nager.Date tatil API · OFAC SDN CSV       │
  │ Reddit r/Accounting, r/Audit, r/FraudPrevention …        │
  │ PwC · Deloitte · EY · KPMG · ACFE kamuya açık sayfalar   │
  └──────────────┬──────────────────────────────────────────┘
                 ▼ dis_veri/*.csv  (robots.txt kontrolü, hız sınırı, çevrimdışı yedek)
  ┌──────────────┴──────────┐     ┌────────────────────────────────────┐
  │ ERP verisi  veri/*.csv  │────▶│ PYTHON ANALİTİK (src/analiz.py)    │
  │ (sentetik + ekilmiş     │     │ Benford ailesi (χ², MAD, KS, toplam)│
  │  hile senaryoları ya da │     │ Isolation Forest (9 özellik)        │
  │  sizin ERP dökümünüz)   │     │ Varlık ağı (union-find)             │
  └──────────────┬──────────┘     │ Bulanık ünvan / yaptırım eşleşmesi  │
                 │                │ Kullanıcı saat dağılımı χ²          │
                 │                │ Lapping zinciri · makbuz boşlukları │
                 ▼                └──────────────┬─────────────────────┘
  ┌────────────────────────────────────────────────────────────────────┐
  │ EXCEL MODELİ (src/excel_model.py)                                   │
  │ Parametreler → T_Odeme(26 test) · T_Tahsilat(8) · T_Org(12) ·        │
  │ T_Personel(4) · Benford · Istatistik · Dinamik_365 → Risk_Panosu     │
  └──────────────────────────────┬─────────────────────────────────────┘
                                 ▼
                  src/dogrula.py → ekilmiş hilelerin yakalanma oranı
```

## 2. Çalıştırma

```bash
pip install -r requirements.txt
```

```bash
python src/calistir.py
```

İnternet erişimi yoksa yedek dış verilerle çalıştırın:

```bash
python src/calistir.py --cevrimdisi
```

Kendi ERP dökümünüzle çalıştırmak için `veri/*.csv` dosyalarını aynı sütun başlıklarıyla değiştirin. Bu modda sentetik veri üretimi ve doğrulama adımı atlanır:

```bash
python src/calistir.py --gercek-veri
```

Excel modeli yeniden hesaplanabilir: veri sayfalarındaki satırlar değiştiğinde ya da `Parametreler` sayfasındaki eşik ve ağırlıklar güncellendiğinde bayraklar, skorlar ve pano kendiliğinden yenilenir. Satır sayısı değişirse adlandırılmış aralıkların yeniden kurulması için `excel_model.py` adımını tekrar çalıştırın.

## 3. Çalışma kitabı

| Sayfa | İçerik |
|---|---|
| **Kilavuz** | Amaç, akış, renk kodu, kaynakça |
| **Risk_Panosu** | KPI'lar, test isabetleri ve grafikleri, en riskli 20 ödeme, 10 tedarikçi ve 10 tahsilat, her kayıtta tetiklenen testler (TEXTJOIN) |
| **Parametreler** | Tüm eşikler ve test ağırlıkları (sarı hücreler). Formüller bu hücrelere adla bağlıdır |
| **T_Odeme** | Her ödeme için T01–T26 testleri. Risk skoru = `SUMPRODUCT(bayraklar; ağırlıklar)` |
| **T_Tahsilat** | H01–H08: lapping, nakdin geç yatırılması, eşik altı nakit, makbuz boşluğu … |
| **T_Org** | G01–G12: VKN/IBAN doğrulama, tedarikçi–personel eşleşmesi, yaptırım, ardışık fatura, Benford χ², RSF |
| **T_Personel** | P01–P04: TCKN algoritması, ortak IBAN (hayalet personel), tedarikçi bağı |
| **Benford** | İlk hane, ilk iki hane ve son iki hane testleri; Z, χ², MAD, Nigrini sınıfı, grafikler |
| **Istatistik** | SUBTOTAL, AGGREGATE persentilleri, IQR, FREQUENCY histogramı, aylık FORECAST.LINEAR sapması, CORREL, gün ve saat örüntüsü, onaylayan analizi |
| **Dinamik_365** | Görseldeki 18 fonksiyonun Microsoft 365 dinamik dizi uygulamaları |
| **Python_Bulgular / Py_IsolationForest** | Python ileri analitik sonuçları |
| **Test_Kutuphanesi** | 59 testin ACFE şeması, mantığı, fonksiyonu, referansı, isabeti ve Reddit/Big Four topluluk sinyali |
| Veri sayfaları | Odeme, Tahsilat, Tedarikci, Personel, Kullanici · Kur, Tatil, Yaptirim, Topluluk, Kazima_Log |

Renk kodu: **mavi** sabit girdi, **siyah** formül, **yeşil** başka sayfaya bağlantı, **sarı zemin** değiştirilebilir varsayım, **kırmızı zemin** tetiklenen bayrak.

## 4. Testler

### Ödeme (T01–T26)
| Grup | Testler |
|---|---|
| Mükerrer | T01 tam eşleşme (COUNTIFS) · T02 bulanık fatura no (SUBSTITUTE ile normalize edilir, baştaki sıfırlar atılır) · T21 aynı tutarın tekrarı |
| Zamanlama | T03 hafta sonu · T04 resmi tatil (kazınan liste) · T05 mesai dışı |
| Yetki / organizasyon | T07 limit altı · T08 bölünmüş ödeme (±N gün, COUNTIFS+SUMIFS) · T09 oluşturan = onaylayan · T10 yetki aşımı · T16 kayıt öncesi ödeme · T17 pasif tedarikçi · T18 riskli tedarikçi |
| Hesaplama | T11 Miktar×BF ≠ Matrah · T12 KDV · T13 brüt · T14 kur sapması (TCMB kuru, `MATCH(…;1)`) · T15 fatura > ödeme tarihi · T26 kümülatif sözleşme aşımı |
| İstatistik / gizli örüntü | T19 **leave-one-out (studentize) log-Z** · T20 IQR uç çit · T22 yeni tedarikçiye yüksek ödeme · T23 ardışık fatura no (kabuk şirket) · T24 tedarikçi Benford χ² · T25 RSF (`AGGREGATE(14;6;…)`) |

### Tahsilat (H01–H08)
Lapping, nakdin geç yatırılması (skimming), raporlama eşiğinin hemen altında nakit (structuring), makbuz numarası boşluğu, hafta sonu/tatil, fazla tahsilat, mükerrer kayıt, yuvarlak nakit.

### Organizasyon doğrulama (G01–G12, P01–P04)
- **VKN** (GİB algoritması) tek bir `SUMPRODUCT` dizi formülüyle doğrulanır.
- **TCKN** 10. ve 11. hane kontrolüyle doğrulanır.
- **IBAN** ISO 13616 mod-97 ile, 28 haneli sayı zincirleme `MOD` kullanılarak doğrulanır.
- Tedarikçi ile personel arasında IBAN, adres ve telefon eşleşmesi aranır.
- Diğer testler: mükerrer VKN, ortak IBAN, yaptırım listesi eşleşmesi, tedarikçiyi açan kullanıcının ödemeyi de onaylaması, yuvarlak tutar oranı.

### İleri istatistik (Python)
- **Benford** ailesi: ilk hane, ikinci hane, ilk iki hane ve son iki hane testleri; χ², MAD (Nigrini) ve Kolmogorov-Smirnov ile ölçülür. Ayrıca toplam (summation) testi yapılır.
- **Isolation Forest**: 9 özellik kullanılır (log tutar, saat, gün, kayıttan geçen süre, Z, yuvarlaklık, tekrar, fatura→ödeme süresi, limit oranı).
- **Varlık ağı**: ortak IBAN, adres, telefon ve VKN üzerinden union-find ile kümeler kurulur.
- **Bulanık eşleşme**: ünvan benzerliği ve yaptırım listesi taraması.
- **Kullanıcı davranışı**: onaylayan bazında saat dağılımı χ² testi.
- **Lapping zinciri**: yönlü graf üzerinde tespit edilir.

## 5. Görseldeki 18 fonksiyonun kullanıldığı yerler

| Fonksiyon | Kullanım |
|---|---|
| TAKE | En riskli 15 ödeme (`TAKE(SORTBY(HSTACK(…)))`), FORECAST.LINEAR'da önceki aylar, TCKN_GECERLI |
| DROP | 16.–30. sıradaki ödemeler (ikinci öncelik dilimi) |
| VSTACK | Tedarikçi ve personel IBAN/telefon/adreslerini birleştirip ortak olanları bulma |
| HSTACK | Rapor kolonlarını kurma, riskli tedarikçi kartı, SCAN çıktısı |
| TOCOL | Telefon ve adresleri tek sütuna indirip tekrarları sayma |
| TOROW | Bayraklı tedarikçi kimlikleri tek satırda |
| SORTBY | Risk sıralamaları (çok anahtarlı) |
| GROUPBY | Onaylayan bazında TL, SoD, yetki aşımı ve limit altı adedi |
| PIVOTBY | Onaylayan × ay risk ısı haritası |
| SUBTOTAL | `SUBTOTAL(9)` ve filtreye duyarlı `SUBTOTAL(109)` |
| AGGREGATE | Koşullu en büyük / 2. en büyük (RSF), min/maks fatura no, P95/P99 |
| FREQUENCY | Tutar histogramı |
| FORECAST.LINEAR | Aylık tahmin sapması testi; MAP ile her ay yalnızca önceki aylardan tahmin |
| CORREL | Ödeme–tahsilat ve adet–tutar korelasyonu |
| LAMBDA | Ad Yöneticisi'nde `VKN_GECERLI`, `TCKN_GECERLI`, `IBAN_GECERLI` (REDUCE ile mod-97), `ZSKOR`, `BENFORD_P` |
| MAP | Tüm tedarikçilerde VKN/IBAN doğrulama, tekrar sayımı |
| BYROW | Tedarikçi Benford MAD'i, risk skoru mutabakatı (0 fark beklenir) |
| SCAN | Seçili tedarikçiye kümülatif ödeme ve sözleşme limitinin aşıldığı an |

Ek olarak LET, FILTER, UNIQUE, XLOOKUP, CHOOSECOLS/CHOOSEROWS, SEQUENCE, REDUCE ve TEXTJOIN kullanılır. Klasik katman (Excel 2010+ ve LibreOffice) bu fonksiyonlara ihtiyaç duymaz. `Dinamik_365` sayfası Microsoft 365 gerektirir; GROUPBY ve PIVOTBY yalnızca Microsoft 365'te bulunur.

## 6. Web kazıma, Reddit ve Big Four

| Kaynak | Kullanım | Not |
|---|---|---|
| TCMB `kurlar/YYYYMM/DDMMYYYY.xml` | T14 kur testi | Döviz alış kuru; tatil ve hafta sonu için önceki iş günü kullanılır |
| Nager.Date API | T04 / H05 | |
| OFAC SDN CSV | G08 yaptırım taraması | Excel'de anahtar içerme, Python'da bulanık eşleşme |
| Reddit JSON araması | Test önceliklendirme | Yalnızca başlık, metin, puan ve bağlantı alınır; **kullanıcı adı toplanmaz**. Üretim ortamında Reddit OAuth API (praw) ve Data API şartları geçerlidir |
| PwC, Deloitte, EY, KPMG, ACFE sayfaları | Test önceliklendirme | URL'ler zamanla değişir; `dis_veri/kaynaklar.json` ile güncelleyin: `[["Kurum","URL"], …]` |

Reddit ve Big Four metinleri, `testler.py` içindeki TR/EN anahtar kelimelerle testlere eşlenir. `Test_Kutuphanesi` sayfasında her test için "Topluluk eşleşme" ve "Topluluk puanı" gösterilir; böylece uygulayıcıların en çok konuştuğu red flag'ler öne çıkar.

`robots.txt` kontrol edilir, istekler arasında bekleme yapılır ve art arda hata alınırsa kazıma erken durur. Erişilemeyen kaynakta `dis_veri/yedek/` kullanılır ve durum `Kazima_Log` sayfasına yazılır.

> ⚠️ Bu depodaki dosyalar ağ erişimi kısıtlı bir ortamda üretildi, bu yüzden tüm kaynaklar **YEDEK** modda çalıştı. Yedek kur serisi **sentetik**, yedek yaptırım listesi **kurgusal**dır ve topluluk verisi boştur (uydurma topluluk verisi yazılmaz). Gerçek denetimde `python src/calistir.py` komutunu internet erişimi olan bir makinede çalıştırın.

## 7. Doğrulama sonuçları (sentetik veri, ekilmiş senaryolar)

- **46 senaryo tipinin 45'inde %100 yakalama.** RSF (T25) %80 yakaladı: bir aykırı değer, tedarikçinin doğal büyük ödemelerinin arasında kaldı.
- **Yüksek risk kesinliği %90:** "Yüksek" seviyedeki 71 ödemenin 64'ü ekilmiş senaryodur. Orta seviyede kesinlik %68.
- VKN, TCKN ve IBAN Excel formülleri Python algoritmalarıyla **0 uyuşmazlık** verdi.
- Isolation Forest'ın işaretlediği aykırıların %71'i ekilmiş senaryodur. Kural tabanlı testleri tamamlayıcı olarak kullanılır.

Bilinçli yanlış pozitif örnekleri de vardır:
- Her ay sabit ödenen kira (V030), T21 ve G12 bayraklarını alır.
- Limit altına doğal olarak düşen ödemeler T07 bayrağını alır.

Bunlar, bayrağın kendisinin değil, denetçi incelemesinin karar verdiğini göstermek içindir.

## 8. Yöntem notları ve sınırlar

- Bayraklar **inceleme önceliği**dir, hile kanıtı değildir (ISA 240). Ağırlıklar şu mantıkla belirlendi: kesin istisnalar (hesaplama hataları, mükerrer, yetki aşımı) ≥ 4, zayıf sinyaller (hafta sonu, yuvarlak tutar) 1–2.
- Nakit raporlama eşiği (30.000 TL) ve onay limiti (50.000 TL) **varsayımdır**. Kurum politikasına ve güncel mevzuata göre güncelleyin.
- Tedarikçi bazlı Benford testi küçük örneklerde güvenilmezdir. Bu yüzden n ≥ 30 ve χ² p < 0,001 şartı aranır.
- Tüm kişi ve şirket adları kurgusaldır.

## 9. Kaynakça
ACFE *Report to the Nations* ve *Fraud Tree* · Nigrini, M. (2012) *Benford's Law*, Wiley · IIA GTAG 13 & 16 · ISA 240, 500, 520 · COSO 2013 · Liu, Ting & Zhou (2008) *Isolation Forest* · GİB VKN ve NVİ TCKN algoritmaları · ISO 13616 (IBAN).

## Dizin yapısı

```
src/
  calistir.py        uçtan uca boru hattı
  kazima.py          web kazıma (TCMB, tatil, OFAC, Reddit, Big Four/ACFE)
  veri_uret.py       sentetik veri + ekilmiş hile senaryoları (ground truth)
  analiz.py          Python ileri istatistik / ML
  excel_model.py     Excel modeli (XlsxWriter, 90 bin formül)
  dogrula.py         recall / precision / algoritma mutabakatı
  yeniden_hesapla.py LibreOffice ile kopya üzerinde yeniden hesaplama
  testler.py         test kütüphanesi (tek kaynak)
  ortak.py           VKN/TCKN/IBAN algoritmaları, normalizasyon
veri/                ERP verisi (CSV) + gercek_anomaliler.csv
dis_veri/            kazınan dış veri + yedek/
cikti/               Hile_Denetim_Modeli.xlsx, dogrulama_raporu.md, py_*.csv
```
