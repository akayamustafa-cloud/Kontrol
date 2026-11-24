# Model Doğrulama Raporu

Ekilen (bilinen) hile senaryolarının Excel formül katmanı tarafından yakalanma durumu. Değerler yeniden hesaplanmış çalışma kitabından okunmuştur.

## 1. Test bazında yakalama oranı (recall)

| Tablo     | Test   |   Ekilen |   Yakalanan | Recall   | OrtaYuksekOncelik   |   ToplamBayrak |
|:----------|:-------|---------:|------------:|:---------|:--------------------|---------------:|
| Odeme     | T01    |       12 |          12 | 100%     | 100%                |             12 |
| Odeme     | T02    |        8 |           8 | 100%     | 100%                |              8 |
| Odeme     | T03    |        5 |           5 | 100%     | 20%                 |              5 |
| Odeme     | T04    |        4 |           4 | 100%     | 0%                  |              5 |
| Odeme     | T05    |        8 |           8 | 100%     | 0%                  |              8 |
| Odeme     | T06    |       37 |          37 | 100%     | 89%                 |             49 |
| Odeme     | T07    |       17 |          17 | 100%     | 100%                |             57 |
| Odeme     | T08    |       12 |          12 | 100%     | 100%                |             43 |
| Odeme     | T09    |       41 |          41 | 100%     | 100%                |             41 |
| Odeme     | T10    |        6 |           6 | 100%     | 100%                |              7 |
| Odeme     | T11    |        6 |           6 | 100%     | 100%                |              6 |
| Odeme     | T12    |        5 |           5 | 100%     | 100%                |              5 |
| Odeme     | T13    |        4 |           4 | 100%     | 100%                |              4 |
| Odeme     | T14    |        7 |           7 | 100%     | 100%                |              7 |
| Odeme     | T15    |        4 |           4 | 100%     | 100%                |              4 |
| Odeme     | T16    |        3 |           3 | 100%     | 100%                |              3 |
| Odeme     | T17    |        4 |           4 | 100%     | 100%                |              4 |
| Odeme     | T18    |       66 |          66 | 100%     | 100%                |             94 |
| Odeme     | T19    |        5 |           5 | 100%     | 100%                |             18 |
| Odeme     | T22    |        1 |           1 | 100%     | 100%                |              1 |
| Odeme     | T23    |       46 |          46 | 100%     | 100%                |             47 |
| Odeme     | T25    |        5 |           4 | 80%      | 100%                |              4 |
| Odeme     | T26    |       19 |          19 | 100%     | 100%                |             19 |
| Personel  | P01    |        2 |           2 | 100%     |                     |              2 |
| Personel  | P02    |        2 |           2 | 100%     |                     |              2 |
| Personel  | P03    |        1 |           1 | 100%     |                     |              1 |
| Tahsilat  | H01    |        6 |           6 | 100%     | 100%                |              6 |
| Tahsilat  | H02    |        6 |           6 | 100%     | 100%                |              6 |
| Tahsilat  | H03    |        6 |           6 | 100%     | 100%                |              6 |
| Tahsilat  | H04    |        4 |           4 | 100%     | 100%                |              4 |
| Tahsilat  | H05    |        3 |           3 | 100%     | 0%                  |              3 |
| Tahsilat  | H06    |        4 |           4 | 100%     | 100%                |              4 |
| Tahsilat  | H07    |        6 |           6 | 100%     | 100%                |              6 |
| Tahsilat  | H08    |        5 |           5 | 100%     | 0%                  |              5 |
| Tedarikci | G01    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G02    |        2 |           2 | 100%     |                     |              2 |
| Tedarikci | G03    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G04    |        2 |           2 | 100%     |                     |              2 |
| Tedarikci | G05    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G06    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G07    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G08    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G09    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G10    |        2 |           2 | 100%     |                     |              2 |
| Tedarikci | G11    |        1 |           1 | 100%     |                     |              1 |
| Tedarikci | G12    |        1 |           1 | 100%     |                     |              2 |

*ToplamBayrak*: testin tüm veri setinde işaretlediği kayıt sayısı (ekilen + doğal/yanlış pozitif).

## 2. Önceliklendirme kesinliği

| Tablo    | Seviye        |   Kayit |   EkilenSenaryo | Kesinlik                           |
|:---------|:--------------|--------:|----------------:|:-----------------------------------|
| Odeme    | Yüksek        |      71 |              64 | 90%                                |
| Odeme    | Orta          |     135 |              92 | 68%                                |
| Odeme    | Ekilen toplam |     176 |             156 | (Orta+Yüksek'e giren ekilen kayıt) |
| Tahsilat | Yüksek        |       0 |               0 | 0%                                 |
| Tahsilat | Orta          |      32 |              32 | 100%                               |
| Tahsilat | Ekilen toplam |      40 |              32 | (Orta+Yüksek'e giren ekilen kayıt) |

## 3. Doğrulama algoritmaları – Excel formülü vs Python

| Kontrol              |   Uyuşmazlık |   Kayıt |
|:---------------------|-------------:|--------:|
| VKN (G01)            |            0 |      67 |
| IBAN tedarikçi (G03) |            0 |      67 |
| TCKN (P01)           |            0 |      80 |
| IBAN personel (P03)  |            0 |      80 |

## 4. Isolation Forest (denetimsiz)

- Ekilen ödemelerin IF tarafından aykırı bulunma oranı: **23%**
- IF aykırı kayıtların ekilen senaryoya ait olma oranı: **71%**

Isolation Forest etiket görmeden çalışır; kural tabanlı testlerin tanımlamadığı örüntüleri yakalamak için tamamlayıcıdır, tek başına karar ölçütü değildir.
