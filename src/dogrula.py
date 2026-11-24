"""Model doğrulama: Excel formüllerinin sonuçlarını ekilen (bilinen) hilelerle karşılaştırır.

1. Her ekilen anomali tipi için ilgili test sütununun yakalama oranı (recall)
2. Ekilen kayıtların risk skoru ile "Orta/Yüksek" önceliğe girme oranı
3. Kesinlik göstergesi: bayraklı kayıtların ne kadarı ekilen bir senaryoya ait (precision)
4. Excel VKN / TCKN / IBAN formüllerinin Python algoritmalarıyla birebir mutabakatı
5. Isolation Forest'ın ekilen ödemeleri yakalama oranı

Önkoşul: çalışma kitabı LibreOffice/Excel ile yeniden hesaplanmış olmalı (calistir.py bunu yapar).
Çıktı: cikti/dogrulama_raporu.md
"""
from __future__ import annotations

import warnings

import openpyxl
import pandas as pd

from ortak import CIKTI, VERI, iban_gecerli, tckn_gecerli, vkn_gecerli

warnings.filterwarnings("ignore")
DOSYA = CIKTI / "Hile_Denetim_Modeli.xlsx"


def sayfa_df(wb, ad):
    satirlar = list(wb[ad].iter_rows(values_only=True))
    return pd.DataFrame(satirlar[1:], columns=satirlar[0])


def calistir():
    wb = openpyxl.load_workbook(DOSYA, data_only=True, read_only=True)
    to, th = sayfa_df(wb, "T_Odeme"), sayfa_df(wb, "T_Tahsilat")
    tg, tp = sayfa_df(wb, "T_Org"), sayfa_df(wb, "T_Personel")
    if to["Risk Skoru"].isna().all():
        raise SystemExit("Çalışma kitabı hesaplanmamış: önce recalc çalıştırın.")
    gt = pd.read_csv(VERI / "gercek_anomaliler.csv")
    to = to.set_index("Ödeme ID")
    th = th.set_index("Tahsilat ID")
    tg = tg.set_index("Tedarikçi ID")
    tp = tp.set_index("Personel ID")

    def kolon(df, kod):
        return next(c for c in df.columns if str(c).split(" ")[0] == kod)

    satir = []
    for (tablo, tip), g in gt.groupby(["Tablo", "Tip"]):
        df = {"Odeme": to, "Tahsilat": th, "Tedarikci": tg, "Personel": tp}[tablo]
        k = kolon(df, tip)
        bayrak = df.loc[g["ID"], k].fillna(0).astype(int)
        oncelik = ""
        if tablo in ("Odeme", "Tahsilat"):
            oncelik = f"{(df.loc[g['ID'], 'Risk Seviyesi'] != 'Düşük').mean():.0%}"
        satir.append(dict(Tablo=tablo, Test=tip, Ekilen=len(g), Yakalanan=int(bayrak.sum()),
                          Recall=f"{bayrak.mean():.0%}", OrtaYuksekOncelik=oncelik,
                          ToplamBayrak=int(df[k].fillna(0).astype(int).sum())))
    rec = pd.DataFrame(satir)

    # kesinlik: ekilen herhangi bir senaryoya ait olan bayraklı kayıt oranı
    kes = []
    for tablo, df in (("Odeme", to), ("Tahsilat", th)):
        ekilen = set(gt.loc[gt["Tablo"] == tablo, "ID"])
        for sev in ("Yüksek", "Orta"):
            secili = df.index[df["Risk Seviyesi"] == sev]
            kes.append(dict(Tablo=tablo, Seviye=sev, Kayit=len(secili),
                            EkilenSenaryo=sum(i in ekilen for i in secili),
                            Kesinlik=f"{(sum(i in ekilen for i in secili) / max(len(secili), 1)):.0%}"))
        kes.append(dict(Tablo=tablo, Seviye="Ekilen toplam", Kayit=len(ekilen),
                        EkilenSenaryo=sum(df.loc[list(ekilen), "Risk Seviyesi"] != "Düşük"),
                        Kesinlik="(Orta+Yüksek'e giren ekilen kayıt)"))
    kes = pd.DataFrame(kes)

    # algoritma mutabakatı
    td = pd.read_csv(VERI / "tedarikciler.csv", dtype=str).set_index("TedarikciID")
    pr = pd.read_csv(VERI / "personel.csv", dtype=str).set_index("PersonelID")
    mut = [
        ("VKN (G01)", sum((tg.loc[i, kolon(tg, "G01")] == 1) != (not vkn_gecerli(td.loc[i, "VKN"])) for i in td.index), len(td)),
        ("IBAN tedarikçi (G03)", sum((tg.loc[i, kolon(tg, "G03")] == 1) != (not iban_gecerli(td.loc[i, "IBAN"])) for i in td.index), len(td)),
        ("TCKN (P01)", sum((tp.loc[i, kolon(tp, "P01")] == 1) != (not tckn_gecerli(pr.loc[i, "TCKN"])) for i in pr.index), len(pr)),
        ("IBAN personel (P03)", sum((tp.loc[i, kolon(tp, "P03")] == 1) != (not iban_gecerli(pr.loc[i, "IBAN"])) for i in pr.index), len(pr)),
    ]
    mut = pd.DataFrame(mut, columns=["Kontrol", "Uyuşmazlık", "Kayıt"])

    ekilen_od = set(gt.loc[gt["Tablo"] == "Odeme", "ID"])
    if_ayk = to["Python IF Aykırı"].fillna(0).astype(int)
    if_recall = if_ayk.loc[list(ekilen_od)].mean()
    if_kes = (if_ayk[if_ayk == 1].index.isin(list(ekilen_od))).mean()

    rapor = ["# Model Doğrulama Raporu", "",
             "Ekilen (bilinen) hile senaryolarının Excel formül katmanı tarafından yakalanma durumu. "
             "Değerler yeniden hesaplanmış çalışma kitabından okunmuştur.", "",
             "## 1. Test bazında yakalama oranı (recall)", "", rec.to_markdown(index=False), "",
             "*ToplamBayrak*: testin tüm veri setinde işaretlediği kayıt sayısı (ekilen + doğal/yanlış pozitif).", "",
             "## 2. Önceliklendirme kesinliği", "", kes.to_markdown(index=False), "",
             "## 3. Doğrulama algoritmaları – Excel formülü vs Python", "", mut.to_markdown(index=False), "",
             "## 4. Isolation Forest (denetimsiz)", "",
             f"- Ekilen ödemelerin IF tarafından aykırı bulunma oranı: **{if_recall:.0%}**",
             f"- IF aykırı kayıtların ekilen senaryoya ait olma oranı: **{if_kes:.0%}**", "",
             "Isolation Forest etiket görmeden çalışır; kural tabanlı testlerin tanımlamadığı örüntüleri yakalamak için "
             "tamamlayıcıdır, tek başına karar ölçütü değildir.", ""]
    (CIKTI / "dogrulama_raporu.md").write_text("\n".join(rapor), encoding="utf-8")
    print("\n".join(rapor))


if __name__ == "__main__":
    calistir()
