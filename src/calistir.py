"""Uçtan uca boru hattı.

    python src/calistir.py                 # canlı kazıma + sentetik veri + analiz + Excel + doğrulama
    python src/calistir.py --cevrimdisi    # internet olmadan (yedek dış veri)
    python src/calistir.py --gercek-veri   # veri/*.csv sizin ERP dökümünüz: sentetik üretim ve doğrulama atlanır
"""
from __future__ import annotations

import argparse
import runpy
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent
sys.path.insert(0, str(SRC))


def adim(baslik, modul, argv=None):
    print(f"\n=== {baslik} ===")
    eski = sys.argv
    sys.argv = [modul] + (argv or [])
    try:
        runpy.run_path(str(SRC / modul), run_name="__main__")
    finally:
        sys.argv = eski


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--yil", default="2025")
    ap.add_argument("--cevrimdisi", action="store_true")
    ap.add_argument("--gercek-veri", action="store_true")
    a = ap.parse_args()

    adim("1/5 Web kazıma", "kazima.py", ["--yil", a.yil] + (["--cevrimdisi"] if a.cevrimdisi else []))
    if not a.gercek_veri:
        adim("2/5 Sentetik veri + ekilen hile senaryoları", "veri_uret.py")
    adim("3/5 Python ileri analitik", "analiz.py")
    adim("4/5 Excel modeli", "excel_model.py")
    if not a.gercek_veri:
        print("\n=== 5/5 Doğrulama ===")
        from ortak import CIKTI
        from yeniden_hesapla import hesapla
        import dogrula
        kopya = hesapla(CIKTI / "Hile_Denetim_Modeli.xlsx")
        if kopya:
            dogrula.DOSYA = kopya
            dogrula.calistir()


if __name__ == "__main__":
    main()
