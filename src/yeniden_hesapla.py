"""LibreOffice (headless) ile çalışma kitabının bir KOPYASINI yeniden hesaplar.

Orijinal dosyaya dokunulmaz (Excel'e özgü dinamik dizi meta verisi korunur); dogrula.py değerleri
bu kopyadan okur. LibreOffice kurulu değilse dosyayı Excel'de açıp kaydetmeniz yeterlidir.
"""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

MAKRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub Hesapla()
  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>"""


def hesapla(kaynak: Path, zaman_asimi: int = 600) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("LibreOffice bulunamadı – yeniden hesaplama atlandı (dosyayı Excel'de açın).")
        return None
    is_dizini = Path(tempfile.mkdtemp(prefix="kontrol-lo-"))
    profil = (is_dizini / "profil").as_uri()
    subprocess.run([soffice, "--headless", "--terminate_after_init", f"-env:UserInstallation={profil}"],
                   capture_output=True, timeout=120)
    makro_dizin = is_dizini / "profil" / "user" / "basic" / "Standard"
    makro_dizin.mkdir(parents=True, exist_ok=True)
    (makro_dizin / "Module1.xba").write_text(MAKRO, encoding="utf-8")
    kopya = is_dizini / kaynak.name
    shutil.copyfile(kaynak, kopya)
    subprocess.run([soffice, "--headless", "--norestore", f"-env:UserInstallation={profil}",
                    "vnd.sun.star.script:Standard.Module1.Hesapla?language=Basic&location=application", str(kopya)],
                   capture_output=True, timeout=zaman_asimi)
    return kopya
