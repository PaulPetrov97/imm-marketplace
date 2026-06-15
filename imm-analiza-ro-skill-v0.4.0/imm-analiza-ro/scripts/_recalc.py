# -*- coding: utf-8 -*-
"""Best-effort: recalculează un workbook Excel (.xlsx) prin Excel COM pe Windows,
ca să scrie ÎN fișier valorile calculate ale formulelor (cache-ul `<v>`).

De ce: openpyxl scrie formule fără valoare-cache, deci în viewere care NU
recalculează (preview Windows, export PDF headless, re-citire openpyxl, unele
mobile) celulele cu formule apar GOALE. Excel desktop recalculează oricum la
deschidere (fullCalcOnLoad), dar pentru livrabile robuste pre-calculăm aici.

No-op tăcut dacă nu e Windows sau dacă Excel/COM lipsește — fișierul rămâne
valid și se recalculează la prima deschidere în Excel.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def recalc_xlsx(path, timeout: int = 120) -> bool:
    """Deschide workbook-ul în Excel (invizibil), forțează recalcul complet și
    salvează. Întoarce True dacă a reușit, False altfel (fără excepție)."""
    p = Path(path).resolve()
    if not sys.platform.startswith("win") or not p.exists():
        return False
    ps = (
        '$ErrorActionPreference="Stop";'
        "try{"
        "$x=New-Object -ComObject Excel.Application;"
        "$x.Visible=$false;$x.DisplayAlerts=$false;"
        f'$wb=$x.Workbooks.Open("{p}");'
        "$x.CalculateFullRebuild();"
        "$wb.Save();$wb.Close($false);$x.Quit();"
        "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($x)|Out-Null;"
        "exit 0"
        "}catch{exit 3}"
    )
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True,
            timeout=timeout,
        )
        return r.returncode == 0
    except Exception:
        return False


if __name__ == "__main__":
    if len(sys.argv) > 1:
        ok = recalc_xlsx(sys.argv[1])
        print("recalc OK" if ok else "recalc skipped/failed (no-op)")
