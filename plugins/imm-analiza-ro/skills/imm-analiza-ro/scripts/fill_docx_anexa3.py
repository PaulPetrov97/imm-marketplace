"""Completare Anexa 3 — Declarație privind încadrarea întreprinderii în categoria IMM.

Pași:
1. shutil.copy(template, out) — copie nouă.
2. Înlocuiește placeholder-urile cu underscore în paragrafele de identificare.
3. Bifează EXACT una din 3 căsuțe (autonomă / parteneră / legată) prin
   înlocuirea Wingdings unchecked (U+F0A8 sau U+F0F8) cu CHECKED (U+F0FE).
4. Completează tabelul financiar (Tabel 2).
5. Adaugă data semnării.
"""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("EROARE: python-docx lipsă. Rulează: pip install python-docx", file=sys.stderr)
    sys.exit(2)


# Caractere Wingdings pentru checkbox (private use area)
# Template-ul folosește un caracter unchecked pe care îl detectăm dinamic la rulare;
# înlocuim cu U+2611 (BALLOT BOX WITH CHECK) sau U+F0FE (Wingdings) după font.
UNCHECKED_CANDIDATES = ["", "", "☐", "☐"]
CHECKED_WINGDINGS = ""
CHECKED_UNICODE = "☑"


TIP_TO_ROW_INDEX = {
    "AUTONOMA": 0,
    "PARTENERA": 1,
    "LEGATA": 2,
}


def _replace_underscore_placeholder(paragraph, label: str, value: str) -> bool:
    """Înlocuiește runul cu '___' dintr-un paragraf care începe cu `label`."""
    if not paragraph.text.startswith(label):
        return False
    for run in paragraph.runs:
        if "___" in run.text:
            # Păstrează un spațiu între etichetă și valoare
            run.text = " " + value
            return True
    return False


def _tick_checkbox(cell, tip: str) -> bool:
    """Bifează căsuța dintr-o celulă — păstrează font Wingdings.

    Strategie: găsește runul care conține caracterul UNCHECKED și înlocuiește
    cu CHECKED_WINGDINGS dacă fontul e Wingdings, altfel cu CHECKED_UNICODE.
    """
    for para in cell.paragraphs:
        for run in para.runs:
            for unchecked in UNCHECKED_CANDIDATES:
                if unchecked in run.text:
                    font_name = (run.font.name or "").lower()
                    is_wingdings = "wing" in font_name
                    replacement = CHECKED_WINGDINGS if is_wingdings else CHECKED_UNICODE
                    run.text = run.text.replace(unchecked, replacement, 1)
                    return True
    return False


def fill_anexa3(
    template_path: Path,
    out_path: Path,
    solicitant: dict,
    totals_ref: dict,
    totals_prev: dict | None,
    tip: str,
    modif_categorie: bool = False,
) -> Path:
    """
    solicitant: {denumire, adresa, cui, reprezentant_nume, reprezentant_functie}
    totals_ref: {salariati, cifra_afaceri_mii_lei, active_totale_mii_lei, an}
    totals_prev: idem, pentru exercițiul anterior (opțional)
    tip: "AUTONOMA" | "PARTENERA" | "LEGATA"
    modif_categorie: True dacă datele financiare au modificat categoria față de N-1
    """
    if tip not in TIP_TO_ROW_INDEX:
        raise ValueError(f"Tip invalid: {tip}. Trebuie AUTONOMA/PARTENERA/LEGATA")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(template_path, out_path)
    doc = Document(out_path)

    # ---- 1. Identificare în paragrafe ----
    placeholders = {
        "Denumirea întreprinderii": solicitant.get("denumire", ""),
        "Adresa sediului social":   solicitant.get("adresa", ""),
        "Cod unic de înregistrare": solicitant.get("cui", ""),
        "Numele și funcția":        f"{solicitant.get('reprezentant_nume','')}, "
                                    f"{solicitant.get('reprezentant_functie','')}",
    }
    for p in doc.paragraphs:
        for label, value in placeholders.items():
            if _replace_underscore_placeholder(p, label, value):
                break

    # ---- 2. Tabel 0: bifa pe rândul corect ----
    if len(doc.tables) < 1:
        raise RuntimeError("Anexa 3 — tabelul de tip întreprindere lipsește")
    t0 = doc.tables[0]
    target_idx = TIP_TO_ROW_INDEX[tip]
    if len(t0.rows) < 3:
        raise RuntimeError(f"Anexa 3 — tabel tip are {len(t0.rows)} rânduri, aștept 3")
    for ri in range(3):
        cell = t0.rows[ri].cells[0]
        if ri == target_idx:
            ok = _tick_checkbox(cell, tip)
            if not ok:
                # Fallback: scrie un X în prima celulă (vizibil chiar dacă font-ul nu randează)
                cell.paragraphs[0].add_run(" X").bold = True

    # ---- 3. Tabel 1: date financiare ----
    if len(doc.tables) < 2:
        raise RuntimeError("Anexa 3 — tabelul financiar lipsește")
    t1 = doc.tables[1]
    if len(t1.rows) < 3:
        raise RuntimeError(f"Anexa 3 — tabel financiar are {len(t1.rows)} rânduri")

    # Antet an de referință (row 0)
    if "an" in totals_ref:
        t1.rows[0].cells[0].text = f"Exercițiul financiar de referință — anul {totals_ref['an']}"

    # Row 2 = primul exercițiu (referință)
    _set_cell(t1.rows[2].cells[0], f"{totals_ref.get('salariati', 0):g}")
    _set_cell(t1.rows[2].cells[1], f"{totals_ref.get('cifra_afaceri_mii_lei', 0):.3f}")
    _set_cell(t1.rows[2].cells[2], f"{totals_ref.get('active_totale_mii_lei', 0):.3f}")

    # Row 3 = exercițiul anterior (dacă există)
    if totals_prev and len(t1.rows) >= 4:
        _set_cell(t1.rows[3].cells[0], f"{totals_prev.get('salariati', 0):g}")
        _set_cell(t1.rows[3].cells[1], f"{totals_prev.get('cifra_afaceri_mii_lei', 0):.3f}")
        _set_cell(t1.rows[3].cells[2], f"{totals_prev.get('active_totale_mii_lei', 0):.3f}")

    # ---- 4. Tabel 2: modificare categorie (Da/Nu) ----
    if len(doc.tables) >= 3:
        t2 = doc.tables[2]
        if len(t2.rows) >= 1 and len(t2.rows[0].cells) >= 2:
            # Adaugă marcajul în celula 2
            mark = "X DA" if modif_categorie else "X NU"
            t2.rows[0].cells[1].paragraphs[0].add_run(f"\n[{mark}]").bold = True

    # ---- 5. Data semnării ----
    today_str = date.today().strftime("%d.%m.%Y")
    for p in doc.paragraphs:
        if "Data semnării" in p.text:
            for run in p.runs:
                if "……" in run.text or "..." in run.text:
                    if "……" in run.text:
                        run.text = run.text.replace("……", " " + today_str, 1)
                    else:
                        run.text = run.text.replace("...", " " + today_str, 1)
                    break
            break

    doc.save(out_path)
    return out_path


def _set_cell(cell, text: str) -> None:
    """Setează textul unei celule fără a strica formatarea — clear + add run."""
    # Curăță paragrafele existente
    for p in cell.paragraphs:
        for run in p.runs:
            run.text = ""
    # Pune textul în primul paragraf (sau adaugă unul)
    if cell.paragraphs:
        cell.paragraphs[0].add_run(text)
    else:
        cell.add_paragraph(text)


# -------------------- CLI smoke test --------------------

def _smoke():
    sys.path.insert(0, str(Path(__file__).parent))
    from _paths import templates_dir
    tpl = templates_dir() / "Anexa-3-Declaratie-incadrare-IMM.docx"
    out = Path.cwd() / f"_smoke_Anexa3_{date.today().isoformat()}.docx"
    solicitant = {
        "denumire": "ROMACTIV BUSINESS CONSULTING SRL",
        "adresa": "Str. Constantin Tănase nr. 12, București, Sector 2",
        "cui": "RO14186770",
        "reprezentant_nume": "POPESCU PAUL",
        "reprezentant_functie": "Administrator",
    }
    totals_ref = {
        "an": 2024,
        "salariati": 8,
        "cifra_afaceri_mii_lei": 1500.000,
        "active_totale_mii_lei": 850.000,
    }
    totals_prev = {
        "salariati": 7,
        "cifra_afaceri_mii_lei": 1200.000,
        "active_totale_mii_lei": 700.000,
    }
    result = fill_anexa3(tpl, out, solicitant, totals_ref, totals_prev, tip="PARTENERA",
                        modif_categorie=False)
    print(f"Smoke OK: {result}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--smoke", action="store_true")
    args = ap.parse_args()
    if args.smoke:
        _smoke()
