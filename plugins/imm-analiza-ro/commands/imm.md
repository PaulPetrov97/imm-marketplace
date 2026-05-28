---
description: Analiză preliminară de încadrare IMM pentru una sau mai multe firme identificate prin CUI. Primul CUI = solicitant. Folosește termene.ro (Chrome MCP) + arborele decizional al legăturilor + completează automat Declarație IMM + Anexa 3 + Anexa 4 + produce sinteză cu verdict.
argument-hint: <CUI1> [CUI2 CUI3 ...] [--an YYYY] [--output <path>]
---

# /imm — analiză preliminară IMM

Invocă skill-ul `imm-analiza-ro` pentru o analiză completă end-to-end de încadrare în categoria IMM (autonomă / parteneră / legată) conform Legii 346/2004 + Recomandării CE 2003/361/EC.

## Argumente

- `CUI1` (obligatoriu) — CUI-ul solicitantului (firma care aplică la finanțare).
- `CUI2 CUI3 ...` (opțional) — CUI-uri suspectate ca parteneri/legate. Dacă nu sunt furnizate, skill-ul le descoperă automat din scrape-ul asociaților.
- `--an YYYY` (opțional) — anul de referință pentru date financiare. Default: anul curent − 1.
- `--output <path>` (opțional) — folder pentru output. Default: cwd / `analize-imm/<YYYY-MM-DD>_<CUI1>/`.

## Exemple

```
/imm 14186770
/imm 14186770 32165478 25896374
/imm 14186770 --an 2024
/imm 14186770 --output C:\analize\PNRR-2026
```

## Workflow declanșat

1. Preflight (Chrome MCP + Python deps).
2. Confirmare scope (CUI, an, folder).
3. Scrape solicitant + recursie controlată pe asociați.
4. Decision engine pe arborele legăturilor.
5. AskUserQuestion pentru cazuri ambigue (rude, piață învecinată, investitori excepție, deținere publică).
6. Completare automată: Excel (Sheet 2 + Sheet 4), Anexa 3 (bifa corectă), Anexa 4 (Secțiunea A + B + fișe duplicate).
7. Sinteză 1–3 pagini cu verdict "Există legături: DA/NU/NECLAR".
8. Audit log per analiză.

## Output

Toate fișierele apar în `<output>/analize-imm/<data>_<CUI>/`:

- `01_raw/` — JSON-uri brute de pe termene.ro (audit + reproducibilitate)
- `02_Sinteza_<CUI>.md` și `.docx` — sinteză cu verdict
- `03_Declaratie_IMM_<CUI>_completata.xlsx`
- `04_Anexa-3-completata.docx`
- `05_Anexa-4-completata.docx`
- `audit.log`

## Instrucțiune pentru Claude

Invocă skill-ul `imm-analiza-ro` cu argumentele furnizate de utilizator. Urmează workflow-ul din `SKILL.md` în ordine strictă. NU sări peste preflight. NU livra rezultate parțiale.

Args: $ARGUMENTS
