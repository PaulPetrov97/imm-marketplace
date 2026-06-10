# Changelog

## v0.3.0 — 2026-06-10

Feedback din utilizare reală (analizele TOPGYN + FRESCOVERDE) — 6 îmbunătățiri:

### Adăugat / schimbat
- **Workbook „Analiza încadrare IMM" = livrabil OBLIGATORIU** la fiecare analiză (`06_Analiza_incadrare_IMM_<denumire>.xlsx`), acum cu:
  - Secțiunea nouă **„5. Recomandare Claude"** (fundal galben): constatări + clarificări rămase + varianta de declarare recomandată.
  - Coloane noi: **CA și Active totale în MII lei și MII euro** (formule live `/1000`), pe lângă lei și euro.
  - Antet cu **Program / apel de finanțare**.
- **Întrebări OBLIGATORII la începutul fiecărui task**: program/apel, an de referință + curs BNR, livrabilele dorite, semnatarul Anexei 3, folder output.
- **Verificare obligatorie a AMBELOR secțiuni termene.ro**: „Asociați/acționari" ȘI „Persoane autorizate" + conexiunile ambelor + „Lista firmelor în care X este acționar" — administratorii pot lega firme fără participație (invarianta 8 + reference/04).
- **Regulă curs valutar**: întotdeauna cursul BNR din **ULTIMA ZI a anului de referință** (31 dec); tabel actualizat în reference/06 (2022=4,9474; 2023=4,9746; 2024=4,9741; 2025=5,0985).
- **Separatori numerici românești** în toate output-urile text: zecimale cu virgulă, mii cu punct (`1.234.567,89`) — helper nou `fmt_ro()` în `sinteza.py` + `fill_analiza_imm.py`.
- **Anexa 3 — etichetele nu se mai șterg**: la blocul de semnătură, „Numele" și „Funcția" rămân, iar numele/funcția semnatarului (persoana care întocmește anexa — parametri noi `semnatar_nume`/`semnatar_functie`) se scriu LÂNGĂ etichetă. Înlocuire sigură pe bază de regex (`[_\.…]{3,}`) care păstrează eticheta chiar dacă stă în același run cu placeholder-ul.

## v0.2.0 — 2026-06-09

Adăugat formatul de output bogat "Analiza încadrare IMM" (învățat din practica RBC, faza de contractare).

### Adăugat
- `scripts/fill_analiza_imm.py` — generator workbook de analiză completă într-un singur sheet:
  - **Harta grupului**: per firmă — asociați, cote, administrator, relații comerciale (DA/NU), adresă, CAEN principal + preponderent.
  - **Verdict legături** colorat (LEGATĂ / PARTENERĂ / AUTONOMĂ / NECLAR) + listă firme + notă.
  - **Tabele financiare consolidate per an** (3-4 ani): CA și active totale în lei ȘI euro (formule live `=lei/curs`), TOTAL care exclude firmele de pe piețe neînvecinate / sub prag.
  - **Categoria IMM finală**.
  - Se construiește de la zero (openpyxl) → scalează la orice număr de firme și ani, spre deosebire de Declaratie_IMM fixă (max 8 parteneri).
- `reference/06-format-analiza-imm.md` — documentația formatului + API generator + constante curs BNR 31.dec (2022=4,9474; 2023=4,9746; 2024=4,9741; 2025=5,0985).
- SKILL.md: pas nou în workflow (output `06_Analiza_incadrare_IMM_<denumire>.xlsx`).

### Notă
- Output-ul de analiză NU bundle-ază date reale de client în zip-ul distribuit (confidențialitate); formatul e capturat de generator + reference doc.

## v0.1.0 — 2026-05-28

Versiune inițială a pluginului `imm-analiza-ro`.

### Adăugat
- Slash command `/imm <CUI1> [CUI2 ...] [--an YYYY] [--output <path>]`.
- Skill `imm-analiza-ro` cu trigger pe fraze românești pentru analiza IMM.
- Arbore decizional complet pentru clasificare autonomă/parteneră/legată conform Legii 346/2004 + Recomandării CE 2003/361/EC.
- Scripts pentru completare automată:
  - `fill_xlsx.py` — Declarație IMM (Sheet 2 Date_partenere + Sheet 4 Date_legate) păstrând formulele.
  - `fill_docx_anexa3.py` — Anexa 3 cu bifa corectă pe tipul de întreprindere.
  - `fill_docx_anexa4.py` — Anexa 4 cu calcul partenere (proporțional) + legate (integral).
  - `sinteza.py` — sinteză MD + DOCX cu tabele și verdict.
- Reference docs:
  - Manualul CE complet (60 pagini) — bilia regulilor.
  - Distilare arbore decizional cu pseudocod.
  - Selectori DOM termene.ro pentru Chrome MCP.
  - Taxonomia cazurilor neclare cu întrebări sample.
- Preflight check (Python deps + template integrity).
- Audit log per analiză.
- Suport distribuție duală: marketplace git + .zip backup.
