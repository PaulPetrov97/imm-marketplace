# Exemplu de analiză — solicitant cu doi parteneri

## Invocare

```
/imm 14186770 32165478 25896374 --an 2024
```

Sau în limbaj natural:

> Analizează RomActiv (CUI 14186770) și verifică legăturile cu firmele 32165478 și 25896374. An de referință: 2024.

## Ce face pluginul

1. **Preflight**:
   - Confirmă Chrome MCP conectat.
   - Verifică openpyxl + python-docx.
   - Probă termene.ro: detectat tier "premium" (utilizator autentificat în PartnerSCAN).

2. **Confirmare scope** (AskUserQuestion):
   - Solicitant: 14186770 ✓
   - Suspectate partenere/legate: 32165478, 25896374
   - An referință: 2024
   - Output: cwd/analize-imm/2026-05-28_14186770/

3. **Scrape solicitant** (14186770):
   - Denumire, adresă, CAEN 7022 (consultanță management), reprezentant POPESCU PAUL.
   - Financiare 2024: 8 salariați, 1.500.000 lei cifră afaceri, 850.000 lei active.
   - Asociați: POPESCU PAUL (PF, 100%).

4. **Recursie pe asociatul PF** POPESCU PAUL:
   - Pluginul navighează la profilul PF pe termene.ro.
   - Detectează: POPESCU PAUL este asociat unic în 2 alte firme: 32165478 (40%) și 25896374 (60%).
   - Scrape ambele:
     - 32165478 — CAEN 7022 (consultanță, ADJACENT) — 12 salariați, 4.5M lei CA, 3.2M active.
     - 25896374 — CAEN 9999 (alte servicii personale, NU adjacent) — 5 salariați, 800k CA, 500k active.

5. **Clarificări AskUserQuestion**:
   - "Firma 25896374 cu CAEN 9999 este pe piață învecinată cu 14186770 (CAEN 7022)?" → Utilizator răspunde "Nu".

6. **Decision engine**:
   - 32165478: asociat unic POPESCU PAUL ≥25% în solicitant ȘI ≥25% în această firmă, CAEN adjacent → LEGATĂ prin PF concertat.
   - 25896374: asociat unic POPESCU PAUL deține 60%, dar piață NU adjacent → NU se cuplează în clasificarea IMM (Art. 4⁴ alin. (2)).
   - Verdict: SOLICITANT este LEGATĂ cu 32165478.
   - Agregat: salariați 8 + 12 = 20, CA 1.500.000 + 4.500.000 = 6.000.000 lei (~€1.2M), active 850.000 + 3.200.000 = 4.050.000 lei (~€814k).
   - Categorie de mărime: **micro** (sub 10 salariați? NU, 20 — trece la "mică", cu CA și active sub praguri).

7. **Completare documente**:
   - Excel `Declaratie_IMM_v1.0.xlsx`:
     - Sheet `Date_legate`: rândul 4 = "FIRMA LEGATA 32165478 SRL", 12 salariați, 4500.000 mii lei CA, 3200.000 mii lei active.
   - Anexa 3:
     - Identificare: ROMACTIV BUSINESS CONSULTING SRL / RO14186770 / Str. Constantin Tănase 12 / POPESCU PAUL, Administrator.
     - Bifa: "Întreprindere legată" ✓
     - Tabel financiar: row 2 = 2024 (8, 1500.000, 850.000), row 3 = 2023.
   - Anexa 4:
     - Secțiunea B (legate): rândul 1 = 32165478 cu datele integrale.
     - Total B: 12 salariați, 4500 mii lei CA, 3200 mii lei active.

8. **Sinteza** (`02_Sinteza_RO14186770.md`):
   - Profil firmă (tabel).
   - Matrice asociați cu cross-reference firmele 32165478 + 25896374.
   - Concluzie: există legături prin asociatul PF POPESCU PAUL cu firma 32165478 (piață identică).
   - Verdict: **Există legături: DA**.
   - Clasificare: LEGATĂ, categorie de mărime **întreprindere mică**.

9. **Audit log**:
   ```
   [2026-05-28T10:32:11] [INFO] Scope: solicitant=14186770, candidați=[32165478, 25896374], an=2024
   [2026-05-28T10:32:12] [INFO] Preflight OK: Chrome MCP / openpyxl / python-docx
   [2026-05-28T10:32:15] [INFO] Scrape 14186770 → termene.ro/firma/14186770 (tier=premium)
   [2026-05-28T10:32:17] [INFO] PF POPESCU PAUL → 2 firme adiționale găsite
   [2026-05-28T10:32:20] [INFO] Scrape 32165478 (CAEN 7022) → adjacent CAEN solicitant
   [2026-05-28T10:32:22] [INFO] Scrape 25896374 (CAEN 9999) → CAEN diferit, întreb user
   [2026-05-28T10:32:30] [INFO] User: 25896374 NU pe piață învecinată
   [2026-05-28T10:32:31] [DECISION] LEGATA via PF concertat: 14186770 ↔ 32165478
   [2026-05-28T10:32:32] [DECISION] 25896374 exclus din clasificare (piață ≠)
   [2026-05-28T10:32:35] [INFO] Excel completat: 03_Declaratie_IMM_14186770_completata.xlsx
   [2026-05-28T10:32:36] [INFO] Anexa 3 completat: 04_Anexa-3-completata.docx
   [2026-05-28T10:32:37] [INFO] Anexa 4 completat: 05_Anexa-4-completata.docx
   [2026-05-28T10:32:38] [INFO] Sinteză produsă: 02_Sinteza_RO14186770.md + .docx
   ```

## Răspuns final către utilizator (inline)

> Analiza completă pentru ROMACTIV BUSINESS CONSULTING SRL (CUI RO14186770):
>
> **Verdict: Există legături — DA**
>
> Clasificare: ÎNTREPRINDERE LEGATĂ, categorie ÎNTREPRINDERE MICĂ.
>
> Motivul legăturii: asociatul unic POPESCU PAUL deține majoritar (60%) firma 32165478 (CAEN 7022, piață identică cu solicitantul).
>
> Firma 25896374 a fost EXCLUSĂ din clasificare — piață NU învecinată (CAEN 9999 vs 7022).
>
> Fișiere generate (toate în `analize-imm/2026-05-28_14186770/`):
> - 02_Sinteza_RO14186770.md / .docx
> - 03_Declaratie_IMM_14186770_completata.xlsx
> - 04_Anexa-3-completata.docx
> - 05_Anexa-4-completata.docx
> - audit.log
