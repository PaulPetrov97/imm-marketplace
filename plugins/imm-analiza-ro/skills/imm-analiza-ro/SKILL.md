---
name: imm-analiza-ro
description: Analiză preliminară de încadrare IMM (autonomă/parteneră/legată) conform Legii 346/2004 și Recomandării CE 2003/361/EC, pentru întreprinderi românești identificate prin CUI. Scrapează date din termene.ro prin Chrome MCP (folosind sesiunea premium/PartnerSCAN a utilizatorului), aplică arborele decizional al legăturilor, completează automat Declarație IMM (Sheet 2 "Date_partenere" și Sheet 4 "Date_legate"), Anexa 3 (declarație încadrare cu bifa corectă autonomă/parteneră/legată) și Anexa 4 (calcul partenere proporțional + legate integral), păstrând neatinse toate formulele existente. Produce o sinteză de 1-3 pagini cu tabele și verdict explicit "există legături: DA/NU/NECLAR". Folosește acest skill ori de câte ori utilizatorul cere — în română sau engleză — "analiză IMM", "analizează CUI", "analiză preliminară întreprindere", "verifică firma", "fă-mi declarația IMM", "încadrare IMM autonomă/parteneră/legată", "Anexa 3 IMM", "Anexa 4 IMM", "calcul partenere legate", "tabel IMM", sau furnizează o listă de CUI-uri în context de finanțare / fonduri europene / PNRR. Skill-ul NU livrează nimic fără verdict explicit (autonomă/parteneră/legată), tabel financiar consolidat și audit.log cu sursele datelor.
---

# imm-analiza-ro — Skill de analiză preliminară IMM

## Prima acțiune

La activare, după ce ai primit lista de CUI-uri (slash command `/imm CUI1 CUI2 ...` sau frază naturală):

1. Rulează `scripts/preflight.py` pentru a verifica: Chrome MCP conectat, `openpyxl` instalat, `python-docx` instalat. Dacă oricare eșuează → STOP cu mesaj clar în română și instrucțiuni de remediere (`pip install openpyxl python-docx` / instalează extensia Claude in Chrome).
2. Confirmă prin AskUserQuestion: (a) primul CUI este solicitantul? (b) anul de referință (default = anul curent − 1, ex. 2024); (c) folder output (default = cwd `analize-imm/<YYYY-MM-DD>_<CUI-solicitant>/`).
3. Probă Chrome MCP: `mcp__Claude_in_Chrome__list_connected_browsers` → navighează la `https://termene.ro/` → detectează dacă utilizatorul este autentificat (Premium/PartnerSCAN) sau nu.

## Invariante CRITICE — niciodată călcate

1. **Template-urile din `templates/` sunt READ-ONLY.** Niciodată `wb.save()` peste ele. Întotdeauna `shutil.copy()` într-un fișier nou în folderul output.
2. **Excel `Declaratie IMM`**: scrie EXCLUSIV pe `Date_partenere!C4:J11` (parteneri, max 8) și `Date_legate!C4:F13` (legate, max 10 — verifică structura la prima rulare). NU atinge `row 12` din Date_partenere (formule). NU atinge alte sheet-uri: `Ipoteze`, `Date_consolidate`, `Calcul partenere & legate`, `Sect A - Intr. partenere`, `Tabel B1 - consolidate`, `Tabel B2 - legate`. NU modifica `number_format` decât cu valoarea originală.
3. **Anexa 3 (Word)**: bifează EXACT una din 3 căsuțe (autonomă/parteneră/legată) prin înlocuire run.text doar pe run-ul Wingdings țintă. NU creează run-uri noi în Tabel 1 (rupe inheritance-ul de font Wingdings → tofu box).
4. **Anexa 4 (Word)**: duplică blocurile "Fișa de parteneriat" / "Fișa privind legătura" prin lxml deepcopy pe elemente `<w:p>` și `<w:tbl>` complete. Niciodată `paragraph.insert_paragraph_before` pentru tabele.
5. **NU inventa date.** Dacă termene.ro lipsește un câmp, marchează `NECUNOSCUT` în JSON-ul de extragere și escaladează la utilizator prin AskUserQuestion sau flag `NECLAR` în sinteză.
6. **NU ghici legăturile ambigue.** Cazurile de "acționare concertată" (rude), "piață învecinată" (CAEN diferite), "investitor excepție" (business angel / universitate / fond) → întotdeauna AskUserQuestion. Dacă utilizatorul nu știe → flag NECLAR explicit în output.
7. **Audit log obligatoriu.** Fiecare analiză scrie `audit.log` cu timestamp, URL-uri scrapate, decizii, clarificări utilizator.

## Workflow — 10 pași

### 1. Preflight + confirmare scope
- Rulează `python scripts/preflight.py` → asigură Chrome MCP, openpyxl, python-docx.
- AskUserQuestion confirmare CUI-uri, an de referință, folder output.

### 2. Pregătire output folder
- Creează `<output>/analize-imm/<YYYY-MM-DD>_<CUI-solicitant>/01_raw/`.
- Inițializează `audit.log` cu antet (CUI-uri, dată, user, scope).

### 3. Scrape solicitant
- `mcp__Claude_in_Chrome__navigate` la `https://termene.ro/firm/<CUI>` (folosește pattern-uri din `reference/04-termene-ro-selectors.md`).
- `mcp__Claude_in_Chrome__get_page_text` + `find` pentru extrageri.
- Salvează raw HTML/JSON în `01_raw/<CUI>.termene.json`.
- Extrage schema: denumire, adresă, CAEN principal+secundare, reprezentant legal, financiare pe ultimii 3 ani, asociați (PF + PJ cu %).

### 4. Recursie controlată (BFS cu cycle detection)
- Coadă inițială: solicitant.
- Pentru fiecare firmă din coadă:
  - Asociați PJ cu ≥25% → adaugă firma respectivă în coadă (depth ≤ 3).
  - Asociați PF cu ≥25% → navighează la profilul PF (`termene.ro/persoana/...`) → extrage `alte_firme` cu % și CAEN.
  - Downstream: firmele în care SOLICITANTUL deține ≥25% (verifică în pagina solicitantului).
- Set vizitat = CUI-uri deja scrapate (evită ciclu).
- Cadență: 1.5–3.5s între request-uri (`mcp__computer-use__wait` sau JavaScript delay). Max 30 pagini per sesiune.

### 5. Construiește graful
- Noduri: companii (CUI) + persoane fizice (nume normalizat).
- Muchii: deținere % (cu direcție).
- Salvează `01_raw/graph.json`.

### 6. Decision engine (`scripts/classify.py`)
Aplică arborele din `reference/03-arbore-decizional.md`:
- Filtrează excepții upstream (business angel <€1.25M, universitate, fond instituțional, autoritate publică); dacă deținere publică cumulată ≥25% → NU_IMM, stop.
- Detectează grupuri "concert" (familie linie directă + clarificări user).
- LEGATĂ: upstream PJ ≥50% în solicitant, SAU downstream ≥50%, SAU grup concertat ≥50% pe piață învecinată, SAU control board/contract.
- PARTENERĂ: relație 25–49% (după ce LEGATĂ a fost eliminată).
- Aplică art. 6 (partener-de-legat și legat-de-partener; partener-de-partener NU).
- Agregare: solicitant + 100% × legate + 50% × parteneri.
- Pragurile IMM: salariați <250, (cifră ≤€50M SAU active ≤€43M).
- Output: clasificare + listă parteneri + listă legate + listă NECLAR.

### 7. Clarificări (AskUserQuestion)
Doar pentru cazurile detectate ca ambigue:
- **Rude care acționează concertat**: "Asociatul X (PF) deține Y% în firma A. Există rude (soț/soție / linie directă) care dețin acțiuni în firme învecinate (B, C, ...)?"
- **Piață învecinată**: "Firmele A (CAEN xxxx) și B (CAEN yyyy) sunt pe piețe învecinate (amonte-aval, complementare)?"
- **Investitor excepție**: "Asociatul HOLDING SA cu 30% în solicitant este: SRL operațional / business angel <€1.25M / fond VC / universitate / investitor instituțional / organism public?"
- **Deținere publică**: "Există ≥25% deținere directă/indirectă de un organism public în solicitant? (Dacă DA → NU este IMM)"

### 8. Completează documentele
- `scripts/fill_xlsx.py` → `03_Declaratie_IMM_<CUI>_completata.xlsx`
- `scripts/fill_docx_anexa3.py` → `04_Anexa-3-completata.docx`
- `scripts/fill_docx_anexa4.py` → `05_Anexa-4-completata.docx`
- `scripts/fill_analiza_imm.py` → `06_Analiza_incadrare_IMM_<denumire>.xlsx` — **workbook de analiză completă** (format RBC contractare): harta grupului (asociați, cote, administrator, relații comerciale DA/NU, CAEN principal+preponderent), verdict legături, tabele financiare consolidate per an (CA + active în lei ȘI euro, formule live `=lei/curs`), categoria IMM finală. Folosește cursurile BNR 31.dec din `reference/06-format-analiza-imm.md` (2023=4,9746; 2024=4,9741; 2025=5,0985). Vezi `reference/06-format-analiza-imm.md` pentru API + structură. Acoperă număr variabil de firme și ani (se construiește de la zero, nu e template fix).

### 9. Sinteza
- `scripts/sinteza.py` → `02_Sinteza_<CUI>.md` (afișat inline) + `02_Sinteza_<CUI>.docx`
- Structura: profil firmă, matrice asociați, clasificare consolidată cu praguri IMM, verdict "Există legături: DA/NU/NECLAR" cu raționament 2–4 propoziții.

### 10. Audit + raportare
- Finalizează `audit.log` cu toate deciziile.
- Răspunde utilizatorului cu sinteza MD inline + listă fișiere generate (cu path-uri clicabile).

## Escalare / cazuri de oprire

- Chrome MCP indisponibil → STOP, instrucțiuni de instalare.
- CUI invalid (regex eșuat sau termene.ro 404) → cere CUI corect.
- Profil PF gated → AskUserQuestion: introduci manual lista de alte firme?
- Bilanț ANAF lipsă pe anul de referință → folosește N-1, warning în sinteză.
- Cycle în recursie → stop ramura, continuă.
- >8 parteneri sau >10 legate (limită template) → STOP, instruiește utilizatorul să restrângă scope-ul.

## Ce să NU faci

- NU începe analiza fără preflight reușit.
- NU livra fișiere fără ca toate cele 6 (raw, sinteză MD+DOCX, Excel Declarație, Anexa 3, Anexa 4, Analiza încadrare IMM) să fie scrise.
- NU bifează mai mult de o căsuță în Anexa 3.
- NU ghici % sau financiare lipsă — întotdeauna întreabă sau marchează NECUNOSCUT.
- NU folosi `wb.save(template_path)` — întotdeauna `out_path` nou.
- NU lăsa `procent` în coloana G ca string — întotdeauna ca decimal (`0.30` pentru 30%), formatul `0.00%` îl afișează corect.

## Referințe în skill

- `reference/01-legea-346-2004.md` — praguri IMM RO + cazuri speciale
- `reference/02-definitie-eu-2003-361.md` — manualul CE complet (60 pagini)
- `reference/03-arbore-decizional.md` — logica distilată cu pseudocod
- `reference/04-termene-ro-selectors.md` — selectori DOM pentru Chrome MCP
- `reference/05-cazuri-neclare.md` — taxonomia ambiguităților + întrebări sample
- `reference/06-format-analiza-imm.md` — formatul workbook-ului de analiză completă + API generator + curs BNR
