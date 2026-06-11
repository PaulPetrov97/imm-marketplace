# imm-analiza-ro — skill pentru analiză preliminară IMM

## Ce face

Automatizează analiza preliminară de încadrare în categoria IMM (autonomă / parteneră / legată) pentru întreprinderi românești identificate prin CUI, conform:

- **Legea 346/2004** (legea IMM-urilor — implementare RO).
- **Recomandarea CE 2003/361/EC** (definiția europeană a IMM-urilor).

## Flux end-to-end

1. **Preflight**: verifică Chrome MCP + dependențele Python.
2. **Scrape termene.ro** (folosind sesiunea utilizatorului — Premium / PartnerSCAN ideal).
3. **Recursie controlată** pe asociații PF (alte firme ale lor) și PJ (recurs cu cycle detection).
4. **Decision engine** aplică arborele din `reference/03-arbore-decizional.md`.
5. **Întrebări de clarificare** pentru cazurile ambigue (rude concertate, piață învecinată, investitori-excepție).
6. **Completare documente** (păstrând formulele și structura intactă):
   - Declarație IMM (Excel — Sheet 2 + Sheet 4).
   - Anexa 3 (Word — declarație de încadrare cu bifa corectă).
   - Anexa 4 (Word — calcul partenere proporțional + legate integral).
7. **Sinteza** 1–3 pagini cu tabele și verdict "Există legături: DA/NU/NECLAR".
8. **Audit log** cu sursele și raționamentele.

## Trigger

- Slash command (dacă instalat ca plugin): `/imm <CUI1> [CUI2 ...] [--an YYYY] [--output <path>]`.
- Frază naturală (oricum): "analiză IMM", "verifică firma", "analizează CUI", "fă-mi declarația IMM", etc.

## Output

În folderul curent: `analize-imm/<YYYY-MM-DD>_<CUI-solicitant>/`

```
01_raw/                        # JSON-uri brute termene.ro
02_Sinteza_<CUI>.md            # sinteza Markdown (afișată inline)
02_Sinteza_<CUI>.docx          # sinteza Word
03_Declaratie_IMM_<CUI>_completata.xlsx
04_Anexa-3-completata.docx
05_Anexa-4-completata.docx
audit.log
```

## Dependențe

- Python 3.10+
- `openpyxl ≥ 3.1`
- `python-docx ≥ 1.1`
- Extensia Chrome `Claude in Chrome` cu sesiune termene.ro activă.

```
pip install openpyxl python-docx
```

## Invariante

- Template-urile bundled (`templates/`) sunt READ-ONLY. Pluginul scrie întotdeauna în fișiere noi.
- Sheet-urile `Date_consolidate`, `Calcul partenere & legate`, `Sect A - Intr. partenere`, `Tabel B1`, `Tabel B2`, `Ipoteze` NU sunt atinse.
- Row 12 din `Date_partenere` (formule) rămâne intactă.
- Maxim 8 parteneri și 10 legate per analiză (limitarea template-ului).
- Audit log obligatoriu.

## Limitări cunoscute

- Termene.ro PF profile pages necesită cont PartnerSCAN. Fără el, cross-reference asociaților PF se face manual (pluginul cere lista).
- Maxim 30 navigări per sesiune (cadență politicoasă).
- Firme străine: pluginul nu recursionează — datele se introduc manual.
- Bilanțul ANAF pentru anul N apare după 30 mai N+1; pentru analize anterioare se folosește N-1 cu warning.

## Suport

Paul Petrov, RomActiv Business Consulting — ecofin@consultant.com.
