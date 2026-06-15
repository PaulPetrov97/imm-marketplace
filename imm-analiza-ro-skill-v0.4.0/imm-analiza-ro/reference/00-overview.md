# Overview — reference docs

Ordinea recomandată de citire de către agent (skill-ul) când are nevoie de context:

| Fișier | Scop | Când să fie consultat |
|---|---|---|
| `01-legea-346-2004.md` | Praguri IMM RO + cazuri speciale | La fiecare clasificare finală |
| `02-definitie-eu-2003-361.md` | Manualul CE complet (60 pagini) | Doar când arborele decizional nu acoperă cazul |
| `03-arbore-decizional.md` | Logica distilată + pseudocod | În fiecare apel al `classify.py` |
| `04-termene-ro-selectors.md` | Selectori DOM + URL patterns | În fiecare apel al `scrape_termene.py` |
| `05-cazuri-neclare.md` | Taxonomia ambiguităților + întrebări sample | Când decision engine returnează `NECLAR` |

Convenție: skill-ul citește DOAR `03` și `04` la fiecare invocare (sunt scurte). `02` (manualul) e fallback pentru cazuri exotice. `01` și `05` sunt referință permanentă pentru utilizator.
