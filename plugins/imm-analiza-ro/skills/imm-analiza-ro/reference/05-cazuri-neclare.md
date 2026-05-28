# Cazuri neclare — taxonomie + întrebări sample

Când decision engine întâlnește un caz neidentificabil din date publice, NU ghicește — întreabă utilizatorul sau marchează NECLAR în sinteză.

## Categoria 1: Acționare concertată (rude)

**Declanșator**: ≥2 asociați PF în solicitant sau în firmele găsite în recursie, cu nume similare (același nume de familie) sau cu indicii de rudenie.

**Întrebare sample (AskUserQuestion)**:
```
Asociatul POPESCU ION (50% în Firma A) și POPESCU MARIA (30% în Firma B, CAEN identic) —
există relație de rudenie linie directă (soți / părinți / copii / bunici / nepoți)?

[Da, sunt rude linie directă]
[Nu, fără relație]
[Nu știu / nu pot confirma]
```

**Consecințe**:
- DA → combine deținerile, verifică prag ≥25% în firme cu piață învecinată → posibil LEGATE.
- NU → tratează ca asociați independenți.
- NU ȘTIU → flag NECLAR, raționament: "Posibilă acționare concertată — necesită confirmare din declarația proprie a solicitantului."

## Categoria 2: Piață învecinată / amonte-aval

**Declanșator**: două firme conectate prin asociați comuni au CAEN diferite la 4 cifre.

**Cazuri**:
- Aceeași diviziune CAEN 2-digit (ex. 47XX = comerț cu amănuntul) → probabil adjacent.
- Diviziuni diferite, dar relație amonte-aval evidentă (ex. fermă vegetală + procesare alimentară; producție textile + retail textile).
- Diviziuni complet diferite (ex. IT + agricultură) → improbabil adjacent.

**Întrebare sample**:
```
Firmele:
- FIRMA A SRL (CAEN 4711 — Comerț cu amănuntul, magazine nespecializate)
- FIRMA B SRL (CAEN 4711 — Identic) → adjacent automat
- FIRMA C SRL (CAEN 5610 — Restaurante)

Activitatea efectivă a FIRMA C este pe piață învecinată cu A/B (ex. produsele A se vând prin C)?

[Da, integrare verticală / piață învecinată confirmată]
[Nu, activități complet independente]
[Nu știu]
```

**Consecințe**:
- DA → conectează în clasificare ca LEGATĂ (dacă control ≥50% prin grup concertat).
- NU → exclude din legăturile transitive prin PF.
- NU ȘTIU → flag NECLAR în sinteză, recomandă verificare suplimentară a obiectului efectiv de activitate.

## Categoria 3: Investitor-excepție

**Declanșator**: un asociat PJ cu denumire neobișnuită (suffix "FOND", "BANK", "UNIVERSITY", "INVEST") deține 25–49% în solicitant.

**Întrebare sample**:
```
Asociatul "VENTURE CAPITAL EAST SA" (RO12345678) deține 35% în solicitant.
Ce categorie de investitor este?

[Societate operațională (numără la 25%)]
[Business angel — investiție <€1.25M (NU numără)]
[Fond de capital de risc / VC (NU numără)]
[Universitate / centru de cercetare nonprofit (NU numără)]
[Fond instituțional / fond de pensii / SIF (NU numără)]
[Organism public / autoritate publică (cumul ≥25% → NU_IMM)]
```

**Consecințe**:
- Categorie exceptată → exclude din calculul pragului 25% (dar verifică să nu fie afiliat).
- Operațional → tratează normal ca partener.
- Public → adaugă la calculul cumul deținere publică.

## Categoria 4: Deținere publică

**Declanșator**: asociat PJ cu nume sugerând stat (ex. "Consiliul Județean", "Primăria", "Ministerul X").

**Întrebare sample**:
```
Asociatul "CONSILIUL LOCAL X" deține 30% în solicitant.
Confirmă: este organism public (în sensul Art. 4(4) Legea 346/2004)?

[Da — organism public]
[Nu — entitate juridică privată cu nume similar]
```

**Consecință**: dacă cumulul deținerii publice efective ≥25% → solicitant NU este IMM. Stop analiza, comunică utilizatorului.

## Categoria 5: Control board / contract (cele 4 tipuri Art. 3(3))

**Declanșator**: utilizatorul indică o relație de control care nu apare în % deținere — ex. contract de management, drept de veto, acord între acționari.

**Întrebare sample (proactivă, în clarificări finale)**:
```
În afara procentelor de deținere identificate automat, există între solicitant și
oricare firmă identificată în analiză:

[Contract care permite influență dominantă (management, franciză cu clauze de control)]
[Acord între acționari pentru exercitarea voturilor]
[Drept statutar de a numi/revoca majoritatea organelor de administrare]
[Niciuna din cele de mai sus]
```

**Consecințe**: dacă DA → firma devine LEGATĂ chiar și fără 50% capital.

## Categoria 6: Asociat străin

**Declanșator**: asociat PJ cu țară ≠ România.

**Întrebare sample**:
```
Asociatul "GLOBAL HOLDING GMBH" (Germania) deține 40% în solicitant.

[Continuă cu cifrele declarate de termene.ro pentru solicitant, fără recursie pe firma străină]
[Furnizez manual financiarele firmei străine (salariați, cifră, active în EUR)]
[Marchez NECLAR — voi obține date ulterior]
```

**Consecințe**: termene.ro nu acoperă firme străine. Default = continuă fără recursie, dar marchează în sinteză ca "Date partener străin neverificate independent".

## Categoria 7: Modificări recente structură

**Declanșator**: termene.ro arată o schimbare a structurii acționariatului în ultimele 6 luni.

**Întrebare sample**:
```
Structura acționariatului solicitantului s-a modificat la {data}.
Datele de referință pentru analiză sunt pentru anul {an}.

[Analizează cu structura curentă (corectă pentru anul calendaristic actual)]
[Analizează cu structura din anul {an} (necesită date istorice — solicitantul trebuie să furnizeze)]
```

**Consecință**: pentru cereri de finanțare RO actuale, default = structura curentă. Pentru analize retrospective, structura istorică.

## Categoria 8: Date financiare lipsă pe anul de referință

**Declanșator**: bilanțul ANAF pentru anul N nu e încă publicat (termen depunere 30 mai an N+1; pentru analize ulterioare 30 mai → date disponibile).

**Decizie automată**:
- Dacă date N lipsesc și data analizei < 1 iunie an N+1 → folosește an N-1, warning în sinteză.
- Dacă date N lipsesc și data analizei ≥ 1 iunie an N+1 → flag NECLAR ("Bilanț {an N} neîncărcat la ANAF — verifică status declarare").
