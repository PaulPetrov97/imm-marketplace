# Format "Analiza încadrare IMM" — output bogat (generator)

Pe lângă cele 3 documente oficiale (Declarație IMM, Anexa 3, Anexa 4), pluginul produce
și un **workbook de analiză completă** într-un singur sheet — formatul intern RBC folosit
la faza de contractare. Generat de `scripts/fill_analiza_imm.py` →
`build_analiza_imm(...)`.

## De ce un generator (nu un template fix)

Declarația IMM fixă suportă max 8 parteneri + 10 legate. Analiza de grup are număr
variabil de firme și ani. De aceea acest output se **construiește de la zero** cu openpyxl,
scalând la orice dimensiune. Coloanele euro rămân **formule live** (`=lei/curs`), iar
TOTAL-ul sumează doar întreprinderile luate în calcul (exclude firmele de pe piețe
neînvecinate / sub prag).

## Structura sheet-ului (4 secțiuni)

### 1. STRUCTURA GRUPULUI
Per firmă din grup, un bloc cu coloanele:
`Denumire societate (+CUI) | Asociati | Cota de participare | Administrator | Relatii comerciale / de alta natura (DA/NU) | Adresa`
+ 2 rânduri: `CAEN principal` și `CAEN preponderent`.

Coloana **Relatii comerciale** (DA/NU) e importantă: semnalează dacă între firme există
flux comercial efectiv — argument pentru piață învecinată / amonte-aval.

### 2. CONCLUZIE LEGĂTURI
Verdict colorat: `LEGATA` / `PARTENERA` / `AUTONOMA` / `NECLAR` + lista firmelor cuplate +
notă explicativă (motivul includerii/excluderii fiecărei firme).

### 3. DATE FINANCIARE CONSOLIDATE (per an)
Câte un tabel pentru fiecare an analizat (tipic 3-4 ani: N-2 … N), fiecare cu:
`Denumire | Nr. angajati | CA (lei) | Active totale (lei) | CA (euro) | Active totale (euro)` + TOTAL.
- **Active totale (lei)** = active imobilizate + active circulante + cheltuieli în avans.
- **CA (euro)** = `=CA_lei / curs`, **Active (euro)** = `=Active_lei / curs` (formule live).
- **curs euro** = cursul BNR la 31 decembrie al anului (vezi constante mai jos).
- **TOTAL** = suma întreprinderilor LUATE ÎN CALCUL (firmele excluse — piață neînvecinată
  sau sub prag — apar în tabel marcate `[EXCLUS]` dar NU intră în TOTAL).

### 4. CATEGORIA IMM
Încadrarea finală: `MICROINTREPRINDERE` / `MICA` / `MIJLOCIE` / `NU ESTE IMM`.

## Constante curs BNR (31 decembrie)

| An | Curs EUR/RON |
|---|---|
| 2022 | 4,9474 |
| 2023 | 4,9746 |
| 2024 | 4,9741 |
| 2025 | 5,0985 |

(Valori din practica RBC / template Hexol. A se confirma cu cursul oficial BNR la depunere
pentru anii noi.)

## API generator

```python
build_analiza_imm(
    out_path,
    solicitant={"denumire": "...", "cui": "..."},
    companies=[{
        "denumire": "...", "cui": "...",
        "asociati": [{"nume": "...", "cota": 0.80}, ...],  # cota ca fracție (0.80 = 80%)
        "administrator": "...",
        "relatii": "DA"/"NU",
        "adresa": "...",
        "caen_principal": "...", "caen_preponderent": "...",
    }, ...],
    verdict={"tip": "LEGATA"/"PARTENERA"/"AUTONOMA"/"NECLAR", "cu": "...", "nota": "..."},
    years=[{
        "an": 2024, "curs": 4.9741,
        "rows": [{"denumire": "...", "angajati": 90,
                  "ca_lei": 58368342, "active_lei": 71583329,
                  "exclus": False, "nota": "..."}, ...],
    }, ...],
    categorie="INTREPRINDERE MIJLOCIE (IMM)",
)
```

## Când îl folosește pluginul

DUPĂ ce decision engine-ul (`classify.py`) a stabilit tipul și a colectat financiarele,
pluginul generează ACEST workbook ca `06_Analiza_incadrare_IMM_<denumire>.xlsx` în folderul
analizei. Este livrabilul de sinteză tehnică (complementar sintezei narative MD/DOCX).

Pentru cazuri NECLAR: include toate firmele în harta grupului, marchează verdictul ca
NECLAR, iar în tabelul financiar pune firma-solicitant ca rând de bază + firmele posibil
legate cu notă, astfel încât ambele scenarii (autonom = doar solicitantul; legat = TOTAL)
să fie vizibile simultan.
