# Format "Analiza încadrare IMM" — livrabil OBLIGATORIU (generator)

Pe lângă cele 3 documente oficiale (Declarație IMM, Anexa 3, Anexa 4), pluginul produce
**la fiecare analiză** și un **workbook de analiză completă** într-un singur sheet —
formatul intern RBC folosit la faza de contractare. Generat de
`scripts/fill_analiza_imm.py` → `build_analiza_imm(...)`.

## De ce un generator (nu un template fix)

Declarația IMM fixă suportă max 8 parteneri + 10 legate. Analiza de grup are număr
variabil de firme și ani. De aceea acest output se **construiește de la zero** cu openpyxl,
scalând la orice dimensiune. Coloanele derivate (euro, mii) = **numere calculate** (`lei/curs`,
`/1000`) — vizibile în orice viewer/PDF (openpyxl nu calculează formule, deci nu le lăsăm
ca formule). TOTAL-ul sumează doar întreprinderile luate în calcul (exclude firmele de
pe piețe neînvecinate / sub prag).

## Structura sheet-ului (5 secțiuni)

### Antet
Solicitant (denumire + CUI), **Program / apel de finanțare** (întrebat la începutul
task-ului), data analizei, sursele.

### 1. STRUCTURA GRUPULUI
Per firmă din grup, un bloc cu coloanele:
`Denumire societate (+CUI) | Asociati | Cota de participare | Administrator | Relatii comerciale / de alta natura (DA/NU) | Adresa`
+ 2 rânduri: `CAEN principal` și `CAEN preponderent`.

**Sursa obligatorie**: termene.ro — AMBELE secțiuni: „Asociați/acționari" **ȘI**
„Persoane autorizate" + conexiunile cu alte firme ale ambelor categorii. Administratorii
pot lega firme fără participație la capital!

Coloana **Relatii comerciale** (DA/NU) e importantă: semnalează dacă între firme există
flux comercial efectiv — argument pentru piață învecinată / amonte-aval.

### 2. CONCLUZIE LEGĂTURI
Verdict colorat: `LEGATA` / `PARTENERA` / `AUTONOMA` / `NECLAR` + lista firmelor cuplate +
notă explicativă (motivul includerii/excluderii fiecărei firme).

### 3. DATE FINANCIARE CONSOLIDATE (per an)
Câte un tabel pentru fiecare an analizat (tipic 3-4 ani: N-2 … N), fiecare cu coloanele:
`Denumire | Nr. angajati | CA (lei) | Active totale (lei) | CA (euro) | Active totale (euro) | CA (mii lei) | Active totale (mii lei) | CA (mii euro) | Active totale (mii euro)` + TOTAL.

- **Active totale (lei)** = active imobilizate + active circulante + cheltuieli în avans.
- **CA (euro)** = `CA_lei / curs` (rotunjit 2 zecimale); **mii** = `/1000` (3 zecimale) — **numere calculate**, nu formule (ca să fie vizibile în orice viewer; curs = BNR 31 dec).
- **curs euro** = cursul BNR din **ULTIMA ZI a anului** (31 decembrie) — vezi mai jos.
  NICIODATĂ curs mediu anual sau curs curent.
- **TOTAL** = suma întreprinderilor LUATE ÎN CALCUL (firmele excluse — piață neînvecinată
  sau sub prag — apar în tabel marcate `[EXCLUS]` dar NU intră în TOTAL).

### 4. CATEGORIA IMM
Încadrarea finală: `MICROINTREPRINDERE` / `MICA` / `MIJLOCIE` / `NU ESTE IMM`.

### 5. RECOMANDARE CLAUDE
Secțiune OBLIGATORIE (fundal galben): concluzia consultantului AI în 2-5 fraze —
ce s-a constatat, ce cazuri NECLAR rămân de clarificat (și cu cine), ce variantă de
declarare se recomandă (și ce documente corespund fiecărui scenariu).

## Separatori numerici (convenția românească)

- Zecimale cu **virgulă**, mii cu **punct**: `1.234.567,89`.
- În **Excel**: number_format standard (`#,##0.00` / `#,##0.000` pentru mii) — separatorii
  sunt redați de Excel conform setărilor regionale (pe Windows România: corect automat).
- În **texte** (MD/DOCX/sinteză): folosește `fmt_ro(value, decimals)` din
  `fill_analiza_imm.py` / `sinteza.py`.

## Curs BNR — ULTIMA ZI a anului de referință (31 decembrie)

| An | Curs EUR/RON (31 dec) |
|---|---|
| 2022 | 4,9474 |
| 2023 | 4,9746 |
| 2024 | 4,9741 |
| 2025 | 5,0985 |

Pentru ani care nu apar în tabel: caută cursul oficial BNR din ultima zi lucrătoare a
anului (bnr.ro / cursbnr.ro) și confirmă-l cu utilizatorul la întrebările de început de
task. Actualizează acest tabel când afli valori noi.

## API generator

```python
build_analiza_imm(
    out_path,
    solicitant={"denumire": "...", "cui": "..."},
    companies=[{
        "denumire": "...", "cui": "...",
        "asociati": [{"nume": "...", "cota": 0.80}, ...],  # cota ca fracție (0.80 = 80%)
        "administrator": "...",      # din "Persoane autorizate"!
        "relatii": "DA"/"NU",
        "adresa": "...",
        "caen_principal": "...", "caen_preponderent": "...",
    }, ...],
    verdict={"tip": "LEGATA"/"PARTENERA"/"AUTONOMA"/"NECLAR", "cu": "...", "nota": "..."},
    years=[{
        "an": 2024, "curs": 4.9741,    # BNR 31 decembrie!
        "rows": [{"denumire": "...", "angajati": 90,
                  "ca_lei": 58368342, "active_lei": 71583329,
                  "exclus": False, "nota": "..."}, ...],
    }, ...],
    categorie="INTREPRINDERE MIJLOCIE (IMM)",
    recomandare="Constatări... Clarificări rămase... Se recomandă declararea ca ...",
    program="PNRR C9 / apel ...",     # întrebat la începutul task-ului
)
```

## Când îl folosește pluginul

DUPĂ ce decision engine-ul (`classify.py`) a stabilit tipul și a colectat financiarele,
pluginul generează ACEST workbook ca `06_Analiza_incadrare_IMM_<denumire>.xlsx` în folderul
analizei — **întotdeauna, la fiecare analiză** (livrabil obligatoriu). Este livrabilul de
sinteză tehnică (complementar sintezei narative MD/DOCX).

Pentru cazuri NECLAR: include toate firmele în harta grupului, marchează verdictul ca
NECLAR, iar în tabelul financiar pune firma-solicitant ca rând de bază + firmele posibil
legate cu notă, astfel încât ambele scenarii (autonom = doar solicitantul; legat = TOTAL)
să fie vizibile simultan. Recomandarea Claude explică explicit ambele scenarii.
