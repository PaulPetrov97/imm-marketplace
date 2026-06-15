# Arbore decizional — clasificare IMM (autonomă/parteneră/legată)

Distilare a manualului CE și a Legii 346/2004 în logică executabilă. Citește acest fișier la fiecare apel al `classify.py`.

## Praguri

```
THRESHOLD_AUT     = 25%   # < acest prag = autonomă
THRESHOLD_LINKED  = 50%   # ≥ acest prag = legată
BUSINESS_ANGEL_MAX_EUR = 1_250_000  # per investitor per firmă
PUBLIC_HOLDING_DISQUALIFY = 25%     # cumulat ≥ → NU_IMM (cu excepții)
```

## Categorii de investitori-excepție (nu se numără la pragul 25% dacă <50%)

```
EXEMPT_INVESTORS = {
  "business_angel": "max €1.25M per firmă; over → numără",
  "vc_fund":        "fond de capital de risc",
  "university":     "universitate / centru cercetare nonprofit",
  "research_nonprofit": "institut cercetare nonprofit",
  "institutional_investor": "fond de pensii, SIF, Fond Proprietatea",
  "small_local_authority": "buget anual <€10M ȘI populație <5000"
}
```

## Algoritm principal

```python
def classify(graf, cui_solicitant, clarificari):
    """
    Intrare:
      graf — noduri (companii + PF), muchii ponderate cu % deținere
      cui_solicitant — CUI-ul firmei solicitante
      clarificari — răspunsuri user la AskUserQuestion (rude, piață învecinată, excepții)
    Ieșire:
      ClassificationResult(
        clasificare ∈ {AUTONOMA, PARTENERA, LEGATA, NU_IMM},
        parteneri[], legate[], neclar[],
        agregat = {salariati, cifra_afaceri, active_totale}
      )
    """
    S = cui_solicitant
    upstream = graf.holders_of(S)      # cine deține pe S
    downstream = graf.held_by(S)       # pe cine deține S

    # =========================================================
    # PASUL 1: Verifică deținere publică cumulată
    # =========================================================
    public_total = sum(h.procent for h in upstream
                       if h.tip == "PJ" and h.investor_type in {"public_institution"})
    # Aplică excepții (autoritate locală mică, fond instituțional public exceptat)
    public_effective = public_total - sum_exempt_public(upstream, clarificari)
    if public_effective >= 25:
        return ClassificationResult("NU_IMM", [], [], [],
            reason="Deținere publică efectivă ≥25% — nu se încadrează ca IMM (Art. 4(4))")

    # =========================================================
    # PASUL 2: Filtrează investitorii-excepție (Art. 3(2))
    # =========================================================
    effective_upstream = []
    for h in upstream:
        if h.investor_type in EXEMPT_INVESTORS and h.procent < 50 and not h.is_affiliated:
            if h.investor_type == "business_angel" and h.investment_eur >= 1_250_000:
                effective_upstream.append(h)  # excepție pierdută
                continue
            continue  # exclus
        effective_upstream.append(h)

    # =========================================================
    # PASUL 3: Detectează grupuri "acționare concertată"
    # =========================================================
    pf_holders = [h for h in effective_upstream if h.tip == "PF"]
    concert_groups = group_by_family(pf_holders, clarificari.family_relations)
    # Familie = soț/soție + ascendenți/descendenți linie directă
    # + clarificări explicite ale userului

    # =========================================================
    # PASUL 4: Detectează LEGĂTURI (≥50% sau cele 4 tipuri art. 3(3))
    # =========================================================
    legate = []
    clasificare = "AUTONOMA"

    # 4a. Upstream PJ ≥50% în solicitant
    for h in effective_upstream:
        if h.tip == "PJ" and h.procent >= 50:
            legate.append(h)
            clasificare = "LEGATA"

    # 4b. Downstream ≥50%
    for d in downstream:
        if d.procent >= 50:
            legate.append(d)
            clasificare = "LEGATA"

    # 4c. Grup concertat ≥50% în solicitant
    for grup in concert_groups:
        proc_in_S = grup.combined_procent_in(S)
        if proc_in_S >= 50:
            # Firmele "soră" controlate de același grup → legate prin piață învecinată
            sibling_companies = grup.other_controlled_companies(min_procent=50)
            for sib in sibling_companies:
                if adjacent_market(S, sib, clarificari):
                    legate.append(sib)
                    clasificare = "LEGATA"
                elif adjacent_market(S, sib, clarificari) == "UNCLEAR":
                    neclar.append((sib, "piață învecinată neconfirmată"))

    # 4d. Control board / contract (cele 4 tipuri art. 3(3))
    # — necesită input utilizator dacă nu e detectabil din termene.ro:
    for h in effective_upstream + downstream:
        for tip_control in clarificari.control_relations.get(h.cui, []):
            if tip_control in {"majoritate_voturi", "numire_admin",
                               "influenta_contract", "acord_voturi"}:
                if h.cui not in [l.cui for l in legate]:
                    legate.append(h)
                    clasificare = "LEGATA"

    # =========================================================
    # PASUL 5: Detectează PARTENERI (25–49%)
    # =========================================================
    parteneri = []
    for h in effective_upstream + downstream:
        already_legata = any(l.cui == h.cui for l in legate)
        if not already_legata and 25 <= h.procent < 50:
            parteneri.append(h)
            if clasificare == "AUTONOMA":
                clasificare = "PARTENERA"

    # =========================================================
    # PASUL 6: Aplică Art. 6 (extensie partener-de-legat, legat-de-partener)
    # =========================================================
    # Partener-de-legat: legăturile firmelor LEGATE → devin partenere ale solicitantului
    # Legat-de-partener: legăturile firmelor PARTENERE → devin partenere (cu % adjustat)
    # Partener-de-partener: NU se includ
    extra = expand_art6(parteneri, legate, graf)
    parteneri += extra.parteneri_indirect
    legate += extra.legate_indirect

    # =========================================================
    # PASUL 7: Calculează agregatul consolidat
    # =========================================================
    agregat = aggregate_consolidated(
        solicitant_data=graf.data(S),
        parteneri=parteneri,
        legate=legate
    )
    # = solicitant + Σ (% × partener) + Σ (100% × legată)

    # =========================================================
    # PASUL 8: Verifică praguri IMM finale
    # =========================================================
    if agregat.salariati >= 250 or (
        agregat.cifra_afaceri_eur > 50_000_000
        and agregat.active_totale_eur > 43_000_000
    ):
        # depășiri în cumulat — nu mai e IMM
        clasificare_marime = "NU_IMM"
    else:
        clasificare_marime = imm_category(agregat)  # micro / mica / mijlocie

    return ClassificationResult(
        clasificare_tip=clasificare,         # autonomă/parteneră/legată
        clasificare_marime=clasificare_marime,  # micro/mica/mijlocie/nu_imm
        parteneri=parteneri,
        legate=legate,
        neclar=neclar,
        agregat=agregat,
    )
```

## Helper: `adjacent_market`

```python
def adjacent_market(cui_a, cui_b, clarificari):
    """Returnează True / False / 'UNCLEAR'."""
    caen_a = lookup_caen(cui_a)
    caen_b = lookup_caen(cui_b)
    # Același CAEN 4-digit → identic, sigur adjacent
    if caen_a == caen_b:
        return True
    # Aceeași diviziune 2-digit → probabil adjacent, întreabă userul
    if caen_a[:2] == caen_b[:2]:
        return clarificari.confirm_adjacent(cui_a, cui_b, default="YES")
    # CAEN total diferit → vertical chain posibil (amonte-aval)
    # Întreabă userul; default UNCLEAR
    return clarificari.confirm_adjacent(cui_a, cui_b, default="UNCLEAR")
```

## Helper: `group_by_family`

```python
def group_by_family(pf_holders, family_relations):
    """
    family_relations: dict de tip {"POPESCU ION": ["POPESCU MARIA (soție)"], ...}
    Generat din răspunsurile user la AskUserQuestion.
    Returnează listă de grupuri concertate, fiecare cu deținerile combinate.
    """
    groups = []
    visited = set()
    for pf in pf_holders:
        if pf.nume in visited:
            continue
        family = family_relations.get(pf.nume, [])
        group_members = [pf] + [find_pf(m) for m in family]
        groups.append(ConcertGroup(members=group_members))
        for m in group_members:
            visited.add(m.nume)
    return groups
```

## Cazuri de oprire timpurie

| Condiție | Verdict |
|---|---|
| Deținere publică efectivă ≥25% | NU_IMM, fără calcul agregat |
| Solicitant nu există pe termene.ro | NECLAR — cere CUI corect |
| >8 parteneri detectați | STOP — template nu suportă, restrânge scope |
| >10 legate detectate | STOP — template nu suportă |
| Cycle de control detectat | Marchează în graf, continuă (cu warning) |
| Asociat PF cu profil gated | NECLAR — cere user lista manuală |

## Reguli de NICIODATĂ

- Nu agrega 100% pentru un partener (Art. 3(2) Recomandare).
- Nu aplica 50% pentru o legată.
- Partener-de-partener NU se include în consolidare.
- Investitor-excepție valabil DOAR dacă <50% și nu e afiliat solicitantului.
- Pragul "salariați <250" este STRICT mai mic (249 OK, 250 NU).
- Pragurile financiare sunt în EUR — convertește RON cu cursul BNR 31 decembrie an de referință.
