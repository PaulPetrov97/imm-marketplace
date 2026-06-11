"""Decision engine — clasificare IMM (autonomă/parteneră/legată) + categorie de mărime.

Implementare a arborelui din `reference/03-arbore-decizional.md`.

Input pur — niciun side effect; folosit de skill cu graf construit din scrape.
"""
from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path


# -------------------- Praguri --------------------

THRESHOLD_AUT = 25.0
THRESHOLD_LINKED = 50.0
BUSINESS_ANGEL_MAX_EUR = 1_250_000.0
PUBLIC_HOLDING_DISQUALIFY = 25.0

EXEMPT_INVESTOR_TYPES = {
    "business_angel",
    "vc_fund",
    "university",
    "research_nonprofit",
    "institutional_investor",
    "small_local_authority",
}

PUBLIC_INVESTOR_TYPES = {
    "public_institution",
}

IMM_THRESHOLDS = {
    "micro":   {"max_salariati": 9,   "max_ca_eur": 2_000_000,    "max_active_eur": 2_000_000},
    "mica":    {"max_salariati": 49,  "max_ca_eur": 10_000_000,   "max_active_eur": 10_000_000},
    "mijlocie":{"max_salariati": 249, "max_ca_eur": 50_000_000,   "max_active_eur": 43_000_000},
}


# -------------------- Structuri --------------------

@dataclass
class Edge:
    """Muchie de deținere: nod sursă deține `procent`% din nodul țintă."""
    sursa: str  # CUI sau "PF:nume"
    tinta: str  # CUI
    procent: float
    rol: str = ""


@dataclass
class Holder:
    nume: str
    tip: str  # "PJ" sau "PF"
    cui: str | None = None
    procent: float = 0.0
    tara: str = "RO"
    rol: str = ""
    investor_type: str | None = None
    is_affiliated: bool = False
    investment_eur: float = 0.0


@dataclass
class CompanyFinancials:
    cui: str
    salariati: float = 0.0
    cifra_afaceri_lei: float = 0.0
    active_totale_lei: float = 0.0
    caen_principal: str = ""


@dataclass
class ClarificationInput:
    """Răspunsuri la AskUserQuestion."""
    family_relations: dict = field(default_factory=dict)  # {"POPESCU ION": ["POPESCU MARIA"]}
    adjacent_market: dict = field(default_factory=dict)   # {("CUI_A","CUI_B"): True/False/"UNCLEAR"}
    investor_types: dict = field(default_factory=dict)    # {"CUI": "business_angel"|"vc_fund"|...}
    control_relations: dict = field(default_factory=dict) # {"CUI": ["majoritate_voturi","numire_admin",...]}
    public_holders: set = field(default_factory=set)      # CUI-uri confirmate ca organism public


@dataclass
class ClassificationResult:
    clasificare_tip: str          # AUTONOMA / PARTENERA / LEGATA / NU_IMM
    clasificare_marime: str       # micro / mica / mijlocie / nu_imm
    parteneri: list = field(default_factory=list)
    legate: list = field(default_factory=list)
    neclar: list = field(default_factory=list)
    agregat: dict = field(default_factory=dict)
    reason: str = ""
    public_holding_effective: float = 0.0


# -------------------- Helpers --------------------

def lei_to_eur(lei: float, curs: float) -> float:
    return lei / curs if curs > 0 else 0.0


def imm_category(salariati: float, ca_eur: float, active_eur: float) -> str:
    """Mapează agregatele la categoria de mărime IMM."""
    if salariati >= 250 or (ca_eur > 50_000_000 and active_eur > 43_000_000):
        return "nu_imm"
    for cat, thr in IMM_THRESHOLDS.items():
        if salariati <= thr["max_salariati"] and (
            ca_eur <= thr["max_ca_eur"] or active_eur <= thr["max_active_eur"]
        ):
            return cat
    return "nu_imm"


def adjacent_market(cui_a: str, cui_b: str, caen_a: str, caen_b: str,
                    clarificari: ClarificationInput) -> bool | str:
    """Determină dacă două firme operează pe piețe învecinate."""
    explicit = clarificari.adjacent_market.get((cui_a, cui_b)) \
               or clarificari.adjacent_market.get((cui_b, cui_a))
    if explicit is not None:
        return explicit
    if not caen_a or not caen_b:
        return "UNCLEAR"
    if caen_a == caen_b:
        return True
    if caen_a[:2] == caen_b[:2]:
        return "UNCLEAR"  # aceeași diviziune — întreabă userul
    return "UNCLEAR"  # diviziuni diferite — întreabă userul


# -------------------- Decision engine --------------------

def classify(
    solicitant_cui: str,
    upstream: list[Holder],
    downstream: list[Holder],
    financials: dict[str, CompanyFinancials],
    clarificari: ClarificationInput,
    curs_eur: float = 4.9759,
) -> ClassificationResult:
    """Aplică arborele decizional. Vezi reference/03-arbore-decizional.md."""

    parteneri: list[dict] = []
    legate: list[dict] = []
    neclar: list[dict] = []

    # =====================================================
    # 1. Deținere publică efectivă
    # =====================================================
    public_total = 0.0
    for h in upstream:
        is_public = (h.investor_type in PUBLIC_INVESTOR_TYPES
                     or (h.cui in clarificari.public_holders))
        if is_public and h.investor_type != "small_local_authority":
            public_total += h.procent

    if public_total >= PUBLIC_HOLDING_DISQUALIFY:
        return ClassificationResult(
            clasificare_tip="NU_IMM",
            clasificare_marime="nu_imm",
            reason=f"Deținere publică cumulată {public_total:.2f}% ≥ 25% — Art. 4(4) Legea 346/2004",
            public_holding_effective=public_total,
        )

    # =====================================================
    # 2. Filtrează investitorii-excepție
    # =====================================================
    effective_upstream: list[Holder] = []
    for h in upstream:
        if (h.investor_type in EXEMPT_INVESTOR_TYPES
                and h.procent < 50.0
                and not h.is_affiliated):
            if (h.investor_type == "business_angel"
                    and h.investment_eur >= BUSINESS_ANGEL_MAX_EUR):
                effective_upstream.append(h)  # plafonul depășit
                continue
            continue
        effective_upstream.append(h)

    # =====================================================
    # 3. LEGĂTURI (≥50% sau cele 4 tipuri art. 3(3))
    # =====================================================
    clasificare_tip = "AUTONOMA"

    def _push_legata(h: Holder, motiv: str):
        if not any(l.get("cui") == h.cui for l in legate):
            legate.append({
                "nume": h.nume,
                "cui": h.cui,
                "procent": h.procent,
                "rol": h.rol,
                "motiv": motiv,
            })

    # 3a. Upstream PJ ≥50% în solicitant
    for h in effective_upstream:
        if h.tip == "PJ" and h.procent >= THRESHOLD_LINKED:
            _push_legata(h, f"Deține {h.procent}% în solicitant (≥50%)")
            clasificare_tip = "LEGATA"

    # 3b. Downstream ≥50% (solicitantul deține majoritar)
    for d in downstream:
        if d.procent >= THRESHOLD_LINKED:
            _push_legata(d, f"Solicitantul deține {d.procent}% (≥50%)")
            clasificare_tip = "LEGATA"

    # 3c. Control board / contract (cele 4 tipuri art. 3(3))
    for h in effective_upstream + downstream:
        ctrl_tipuri = clarificari.control_relations.get(h.cui or h.nume, [])
        if any(t in {"majoritate_voturi", "numire_admin",
                     "influenta_contract", "acord_voturi"} for t in ctrl_tipuri):
            _push_legata(h, f"Control {', '.join(ctrl_tipuri)} (Art. 3(3))")
            clasificare_tip = "LEGATA"

    # =====================================================
    # 4. PARTENERI (25–49%, după eliminarea legatelor)
    # =====================================================
    def _push_partener(h: Holder, sens: str):
        if any(l.get("cui") == h.cui for l in legate):
            return
        if any(p.get("cui") == h.cui for p in parteneri):
            return
        parteneri.append({
            "nume": h.nume,
            "cui": h.cui,
            "tip": h.tip,
            "procent": h.procent,
            "rol": h.rol,
            "sens": sens,
        })

    for h in effective_upstream:
        if THRESHOLD_AUT <= h.procent < THRESHOLD_LINKED:
            _push_partener(h, "upstream")
    for d in downstream:
        if THRESHOLD_AUT <= d.procent < THRESHOLD_LINKED:
            _push_partener(d, "downstream")

    if parteneri and clasificare_tip == "AUTONOMA":
        clasificare_tip = "PARTENERA"

    # =====================================================
    # 5. Agregare — solicitant + 100% × legate + % × parteneri
    # =====================================================
    sol = financials.get(solicitant_cui, CompanyFinancials(cui=solicitant_cui))
    agr_salariati = sol.salariati
    agr_ca = sol.cifra_afaceri_lei
    agr_active = sol.active_totale_lei

    for l in legate:
        f = financials.get(l["cui"] or "", CompanyFinancials(cui=l.get("cui") or ""))
        agr_salariati += f.salariati
        agr_ca += f.cifra_afaceri_lei
        agr_active += f.active_totale_lei

    for p in parteneri:
        f = financials.get(p["cui"] or "", CompanyFinancials(cui=p.get("cui") or ""))
        proc = p["procent"] / 100.0
        agr_salariati += f.salariati * proc
        agr_ca += f.cifra_afaceri_lei * proc
        agr_active += f.active_totale_lei * proc

    agr_ca_eur = lei_to_eur(agr_ca, curs_eur)
    agr_active_eur = lei_to_eur(agr_active, curs_eur)

    clasificare_marime = imm_category(agr_salariati, agr_ca_eur, agr_active_eur)

    agregat = {
        "salariati": round(agr_salariati, 2),
        "cifra_afaceri_lei": round(agr_ca, 2),
        "active_totale_lei": round(agr_active, 2),
        "cifra_afaceri_eur": round(agr_ca_eur, 2),
        "active_totale_eur": round(agr_active_eur, 2),
        "curs_eur": curs_eur,
    }

    if clasificare_marime == "nu_imm" and clasificare_tip != "NU_IMM":
        clasificare_tip_final = clasificare_tip
        reason = (
            f"Tip relațional: {clasificare_tip}. Cumulat depășește pragurile IMM "
            f"(salariați {agr_salariati:.0f}, CA €{agr_ca_eur:,.0f}, "
            f"active €{agr_active_eur:,.0f})."
        )
    else:
        clasificare_tip_final = clasificare_tip
        reason = f"Clasificare după Art. 4 Legea 346/2004 + Recomandarea CE 2003/361."

    return ClassificationResult(
        clasificare_tip=clasificare_tip_final,
        clasificare_marime=clasificare_marime,
        parteneri=parteneri,
        legate=legate,
        neclar=neclar,
        agregat=agregat,
        reason=reason,
        public_holding_effective=public_total,
    )


# -------------------- CLI smoke test --------------------

def _smoke():
    # Caz: solicitant autonom (asociat unic PF cu 100%, fără alte firme)
    sol = "14186770"
    upstream = [Holder(nume="POPESCU ION", tip="PF", procent=100.0, rol="Asociat unic")]
    downstream: list[Holder] = []
    fin = {sol: CompanyFinancials(cui=sol, salariati=8, cifra_afaceri_lei=1_500_000,
                                  active_totale_lei=900_000)}
    clarif = ClarificationInput()
    r = classify(sol, upstream, downstream, fin, clarif)
    assert r.clasificare_tip == "AUTONOMA", r
    assert r.clasificare_marime == "micro", r
    print(f"Smoke 1: {r.clasificare_tip} / {r.clasificare_marime} — OK")

    # Caz: solicitant cu partener PJ 30%
    upstream2 = [Holder(nume="HOLDING SRL", tip="PJ", cui="9988776",
                        procent=30.0, rol="Asociat")]
    fin2 = dict(fin)
    fin2["9988776"] = CompanyFinancials(cui="9988776", salariati=50,
                                         cifra_afaceri_lei=8_000_000,
                                         active_totale_lei=5_000_000)
    r2 = classify(sol, upstream2, downstream, fin2, clarif)
    assert r2.clasificare_tip == "PARTENERA", r2
    print(f"Smoke 2: {r2.clasificare_tip} / {r2.clasificare_marime} — OK "
          f"(agregat sal: {r2.agregat['salariati']})")
    expected_sal = 8 + 50 * 0.30
    assert abs(r2.agregat["salariati"] - expected_sal) < 0.01, r2.agregat

    # Caz: solicitant legată (PJ deține 60%)
    upstream3 = [Holder(nume="PARENT SA", tip="PJ", cui="11223344",
                        procent=60.0, rol="Acționar majoritar")]
    fin3 = dict(fin)
    fin3["11223344"] = CompanyFinancials(cui="11223344", salariati=300,
                                          cifra_afaceri_lei=60_000_000,
                                          active_totale_lei=40_000_000)
    r3 = classify(sol, upstream3, downstream, fin3, clarif)
    assert r3.clasificare_tip == "LEGATA", r3
    # 300 + 8 = 308 salariați cumulat → nu IMM
    assert r3.clasificare_marime == "nu_imm", r3
    print(f"Smoke 3: {r3.clasificare_tip} / {r3.clasificare_marime} — OK")

    # Caz: deținere publică 30%
    upstream4 = [Holder(nume="CONSILIUL JUDETEAN X", tip="PJ", cui="22334455",
                        procent=30.0, investor_type="public_institution")]
    r4 = classify(sol, upstream4, downstream, fin, clarif)
    assert r4.clasificare_tip == "NU_IMM", r4
    print(f"Smoke 4: {r4.clasificare_tip} — OK (public_eff = {r4.public_holding_effective}%)")

    print("\nToate smoke tests passed.")


if __name__ == "__main__":
    _smoke()
