"""Orchestrator pentru scrape-ul termene.ro prin Chrome MCP.

NOTĂ IMPORTANTĂ: Acest fișier NU face request-uri HTTP directe. Este un
orchestrator care documentează pașii pe care Claude îi execută prin tool-urile
`mcp__Claude_in_Chrome__*`. Funcțiile aici sunt:
- (a) helper-uri pentru parsare DOM extras de Claude in Chrome,
- (b) validare/normalizare CUI,
- (c) schema JSON pentru output,
- (d) detecție pattern-uri de rate-limit / Cloudflare în text-ul paginii.

Claude apelează aceste funcții pasând conținutul pagini extras prin
get_page_text / read_page / find.
"""
from __future__ import annotations

import json
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional


# -------------------- CUI normalization --------------------

CUI_RE = re.compile(r"^(RO)?\s*(\d{2,10})$", re.IGNORECASE)


def normalize_cui(raw: str) -> str:
    """Normalizează CUI → formă canonică fără prefix RO, doar digiți."""
    raw = raw.strip().upper().replace(" ", "")
    m = CUI_RE.match(raw)
    if not m:
        raise ValueError(f"CUI invalid: {raw}")
    return m.group(2)


def validate_cui_checksum(cui: str) -> bool:
    """Verifică algoritmul de control ANAF pentru CUI.

    Algoritmul ANAF: înmulțește digiții cu cheia 753217532, sumă mod 11.
    Dacă rest == 10 → cifra de control e 0, altfel == rest.
    """
    if not cui.isdigit() or len(cui) < 2 or len(cui) > 10:
        return False
    key = "753217532"
    main = cui[:-1].rjust(9, "0")  # rjust pentru CUI scurte
    digits = main[-9:]
    s = sum(int(d) * int(k) for d, k in zip(digits, key))
    control = (s * 10) % 11
    if control == 10:
        control = 0
    return control == int(cui[-1])


def termene_url(cui: str) -> str:
    return f"https://termene.ro/firma/{normalize_cui(cui)}"


# -------------------- Schema JSON --------------------

@dataclass
class Financiar:
    an: int
    salariati: Optional[float] = None
    cifra_afaceri_lei: Optional[float] = None
    active_totale_lei: Optional[float] = None
    profit_net_lei: Optional[float] = None
    sursa: str = "termene.ro / bilanț ANAF"


@dataclass
class Asociat:
    tip: str  # "PF" sau "PJ"
    nume: str
    cui: Optional[str] = None  # doar pentru PJ
    procent: float = 0.0
    tara: str = "RO"
    rol: Optional[str] = None
    alte_firme: list[dict] = field(default_factory=list)
    investor_type: Optional[str] = None  # set after clarification: business_angel / vc_fund / etc.
    recurse_status: str = "pending"  # pending / done / gated / not_found


@dataclass
class FirmaData:
    cui: str
    denumire: str = ""
    adresa: str = ""
    stare: str = ""
    data_infiintare: str = ""
    caen_principal: dict = field(default_factory=dict)
    caen_secundar: list[dict] = field(default_factory=list)
    reprezentant_legal: list[dict] = field(default_factory=list)
    financiare: list[Financiar] = field(default_factory=list)
    asociati: list[Asociat] = field(default_factory=list)
    scraped_at: str = ""
    source_tier: str = "unknown"  # premium / public / anonymous
    warnings: list[str] = field(default_factory=list)
    source_url: str = ""

    def to_json(self) -> dict:
        d = asdict(self)
        return d


# -------------------- Anti-bot detection --------------------

ANTI_BOT_MARKERS = [
    "cf-challenge",
    "Verifying you are human",
    "Just a moment",
    "Checking your browser",
    "DDoS protection by Cloudflare",
    "Too Many Requests",
    "429",
]


def detect_anti_bot(page_text: str) -> Optional[str]:
    for m in ANTI_BOT_MARKERS:
        if m.lower() in page_text.lower():
            return m
    return None


# -------------------- Tier detection --------------------

PREMIUM_MARKERS = [
    "shareholders-detail",
    "actionari-detaliu",
    "PartnerSCAN",
    "data-section=\"shareholders\"",
    "actionari-completi",
]

LOGIN_MARKERS = [
    "data-test=\"user-menu\"",
    "account-link",
    "Contul meu",
    "Deconectare",
    "Logout",
]


def detect_tier(homepage_text: str, firm_page_text: str) -> str:
    logged_in = any(m in homepage_text for m in LOGIN_MARKERS)
    premium = any(m in firm_page_text for m in PREMIUM_MARKERS)
    if logged_in and premium:
        return "premium"
    if logged_in:
        return "logged_in_basic"
    return "anonymous"


# -------------------- Save raw --------------------

def save_raw(data: FirmaData, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    out = output_dir / f"{data.cui}.termene.json"
    out.write_text(json.dumps(data.to_json(), ensure_ascii=False, indent=2),
                   encoding="utf-8")
    return out


# -------------------- Audit log --------------------

def audit_log(audit_path: Path, message: str, level: str = "INFO") -> None:
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().isoformat(timespec="seconds")
    line = f"[{ts}] [{level}] {message}\n"
    with audit_path.open("a", encoding="utf-8") as f:
        f.write(line)


# -------------------- CLI smoke test --------------------

def _smoke():
    """Test CUI normalization + checksum."""
    test_cases = [
        ("RO 14186770", "14186770"),
        ("14186770", "14186770"),
        ("  RO14186770  ", "14186770"),
    ]
    for raw, expected in test_cases:
        got = normalize_cui(raw)
        assert got == expected, f"{raw} → {got} ≠ {expected}"
        print(f"OK: {raw!r} -> {got}")
    # Checksum: termene.ro lists CUI 14186770 (RomActiv); test that real CUIs validate
    # Known valid: 14186770 (RomActiv Business Consulting)
    print(f"Checksum 14186770: {validate_cui_checksum('14186770')}")


if __name__ == "__main__":
    _smoke()
