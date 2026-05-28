"""Path helpers — cross-platform pentru skill-ul imm-analiza-ro."""
from __future__ import annotations

import os
import sys
from datetime import date
from pathlib import Path


def skill_root() -> Path:
    """Path-ul rădăcină al skill-ului (folderul care conține SKILL.md).

    Preferă ${CLAUDE_PLUGIN_ROOT} dacă e setat (instalare prin marketplace),
    altfel deduce din locația acestui fișier.
    """
    env = os.environ.get("CLAUDE_PLUGIN_ROOT")
    if env:
        return Path(env) / "skills" / "imm-analiza-ro"
    return Path(__file__).resolve().parent.parent


def templates_dir() -> Path:
    return skill_root() / "templates"


def reference_dir() -> Path:
    return skill_root() / "reference"


def template_path(name: str) -> Path:
    p = templates_dir() / name
    if not p.exists():
        raise FileNotFoundError(f"Template missing: {p}")
    return p


def output_root(cwd: Path | None = None) -> Path:
    cwd = cwd or Path.cwd()
    return cwd / "analize-imm"


def analysis_dir(cui_solicitant: str, cwd: Path | None = None,
                 date_override: date | None = None) -> Path:
    d = (date_override or date.today()).isoformat()
    safe_cui = "".join(c for c in cui_solicitant if c.isalnum())
    p = output_root(cwd) / f"{d}_{safe_cui}"
    p.mkdir(parents=True, exist_ok=True)
    (p / "01_raw").mkdir(exist_ok=True)
    return p


def python_exe() -> str:
    """Cross-platform Python executable."""
    return sys.executable or ("python.exe" if os.name == "nt" else "python3")


if __name__ == "__main__":
    print(f"skill_root      = {skill_root()}")
    print(f"templates_dir   = {templates_dir()}")
    print(f"reference_dir   = {reference_dir()}")
    print(f"output_root(cwd)= {output_root()}")
