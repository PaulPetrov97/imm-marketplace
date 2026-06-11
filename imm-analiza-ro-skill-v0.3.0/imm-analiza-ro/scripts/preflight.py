"""Preflight check pentru imm-analiza-ro.

Verifică:
1. Python ≥3.10
2. openpyxl + python-docx instalate
3. Template-urile există în plugin
4. Chrome MCP — semnal indirect (instrucțiuni pentru Claude, nu test executabil din Python)

Output: JSON pe stdout + exit code (0 = OK, 1 = fail).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


def check_python() -> tuple[bool, str]:
    v = sys.version_info
    ok = (v.major, v.minor) >= (3, 10)
    return ok, f"Python {v.major}.{v.minor}.{v.micro}"


def check_module(name: str) -> tuple[bool, str]:
    try:
        __import__(name)
        mod = sys.modules[name]
        version = getattr(mod, "__version__", "unknown")
        return True, f"{name} {version}"
    except ImportError:
        return False, f"{name} NU este instalat — rulează: pip install {name}"


def check_templates() -> tuple[bool, list[str]]:
    try:
        from _paths import templates_dir
    except ImportError:
        sys.path.insert(0, str(Path(__file__).parent))
        from _paths import templates_dir
    required = [
        "Declaratie_IMM_v1.0.xlsx",
        "Anexa-3-Declaratie-incadrare-IMM.docx",
        "Anexa-4-Calcul-intreprinderi.docx",
    ]
    tdir = templates_dir()
    missing = [r for r in required if not (tdir / r).exists()]
    return len(missing) == 0, missing


def main() -> int:
    results: dict = {"checks": [], "overall": "PASS"}

    py_ok, py_msg = check_python()
    results["checks"].append({"name": "python_version", "ok": py_ok, "info": py_msg})

    for mod in ["openpyxl", "docx", "lxml"]:
        # python-docx imports as "docx"
        ok, msg = check_module(mod)
        pkg_name = "python-docx" if mod == "docx" else mod
        results["checks"].append({"name": pkg_name, "ok": ok, "info": msg})

    tpl_ok, missing = check_templates()
    results["checks"].append({
        "name": "templates",
        "ok": tpl_ok,
        "info": "Toate template-urile prezente" if tpl_ok else f"Lipsesc: {missing}",
    })

    # Chrome MCP — instrucțiune doar (nu poate fi testată din Python)
    results["checks"].append({
        "name": "chrome_mcp",
        "ok": None,
        "info": "Chrome MCP trebuie verificat din Claude prin mcp__Claude_in_Chrome__list_connected_browsers",
    })

    if any(c["ok"] is False for c in results["checks"]):
        results["overall"] = "FAIL"

    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0 if results["overall"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
