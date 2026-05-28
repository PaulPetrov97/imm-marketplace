"""Build backup zip — `imm-analiza-ro-skill-v<version>.zip`.

Zipează DOAR conținutul folderului skill (fără folderele plugin/marketplace
wrappers), pentru colegii care nu folosesc marketplace git și vor să dropeze
folderul în `~/.claude/skills/`.
"""
from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path


def main() -> int:
    root = Path(__file__).parent.resolve()
    plugin_manifest = root / "plugins" / "imm-analiza-ro" / ".claude-plugin" / "plugin.json"
    if not plugin_manifest.exists():
        print(f"EROARE: nu găsesc {plugin_manifest}", file=sys.stderr)
        return 1
    version = json.loads(plugin_manifest.read_text(encoding="utf-8")).get("version", "0.0.0")

    skill_dir = root / "plugins" / "imm-analiza-ro" / "skills" / "imm-analiza-ro"
    if not skill_dir.exists():
        print(f"EROARE: nu găsesc {skill_dir}", file=sys.stderr)
        return 1

    out_zip = root / f"imm-analiza-ro-skill-v{version}.zip"
    if out_zip.exists():
        out_zip.unlink()

    count = 0
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(skill_dir.rglob("*")):
            if path.is_file() and "__pycache__" not in path.parts:
                arcname = Path("imm-analiza-ro") / path.relative_to(skill_dir)
                zf.write(path, arcname.as_posix())
                count += 1

    size_kb = out_zip.stat().st_size / 1024
    print(f"OK: {out_zip} ({count} fișiere, {size_kb:.1f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
