# Changelog

## v0.1.0 — 2026-05-28

Versiune inițială a pluginului `imm-analiza-ro`.

### Adăugat
- Slash command `/imm <CUI1> [CUI2 ...] [--an YYYY] [--output <path>]`.
- Skill `imm-analiza-ro` cu trigger pe fraze românești pentru analiza IMM.
- Arbore decizional complet pentru clasificare autonomă/parteneră/legată conform Legii 346/2004 + Recomandării CE 2003/361/EC.
- Scripts pentru completare automată:
  - `fill_xlsx.py` — Declarație IMM (Sheet 2 Date_partenere + Sheet 4 Date_legate) păstrând formulele.
  - `fill_docx_anexa3.py` — Anexa 3 cu bifa corectă pe tipul de întreprindere.
  - `fill_docx_anexa4.py` — Anexa 4 cu calcul partenere (proporțional) + legate (integral).
  - `sinteza.py` — sinteză MD + DOCX cu tabele și verdict.
- Reference docs:
  - Manualul CE complet (60 pagini) — bilia regulilor.
  - Distilare arbore decizional cu pseudocod.
  - Selectori DOM termene.ro pentru Chrome MCP.
  - Taxonomia cazurilor neclare cu întrebări sample.
- Preflight check (Python deps + template integrity).
- Audit log per analiză.
- Suport distribuție duală: marketplace git + .zip backup.
