# imm-marketplace

Marketplace privat Claude Code pentru consultanță RomActiv — pluginuri de automatizare a analizelor preliminare de încadrare IMM.

## Conține

| Plugin | Versiune | Scop |
|---|---|---|
| `imm-analiza-ro` | 0.1.0 | Analiză preliminară de încadrare IMM (autonomă/parteneră/legată) conform Legii 346/2004 + Recomandării CE 2003/361/EC, cu scrape termene.ro prin Chrome MCP, completare automată Declarație IMM + Anexa 3 + Anexa 4, sinteză 1-3 pagini cu verdict |

---

## Instalare — metoda 1 (recomandată: marketplace git)

În Claude Code:

```
/plugin marketplace add github:paul-petrov/imm-marketplace
/plugin install imm-analiza-ro@imm-marketplace
```

Pentru update:

```
/plugin update imm-analiza-ro
```

## Instalare — metoda 2 (zip, fără git)

1. Descarcă `imm-analiza-ro-skill-v0.1.0.zip` (atașat în Releases sau primit prin email).
2. Extrage conținutul în `~/.claude/skills/imm-analiza-ro/` (sau `%USERPROFILE%\.claude\skills\imm-analiza-ro\` pe Windows).
3. Repornește Claude Code.
4. Verifică instalarea: scrie în Claude `/imm` — ar trebui să apară hint-ul comenzii.

> Notă: metoda zip NU include slash command-ul `/imm` (acela trăiește în plugin). Triggerul rămâne pe frază naturală — „analiză IMM", „verifică firma 14186770", etc.

## Dependențe

- Claude Code instalat și autentificat.
- **Chrome MCP** (extensia Claude in Chrome): https://www.anthropic.com/claude-code — necesară pentru scrape termene.ro.
- **Python 3.10+** cu `openpyxl` și `python-docx`:
  ```
  pip install openpyxl python-docx
  ```
- Cont termene.ro (recomandat: PartnerSCAN / Premium pentru cross-reference asociați PF).

## Utilizare rapidă

După instalare, în Claude Code:

```
/imm 14186770
```

Sau:

```
/imm 14186770 32165478 25896374 --an 2024
```

Sau în limbaj natural:

> Analizează firma cu CUI 14186770 pentru încadrare IMM.

Plugin-ul va:
1. Confirma scope-ul la începutul fiecărui task (CUI-uri, program/apel, an + curs BNR ultima zi a anului, livrabile, semnatar Anexa 3, folder output).
2. Scrape termene.ro pentru solicitant + asociați + administratori (AMBELE secțiuni: „Asociați/acționari" și „Persoane autorizate") + alte firme ale acestora.
3. Aplica arborele decizional (Legea 346/2004 + Recomandarea CE 2003/361).
4. Întreba clarificări pentru cazuri ambigue (rude, piață învecinată, investitori-excepție).
5. Completa automat Declarație IMM + Anexa 3 (semnatarul lângă eticheta „Numele") + Anexa 4.
6. Genera workbook-ul „Analiza încadrare IMM" (livrabil obligatoriu): harta grupului, verdict, financiare în lei/euro/mii lei/mii euro, categoria IMM, Recomandare Claude.
7. Produce o sinteză 1-3 pagini cu verdict.

Output-ul ajunge în `<cwd>/analize-imm/<YYYY-MM-DD>_<CUI-solicitant>/`.

## Confidențialitate

- Toate datele rămân local pe mașina utilizatorului.
- Pluginul nu trimite date către niciun server extern (în afara cererilor inițiate de tine prin Chrome MCP către termene.ro).
- Template-urile bundled sunt READ-ONLY; pluginul scrie doar copii noi în folderul output.
- Audit log per analiză.

## Suport

Contact: ecofin@consultant.com (Paul Petrov, RomActiv Business Consulting).

## Licență

MIT — vezi `LICENSE`.
