# termene.ro — selectori DOM + URL patterns

Pentru utilizare cu `mcp__Claude_in_Chrome__navigate` + `get_page_text` / `find`.

## URL patterns

| Tip | URL | Note |
|---|---|---|
| Pagina firmă | `https://termene.ro/firma/{CUI}` | CUI fără prefix "RO" |
| Pagina firmă alt | `https://termene.ro/firm/{CUI}` | Alias |
| Pagina PF (persoană fizică) | `https://termene.ro/persoana/{slug-nume}` | Necesită PartnerSCAN |
| Căutare | `https://termene.ro/cauta/{query}` | Fallback dacă URL direct eșuează |

## Detecție tier acces

Pentru a determina dacă utilizatorul are cont premium/PartnerSCAN, după login:

1. Navighează la `https://termene.ro/` și caută selector `[data-test="user-menu"]` sau `.account-link` sau text "Contul meu".
2. Navighează la o pagină firmă de test → caută selector pentru asociați detaliat:
   - **Premium**: `.shareholders-detail`, `.actionari-detaliu`, `.partneri-list`, sau secțiune `<div data-section="shareholders">` cu tabel complet.
   - **Public**: doar listă nume asociați, fără %, fără cross-reference.

## Secțiuni de extras din pagina firmă

### Identificare firmă
- `h1.company-name` → denumire
- `.company-details .cui` sau `[data-field="cui"]` → CUI verificat
- `.company-details .address` sau `[data-field="address"]` → adresă sediu
- `.company-details .status` → stare (Funcțională / Radiată / În insolvență)
- `.company-details .data-infiintare` → data înregistrării

### CAEN
- `.caen-main` sau `[data-section="caen"] .main-code` → cod + descriere CAEN principal
- `.caen-secondary li` sau `[data-section="caen"] .secondary-code li` → CAEN-uri secundare (poate fi >20)

### Reprezentant legal
- `.administrators-list .person` sau `[data-section="administratori"]`
- Câmpuri per administrator: nume, funcție (Administrator unic / Director general / Președinte CA / membru CA), data numirii.

### Date financiare (bilanț ANAF)
- `[data-section="bilant"]` sau `.financial-table`
- Pentru fiecare an disponibil (ultimii 3–5 ani):
  - `Cifra de afaceri` / `Cifră de afaceri netă` (în lei)
  - `Active totale` / `Total activ` (în lei)
  - `Număr mediu salariați` / `Salariați`
  - `Profit net` / `Pierdere`

**ATENȚIE**: pe termene.ro valorile pot fi în RON full sau în mii RON — verifică unitatea afișată în antet de coloană. Pentru Excel-ul nostru avem nevoie în **mii lei**.

### Asociați / Acționari
- `[data-section="actionari"]` sau `.shareholders-list`
- Per asociat:
  - Tip: PF sau PJ (uneori marcat `.badge-pf` / `.badge-pj`)
  - Nume / Denumire
  - CUI (dacă PJ) — link către `/firma/{CUI}`
  - Procent deținere (în %)
  - Țară (default: România)
  - Rol (Asociat unic / Asociat / Acționar)
- Dacă tier = PREMIUM și asociat = PF: există link `/persoana/{slug}` cu lista **alte firme**.

### Alte firme ale unui asociat PF (premium)
- `.persoana-firme-list .firma-entry`
- Per intrare: denumire, CUI, procent deținere, CAEN principal, rol.

## Anti-bot / cadență

- Cadență: 1.5–3.5s între navigări. Folosește JavaScript delay sau `mcp__computer-use__wait`.
- Max 30 navigări per sesiune. La depășire: opritor obligatoriu, comunică utilizatorului.
- Dacă apare Cloudflare interstitial (selector `.cf-challenge` sau text "Verifying you are human") → STOP, cere user să rezolve manual challenge-ul în browser-ul lui, apoi reia.
- Nu deschide pagini în paralel — strict secvențial.

## Recovery / erori

| Eroare | Acțiune |
|---|---|
| 404 / "Firmă inexistentă" | Reintreabă userul CUI corect |
| 403 / "Acces interzis" | Probă login, sau cere user să se autentifice |
| 429 / "Too Many Requests" | Pauză 5 min, reia. Dacă persistă: STOP sesiune. |
| Time-out (>30s pagină) | Retry 1x; dacă persistă: marchează firma cu `scrape_status="timeout"` |
| Date bilanț lipsă pentru an de referință | Folosește an N-1, warning în sinteză |

## Schema JSON extragere (per firmă)

Vezi `scripts/scrape_termene.py` pentru schema exactă serializată. Câmpurile minime:

```json
{
  "cui": "RO12345678",
  "denumire": "...",
  "adresa": "...",
  "stare": "Funcțională",
  "caen_principal": {"cod": "6201", "label": "..."},
  "caen_secundar": [{"cod": "...", "label": "..."}],
  "reprezentant_legal": [{"nume": "...", "functie": "...", "data_numire": "..."}],
  "financiare": [
    {"an": 2024, "salariati": 12, "cifra_afaceri_lei": 4523187.55, "active_totale_lei": 2891044.10}
  ],
  "asociati": [
    {"tip": "PF|PJ", "nume": "...", "cui": "RO...", "procent": 60.0, "tara": "RO", "rol": "...",
     "alte_firme": [{"cui": "...", "denumire": "...", "procent_in": 30.0, "caen": "..."}]}
  ],
  "scraped_at": "2026-05-28T10:32:11+03:00",
  "source_tier": "premium|public",
  "warnings": []
}
```
