# Division 296 CGT Reset — website

A static, single-page site for the Division 296 cost-base reset model: a
plain-English explainer framed around the reset **valuation date (30 June 2026)**
— which is *not* the election deadline; the election is made with the fund's
2026-27 return — and a **live in-browser calculator** that reproduces the workbook's headline
comparison (Div 296 tax *if no reset* vs *if elected to reset*, with the signed
difference, per-member split, and most-affected assets — including the
"reset trap" flag).

No backend, no build step, no tracking. All computation runs client-side.

## Files

| File | Purpose |
|------|---------|
| `index.html` | Page structure and copy |
| `styles.css` | Styling (echoes the workbook palette / sign-colour convention) |
| `calcs.js` | **1:1 port of `src/div296/calcs.py`** — the calc engine |
| `app.js` | UI: editable state → render, countdown |
| `tests/parity.test.js` | Pins the §12 acceptance numbers against `calcs.js` |

## Source of truth

`src/div296/calcs.py` (the Python module, tested by `tests/test_calcs.py`)
remains the canonical calc engine. `calcs.js` is a hand-maintained mirror —
**change it in lockstep with the Python.** `tests/parity.test.js` reproduces
the same §12 acceptance numbers the pytest suite asserts, so drift is caught:

```bash
node web/tests/parity.test.js
# → ✓ all 23 parity checks passed — calcs.js matches §12 acceptance numbers
```

## Run locally

It's plain static files (ES modules, so serve over HTTP rather than `file://`):

```bash
cd web
python -m http.server 8000
# → open http://localhost:8000
```

## Deploy

Any static host. For **GitHub Pages**, point Pages at this `web/` directory
(Settings → Pages → Deploy from branch → `/web`), or copy these files to the
Pages source you use. No environment variables or secrets required.

---

> **Illustrative only — not financial, tax, or legal advice.** Realised-basis,
> Year 1, 100% accumulation phase. Confirm against the final ATO method and a
> registered tax agent before relying on any figure.
