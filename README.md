# Chef Abbys × Côte d'Ivoire — AToure

AToure is delivering the experience and logistics for the **Côte d'Ivoire stop of Chef Abbys' West Africa Food Tour**, 19–21 October 2026.

| Folder | What's inside |
|---|---|
| [`01-research/chef-abbys-profile.md`](01-research/chef-abbys-profile.md) | Who she is, her tours, what she has already filmed in Abidjan, Ivorian angles that are current in 2026, people to approach, practical facts, sources |
| [`02-proposition/experience-proposition.md`](02-proposition/experience-proposition.md) | **The Wednesday draft**: *"Who Feeds Côte d'Ivoire?"*, a 3-day programme with budget, sponsorship approach, timeline and open questions |
| [`02-proposition/AToure-ChefAbbys-CotedIvoire-Proposition-DRAFT.pdf`](02-proposition/AToure-ChefAbbys-CotedIvoire-Proposition-DRAFT.pdf) | The same proposition as a branded 5-page PDF, ready to send to her management |
| [`03-budget/ground-budget.xlsx`](03-budget/ground-budget.xlsx) | **v0.2** live budget for the merged programme (team Parcours + Olivia Yacé goûter + Jiaan Wu maquis; 4 nights). Core ≈ $27.5k. Change the headcount, nights, exchange rate and fee % on *Assumptions*, and switch lines on or off (Y/N) on *Budget* |
| [`04-collaborations/foundation-and-creators.md`](04-collaborations/foundation-and-creators.md) | Working notes: Olivia Yacé foundation masterclass, Jiaan Wu at the maquis night, ministry patronage |
| `tools/` | Sources for rebuilding the files: `build_budget.py` regenerates the spreadsheet, and `proposition.html` is printed to the PDF (it needs the `atoure` brand-kit repo checked out alongside this one) |

## Status

- **v0.2 deck** (merged programme) is the current version, as a Slides artifact. The v0.1 Markdown/PDF is kept for reference.
- **Draft v0.1.** Every partner, person and venue named is a *target*. Nothing has been approached or confirmed yet.
- Budget figures are **estimates before supplier quotes**, for a party of 6. The AToure fee is set to 0% until it has been discussed with the client. Logistics-only total: 3,789,500 XOF (≈ $6,707).
- Credential: use **TIME100 Creators 2025** (it is pinned on her Instagram). I found no Forbes listing.

## Rebuild

```bash
python3 tools/build_budget.py                      # → 03-budget/ground-budget.xlsx
chromium --headless --no-sandbox --no-pdf-header-footer \
  --print-to-pdf=02-proposition/AToure-ChefAbbys-CotedIvoire-Proposition-DRAFT.pdf \
  tools/proposition.html
```
