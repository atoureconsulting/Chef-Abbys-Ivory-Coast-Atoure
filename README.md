# Chef Abbys × Côte d'Ivoire — AToure

AToure is delivering the experience and logistics for the **Côte d'Ivoire stop of Chef Abbys' West Africa Food Tour**, 19–21 October 2026.

| Folder | What's inside |
|---|---|
| [`01-research/chef-abbys-profile.md`](01-research/chef-abbys-profile.md) | Who she is, her tours, what she has already filmed in Abidjan, Ivorian angles that are current in 2026, people to approach, practical facts, sources |
| [`02-proposition/experience-proposition.md`](02-proposition/experience-proposition.md) | **The Wednesday draft**: *"Who Feeds Côte d'Ivoire?"*, a 3-day programme with budget, sponsorship approach, timeline and open questions |
| [`02-proposition/AToure-ChefAbbys-CotedIvoire-Proposition-DRAFT.pdf`](02-proposition/AToure-ChefAbbys-CotedIvoire-Proposition-DRAFT.pdf) | The same proposition as a branded 5-page PDF, ready to send to her management |
| [`03-budget/ground-budget.xlsx`](03-budget/ground-budget.xlsx) | Live budget. Change the headcount, nights, exchange rate and fee % on *Assumptions*, and switch lines on or off (Y/N) on *Budget* |
| [`04-collaborations/foundation-and-creators.md`](04-collaborations/foundation-and-creators.md) | Working notes: Olivia Yacé foundation masterclass, Jiaan Wu at the maquis night, ministry patronage |
| `tools/` | Sources for rebuilding the files: `build_budget.py` regenerates the spreadsheet, and `proposition.html` is printed to the PDF (it needs the `atoure` brand-kit repo checked out alongside this one) |

## Status

- **Draft v0.1.** Every partner, person and venue named is a *target*. Nothing has been approached or confirmed yet.
- Budget figures are **estimates before supplier quotes**, for a party of 6. The AToure fee (15%) is a placeholder.
- One credential needs checking: the brief says "Forbes-recognised", but the public record confirms **TIME 100 Most Influential Creators (2025)**.

## Rebuild

```bash
python3 tools/build_budget.py                      # → 03-budget/ground-budget.xlsx
chromium --headless --no-sandbox --no-pdf-header-footer \
  --print-to-pdf=02-proposition/AToure-ChefAbbys-CotedIvoire-Proposition-DRAFT.pdf \
  tools/proposition.html
```
