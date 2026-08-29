# Phase 1 Outputs Summary

> **Created**: 2026-08-29 (Session 1)
> **Purpose**: Handoff document for Phase 2 (slide building, narrative finalisation, Q&A answers, rehearsal)
> **Repo**: `~/Grad-Challenge/` (GitHub: lukemoltbot/Grad-Challenge)

## Phase 1 Deliverables — 8 Analysis Documents (3,993 lines total)

All files in `~/Grad-Challenge/analysis/`:

| File | Lines | Workstream | Content |
|------|-------|------------|---------|
| `01_workbook_audit.md` | 550 | WS1 Task 1.1 | Full 15-sheet Complex Valuation Model audit. Cell counts, cross-sheet links, Brave Blossom Capital deep-dive. **Corrected**: unit costs sum to $387.9M (not $589M) — matches vault file 07 within $1M. |
| `01b_financial_model_analysis.md` | 350 | WS1 Tasks 1.2–1.11 | NPV/IRR scenarios, populated capital schedule, sensitivity analysis, DCF dashboard. **Populated workbook**: `workbooks/Complex_Valuation_Model_POPULATED.xlsx` |
| `02_brave_blossom_swot.md` | 344 | WS2 Tasks 2.1–2.15 | 11-area SWOT + valuation narrative + $389M capital review + 1:2.31 capital-to-liability ratio |
| `03_closure_liability_review.md` | 511 | WS3 Tasks 3.1–3.16 | $900M review, 8 anomalies, 6 SMART targets with compliance checks, cashflow/risk |
| `04_other_projects_and_workplan.md` | 786 | WS4 Tasks 4.1–4.11 | Brumby/Bronco/Option E SWOTs, cross-option matrix A–E, stage-gated workplan |
| `05_recommendation_gonogo.md` | 360 | WS5 Tasks 5.1–5.10 | Do-nothing scenario, 3 recommendations with cashflow + risks, timeline |
| `06_slide_structure_outline.md` | 564 | WS6 Task 6.1 | 14-slide outline (900s total) + 6 appendix slides, speaker notes, timing |
| `07_qa_preparation.md` | 528 | WS7 Tasks 7.1, 7.5 | 64 board questions, 10 model answers, 8 hostile question prep |

## Key Corrected Figures (use these in Phase 2 — NOT vault originals)

### Financial Model — Corrected with Capital Schedule

| Metric | Original (Workbook) | Corrected (Phase 1) | Notes |
|--------|---------------------|---------------------|-------|
| Brave Blossom NPV | $1,320M | **$997M** | Base case: 30% contingency + tax shield |
| Combined NPV | $1,670M | **$1,277M** | Includes capital costs |
| Brave Blossom IRR | 77.4% | **~52%** (estimated) | |
| Scenario range | — | **$859M–$1,072M** | All scenarios positive |
| Total value with deferral | — | **~$1,505M** | NPV + closure deferral benefit |

### Capital Schedule

| Metric | Value | Notes |
|--------|-------|-------|
| Unit costs sum | $387.9M | Matches vault file 07 ($388.9M) within $1M |
| Schedule total (no contingency) | $518.4M | Higher due to fractional unit counts |
| With 30% contingency | $673.9M | Full risked capital estimate |
| With tax shield | $518.4M | Base case for NPV |
| **Board reporting figure** | **$388.9M** direct, $673.9M risked | Use range in presentation |

### Other Key Figures (unchanged from vault)

| Metric | Value |
|--------|-------|
| Closure liability | $900M ($494M direct + $173M contingency + $233M holding) |
| SMART reductions | $180.3M (20% of $900M) — 6 targets, all implementable by end 2027 |
| Closure deferral NPV | $508M (at 7% discount) |
| Capital-to-liability ratio | 1:2.31 ($1 invested defers $2.31 closure) |
| Coal price | $211.20/t (33% below Springbok) |
| Discount rate | 8% |
| Emissions | 20.1M tCO₂e above baseline (20 years) |

## Phase 2 Tasks (what remains)

From TASK_LIST.md, Phase 2 covers:
- **WS1 Tasks 1.12–1.15**: Finalise DCF dashboard, sensitivity tables, carbon/emissions summary
- **WS2 Tasks 2.16–2.20**: Cross-reference SWOT with financial model, finalise narratives
- **WS5 Tasks 5.11–5.15**: Finalise recommendation cashflows with corrected NPV figures
- **WS6 Tasks 6.2–6.10**: Build actual slide deck (PPTX), write speaker notes, timing
- **WS7 Tasks 7.2–7.10**: Write full Q&A model answers, appendix slides

## Blocked Items (3 tasks, WS8)
1. Conceptual Mining Study document — not provided
2. Seam thickness contours — only in figures, not extractable
3. Depth of cover values — only in figures, not extractable

## Verification
- 73/73 ad-hoc checks pass (workbook population, capital schedule, NPV scenarios, SMART totals)
- 19/19 ad-hoc checks pass (patched file corrections)
- All subagent verification: 25/25, 53/53, 8/8, 31/31, 20/20

## Python Scripts (in `analysis/`)
- `_populate_capital.py` — populates Brave Blossom Capital sheet from vault file 07
- `_compute_npv_scenarios.py` — computes NPV/IRR scenarios with populated capital
- Both execute cleanly (exit 0)

## Populated Workbook
- `workbooks/Complex_Valuation_Model_POPULATED.xlsx` — Brave Blossom Capital tab populated with item names, unit costs, and year-by-year scheduling from vault file 07
