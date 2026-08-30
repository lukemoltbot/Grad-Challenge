# 2026 Graduate Challenge — Wallaby Mining Board Presentation

> **Scenario**: Business strategy simulation in the Australian coal mining industry. You are the management team of **Wallaby Mining (WM)**, presenting to the WM Board of Directors.
>
> **Repo**: [github.com/lukemoltbot/Grad-Challenge](https://github.com/lukemoltbot/Grad-Challenge)

---

## 🎯 Presentation Deliverables — Quick Links

| Deliverable | Link | Notes |
|-------------|------|-------|
| **HTML Deck** (20 slides, presenter mode) | [slides/deck.html](slides/deck.html) | Open in browser — keyboard navigation, press **N** for presenter notes |
| **PowerPoint** (20 slides, 126 KB) | [slides/Wallaby_Mining_Board_Presentation.pptx](slides/Wallaby_Mining_Board_Presentation.pptx) | Native python-pptx charts, editable in PowerPoint |
| Speaker Notes (900s timed) | [slides/speaker_notes.md](slides/speaker_notes.md) | Per-slide notes with key phrases and timing |

---

## Table of Contents

1. [Start Here — Reading Guide](#start-here--reading-guide)
2. [Overview](#overview)
3. [Repository Structure](#repository-structure)
4. [Complete Repository Contents — Every File Explained](#complete-repository-contents--every-file-explained)
5. [Key Files](#key-files)
6. [The Brief](#the-brief)
7. [Analysis Process & Methodology](#analysis-process--methodology)
8. [Excel Workbook Content](#excel-workbook-content)
9. [Key Financial Figures](#key-financial-figures)
10. [Deliverables Mapping](#deliverables-mapping)
11. [Status](#status)
12. [Calculation Walkthroughs](#calculation-walkthroughs)

---

## Start Here — Reading Guide

**New to this project? Read these two documents first, in this order:**

1. [`analysis/08_plain_english_evaluation.md`](analysis/08_plain_english_evaluation.md) — **Plain-English Evaluation & Recommendations**

   Start here. This explains, in plain English, what each of the five options (A–E) means, what's good and bad about each, and why the team recommends what it recommends. No jargon, no spreadsheets — just clear reasoning. If you read only one document, read this one.

2. [`PLAIN_ENGLISH_GUIDE.md`](PLAIN_ENGLISH_GUIDE.md) — **Plain English Guide: How We Reached Our Conclusions**

   Read this second. This builds on the evaluation above and walks through *how* the conclusions were reached — the three big recommendations, how the financial numbers were calculated (NPV, IRR, capital), what the risks are, and what happens to the workers and the town. Includes a plain-English glossary of every acronym used in the project.

3. [`OPTION_A_CALCULATION_WALKTHROUGH.md`](OPTION_A_CALCULATION_WALKTHROUGH.md) — **Option A Calculation Walkthrough: Brave Blossom Underground Mine**

   Read this third. A cell-by-cell walkthrough of the entire Brave Blossom DCF model. Explains every financial term (discount rate, NPV, IRR, tax shield, carbon liability), references exact workbook cells (`Brave Blossom !D7`, `Assumptions!B2`), and shows how revenue, costs, capital, and cashflow combine to produce the corrected $997M NPV. Open the Excel workbook alongside this document and follow the live formulas.

4. [`OPTION_C_CALCULATION_WALKTHROUGH.md`](OPTION_C_CALCULATION_WALKTHROUGH.md) — **Option C Calculation Walkthrough: Phased Go — Stage-Gated Brave Blossom**

   Read this fourth. Covers what's unique to Option C: the three stage-gate decision framework overlaid on the same DCF model, the $20M initial commitment, gate pass criteria, real option value, and the separate closure-liability deferral NPV ($508M at 7% discount). Cross-references the shared DCF chain from the Option A walkthrough rather than repeating it.

After these four guides, the rest of this README and the `/analysis/` folder provide the full technical detail behind every figure and recommendation.

---

## Overview

WM acquired 60% of the Waratah Resources Joint Venture, gaining three assets:

| Asset | Status |
|-------|--------|
| **Springbok Coal Complex** | Operational underground coking coal mine, ~5 years from closure (Q4 2031) |
| **Brumby Open Cut** | Historic depleted mine + development project targeting Xanadu seam (1.5 Mtpa, 25 km NE) |
| **Brave Blossom Project** | Concept-stage underground, ~6 Mtpa saleable HCC until ~2052, ~$389M capital |

The acquisition was conditional on **assuming closure liability from 2032 onwards** ($900M rehabilitation estimate). In return, WM receives 60% of revenue from pre-existing infrastructure (CHPP, TLO, rail/port contracts).

### Two Core Tasks

| Task | Description |
|------|-------------|
| **Task 1 — Strategic Review** | Evaluate all life-extension options (Brave Blossom, Brumby, Bronco acquisition, future exploration, or no-go). Present SWOT across 11 areas for each. Recommend go/no-go with forward work plan timeline. |
| **Task 2 — Closure Liability** | Evaluate the $900M closure estimate and post-mining carbon liability. Present SMART-aligned recommendations to reduce rehabilitation and emissions liabilities. All targets implementable by end of 2027. |

### Deliverable Format

- **15-minute** board presentation (max) + **5-minute** Q&A
- Professional, creative, influential — a board pitch, not a technical report
- Four required sections: (1) Brave Blossom valuation & SWOT, (2) Post-mining liability reduction, (3) Other projects & forward workplan, (4) Go/no-go recommendation & timeline
- Top 3 recommendations, each with cashflow analysis and implementation risks

---

## Repository Structure

```
Grad-Challenge/
├── 2026_Graduate_Challenge_Brief.pdf      # Original 18-page challenge brief
├── TASK_LIST.md                             # Master task list (95 tasks across 8 workstreams)
├── README.md                                # This file
├── PLAIN_ENGLISH_GUIDE.md                   # Plain English Guide: How We Reached Our Conclusions
├── OPTION_A_CALCULATION_WALKTHROUGH.md       # Cell-by-cell DCF walkthrough (Brave Blossom, $997M NPV)
├── OPTION_C_CALCULATION_WALKTHROUGH.md       # Stage-gated phased go walkthrough (gates, deferral NPV)
├── vault/                                   # 23-file knowledge base (INDEX + 01–23)
│   ├── INDEX.md
│   ├── 01_Challenge_Overview.md
│   ├── 02_Company_Profile.md
│   ├── ...
│   └── 23_Phase1_Outputs.md
├── workbooks/                               # Excel financial models
│   ├── Complex_Valuation_Model.xlsx         # Original (15 sheets, as provided)
│   ├── Complex_Valuation_Model_POPULATED.xlsx  # Populated with Brave Blossom capital schedule
│   └── 2024_Springbok_Planned_Closure_Costs.xlsx  # $900M ERC closure cost model (6 sheets)
├── analysis/                                # 9 analysis documents + Python scripts
│   ├── 01_workbook_audit.md
│   ├── 01b_financial_model_analysis.md
│   ├── 02_brave_blossom_swot.md
│   ├── 03_closure_liability_review.md
│   ├── 04_other_projects_and_workplan.md
│   ├── 05_recommendation_gonogo.md
│   ├── 06_slide_structure_outline.md
│   ├── 07_qa_preparation.md
│   ├── 08_plain_english_evaluation.md        # Plain-English evaluation of all 5 options (START HERE)
│   ├── _audit_*.py                          # Workbook audit scripts
│   ├── _populate_capital.py                 # Populates Brave Blossom Capital tab
│   └── _compute_npv_scenarios.py            # NPV/IRR scenario calculator
├── slides/                                  # Presentation deliverables
│   ├── Wallaby_Mining_Board_Presentation.pptx  # 20-slide PowerPoint deck
│   ├── deck.html                               # 20-slide HTML deck (presenter mode)
│   ├── speaker_notes.md                         # 900s timed speaker notes
│   └── _build_pptx.py                           # PPTX build script
└── verify_phase2.py                         # 55-check verification suite
```

---

## Complete Repository Contents — Every File Explained

> Every file in this repository, grouped by directory, with a one-line purpose so you know what each document is for before opening it. Use `Ctrl+F` to search for a filename or keyword.

### Root Directory (`/`)

| File | Purpose |
|------|---------|
| [`README.md`](README.md) | **This file.** Project overview, reading guide, methodology, financial figures, and complete file index. Start here. |
| [`EXECUTIVE_SUMMARY.md`](EXECUTIVE_SUMMARY.md) | **Executive summary** — the entire project distilled to one page. The situation, the five options, the three recommendations, key financials, risks, timeline, and bottom line. If you have 2 minutes, read this. |
| [`PLAIN_ENGLISH_GUIDE.md`](PLAIN_ENGLISH_GUIDE.md) | **Plain English Guide: How We Reached Our Conclusions.** Walks through the three recommendations, how the financial numbers were calculated (NPV, IRR, capital), what the risks are, and what happens to workers and the town. Includes a plain-English glossary of every acronym. Read after the executive summary. |
| [`OPTION_A_CALCULATION_WALKTHROUGH.md`](OPTION_A_CALCULATION_WALKTHROUGH.md) | **Cell-by-cell DCF walkthrough for Brave Blossom (Option A).** Explains every financial term (discount rate, NPV, IRR, tax shield, carbon liability) with exact workbook cell references. Shows how revenue, costs, capital, and cashflow combine to produce the corrected **$997M NPV** (~52% IRR). 826 lines. Open the Excel workbook alongside this guide and follow the live formulas. |
| [`OPTION_C_CALCULATION_WALKTHROUGH.md`](OPTION_C_CALCULATION_WALKTHROUGH.md) | **Stage-gated phased go walkthrough (Option C).** Covers the three stage-gate decision framework, $20M initial commitment, gate pass criteria, real option value, and the separate **closure-liability deferral NPV of $508M** (at 7% discount). Cross-references the shared DCF chain from the Option A walkthrough. 658 lines. |
| [`TASK_LIST.md`](TASK_LIST.md) | **Master task list** — 95 tasks across 8 workstreams, classified by readiness (✅ READY / 🟡 DRAFT-READY / 🔴 BLOCKED). Tracks the full pipeline from analysis through presentation. |
| [`2026_Graduate_Challenge_Brief.pdf`](2026_Graduate_Challenge_Brief.pdf) | **Original challenge brief** — 18-page document (+ 8-page supplement by Shaun Booth) defining the scenario, the two tasks, deliverable format, and provided financial models. The source of truth for all requirements. |
| [`verify_phase2.py`](verify_phase2.py) | **55-check verification suite** — programmatically verifies all Phase 2 deliverables (slide counts, NPV figures, file presence, content checks). All 55 checks pass. |

### `slides/` — Presentation Deliverables

| File | Purpose |
|------|---------|
| [`slides/deck.html`](slides/deck.html) | **Primary HTML presentation deck** — 20 slides with presenter mode, CSS charts, keyboard navigation, and presenter notes panel (press **N**). This is the main deck for live presentation. 1,446 lines. |
| [`slides/Wallaby_Mining_Board_Presentation.pptx`](slides/Wallaby_Mining_Board_Presentation.pptx) | **PowerPoint deck** — 20 slides (126 KB) with native python-pptx charts (bar, waterfall, timeline). Editable in PowerPoint for manual adjustments. |
| [`slides/presentation-annotated.html`](slides/presentation-annotated.html) | **Annotated HTML deck** — same 20 slides as the main deck, but with interactive tooltips. Hover any dotted-underlined value to see calculation details, source cell references, and explanatory notes. Use for self-review or sharing with reviewers who want to verify figures. 926 lines. |
| [`slides/presentation.html`](slides/presentation.html) | **Simplified HTML deck** — earlier non-presenter version of the 20-slide deck. Simpler layout without presenter mode. Useful for quick viewing or embedding. 793 lines. |
| [`slides/speaker_notes.md`](slides/speaker_notes.md) | **Speaker notes** — 900-second timed notes for all 14 content slides + 6 appendix slides. Per-slide key phrases, transitions, and timing markers. |
| [`slides/_build_pptx.py`](slides/_build_pptx.py) | **PPTX build script** — generates the PowerPoint deck programmatically using python-pptx. Re-run to regenerate the PPTX from source data. |

### `analysis/` — Analysis Documents & Scripts

| File | Purpose |
|------|---------|
| [`analysis/01_workbook_audit.md`](analysis/01_workbook_audit.md) | **Full 15-sheet Complex Valuation Model audit.** Cell counts, formula counts, cross-sheet link verification. Key finding: Brave Blossom Capital tab was a blank template shell. Unit costs verified at $387.9M. 550 lines. |
| [`analysis/01b_financial_model_analysis.md`](analysis/01b_financial_model_analysis.md) | **NPV/IRR scenario analysis** with populated capital schedule. Corrected NPV: $997M base case. Five scenarios ($859M–$1,072M), sensitivity analysis, DCF dashboard. 350 lines. |
| [`analysis/02_brave_blossom_swot.md`](analysis/02_brave_blossom_swot.md) | **11-area SWOT for Brave Blossom** (Option A: Full Go). Valuation narrative, $389M capital review, 1:2.31 capital-to-liability ratio analysis. 344 lines. |
| [`analysis/03_closure_liability_review.md`](analysis/03_closure_liability_review.md) | **$900M closure liability review.** All 8 anomalies documented, 6 SMART targets with compliance checks (total: $180.3M / 20% reduction), cashflow impact and implementation risk. 511 lines. |
| [`analysis/04_other_projects_and_workplan.md`](analysis/04_other_projects_and_workplan.md) | **Other projects & forward workplan.** SWOT for Brumby (Option C), Bronco (Option D), and future exploration (Option E). Cross-option comparison matrix A–E. Stage-gated workplan 2027–2033. 787 lines. |
| [`analysis/05_recommendation_gonogo.md`](analysis/05_recommendation_gonogo.md) | **Go/no-go recommendation.** Do-nothing scenario (Option B), top 3 recommendations with cashflow analysis and implementation risks, proposed timeline. 361 lines. |
| [`analysis/06_slide_structure_outline.md`](analysis/06_slide_structure_outline.md) | **Slide structure outline** — 14 content slides + 6 appendix, 900-second timing budget, per-slide content and speaker notes guidance. 564 lines. |
| [`analysis/07_qa_preparation.md`](analysis/07_qa_preparation.md) | **Q&A preparation** — 64 anticipated board questions across 8 categories, 10 model answers, 8 hostile question prep, solo presenter framework. 527 lines. |
| [`analysis/08_plain_english_evaluation.md`](analysis/08_plain_english_evaluation.md) | **Plain-English evaluation of all 5 options (A–E)** — what each means, pros, cons, and why the team recommends what it recommends. No jargon, no spreadsheets. **Start here.** 434 lines. |
| [`analysis/_audit_deep.py`](analysis/_audit_deep.py) | **Deep audit script** — dumps cell-level formulas and values for OC Clean Capital, Springbok, Brave Blossom, and Decommissioned Mine sheets. Used to generate the workbook audit report (01). |
| [`analysis/_audit_raw.py`](analysis/_audit_raw.py) | **Raw audit script** — comprehensive dump of all 15 sheets: every cell with formulas, values, and structured JSON output. Most thorough audit script. |
| [`analysis/_audit_structure.py`](analysis/_audit_structure.py) | **Structural audit script** — analyses sheet structure: dimensions, formula patterns, cross-sheet link mapping. Used to map the workbook architecture. |
| [`analysis/_audit_targeted.py`](analysis/_audit_targeted.py) | **Targeted audit script** — extracts specific data: Safeguard baselines, NPV result cells, Summary tab, Assumptions tab. Used for final figure verification. |
| [`analysis/_populate_capital.py`](analysis/_populate_capital.py) | **Capital schedule populator** — writes 12 capital items with unit costs and year-by-year fractional scheduling into the Brave Blossom Capital tab, producing the populated workbook. |
| [`analysis/_compute_npv_scenarios.py`](analysis/_compute_npv_scenarios.py) | **NPV/IRR scenario calculator** — computes 5 NPV scenarios (original, with/without contingency, with/without tax shield) from the populated workbook. Produces the corrected $997M base case. |

### `vault/` — Structured Knowledge Base (23 Files + Index)

> The vault decomposes the 18-page brief + 8-page supplement into 23 structured topic files. Start with the index, then jump to the file for the topic you need.

| File | Topic | Purpose |
|------|-------|---------|
| [`vault/INDEX.md`](vault/INDEX.md) | **Navigation index** | Links to all 23 vault files with one-line summaries. Use this to find the right vault file quickly. |
| [`vault/01_Challenge_Overview.md`](vault/01_Challenge_Overview.md) | Challenge overview | What the challenge is, who you represent (WM management team), the two core tasks, and deliverable format. |
| [`vault/02_Company_Profile.md`](vault/02_Company_Profile.md) | Company profile | Wallaby Mining, Waratah Resources JV structure, 60% acquisition terms, asset portfolio summary. |
| [`vault/03_Springbok_Operations.md`](vault/03_Springbok_Operations.md) | Springbok operations | Current mine status, production, challenges, infrastructure (CHPP, TLO, rail/port), logistics, closure timeline (Q4 2031). |
| [`vault/04_Brave_Blossom_Project.md`](vault/04_Brave_Blossom_Project.md) | Brave Blossom project | Conceptual underground mine — geology, seam details, mine design, ~6 Mtpa saleable HCC, ~20-year life, capital estimate. |
| [`vault/05_Key_Considerations.md`](vault/05_Key_Considerations.md) | Key considerations | CHPP capacity & life, logistics (rail/port), carbon pricing, closure, social impact, neighbouring mines (Bronco). |
| [`vault/06_Coal_Quality_Specs.md`](vault/06_Coal_Quality_Specs.md) | Coal quality specs | Detailed coal quality parameters table — ash, VM, CSN, sulfur, phosphorus, crucible swelling, etc. for Springbok premium HCC. |
| [`vault/07_Capital_Estimate.md`](vault/07_Capital_Estimate.md) | Capital estimate | $389M concept capital estimate — 12 items broken down by year and category (development, equipment, infrastructure). Source for the populated capital schedule. |
| [`vault/08_Task1_Strategic_Review.md`](vault/08_Task1_Strategic_Review.md) | Task 1: Strategic review | Life extension strategic assessment — options A–E, SWOT framework, 11 evaluation areas, go/no-go criteria. |
| [`vault/09_Task2_Closure_Liability.md`](vault/09_Task2_Closure_Liability.md) | Task 2: Closure liability | Post-mining liability mitigation — $900M estimate, decommissioned mine emissions, SMART recommendation framework. |
| [`vault/10_Deliverables_Format.md`](vault/10_Deliverables_Format.md) | Deliverables format | Presentation requirements — 15+5 min format, four required sections, SWOT/SMART criteria, professional/creative standard. |
| [`vault/11_Outstanding_Information.md`](vault/11_Outstanding_Information.md) | Outstanding information | What's missing from the brief — gaps between what's provided and what's needed to fully complete the challenge (conceptual mining study, seam contours, depth of cover). |
| [`vault/12_Proposed_Solutions.md`](vault/12_Proposed_Solutions.md) | Proposed solutions | Strategic options matrix — Options A–E with capital, revenue extension, closure impact, and risk level comparison. |
| [`vault/13_SWOT_Framework.md`](vault/13_SWOT_Framework.md) | SWOT framework | Standardised SWOT structure for evaluating each option across the 11 areas specified in the brief. |
| [`vault/14_Deliverables_Checklist.md`](vault/14_Deliverables_Checklist.md) | Deliverables checklist | Master checklist consolidating all required presentation elements from both the 18-page brief and 8-page supplement. |
| [`vault/15_Brumby_Open_Cut.md`](vault/15_Brumby_Open_Cut.md) | Brumby open cut | Brumby open cut development project — Xanadu seam target, 1.5 Mtpa, 25 km NE. Reconciles dual status (historic mine + development project). |
| [`vault/16_Supplement_Notes.md`](vault/16_Supplement_Notes.md) | Supplement notes | 8-page supplemental PDF analysis — new details, corrections, and reconciliations relative to the 18-page brief. |
| [`vault/17_2024_Planned_Closure_Costs.md`](vault/17_2024_Planned_Closure_Costs.md) | Closure cost workbook | Full breakdown of the $900M ERC closure cost workbook — 6 sheets, 9 domains, 284 line items, holding costs, rate library, cost reduction opportunities. |
| [`vault/18_Audit_Results.md`](vault/18_Audit_Results.md) | Audit results | Calculation audit of all 284 closure cost line items — domain subtotals, grand total ($900M), contingency math, holding costs verified. **0 arithmetic errors.** |
| [`vault/19_Issues_and_Anomalies.md`](vault/19_Issues_and_Anomalies.md) | Issues & anomalies | 8 issues found during audit — duplicate TSF ($40M+), duplicate 4I/4J ($4.8M), missing execution-phase detail, outdated rate escalations, fleet rate substitution. |
| [`vault/20_Vault_Crossrefs_and_SMART.md`](vault/20_Vault_Crossrefs_and_SMART.md) | Cross-refs & SMART targets | Cross-references to vault files 05/07/08/09/12/13/14/15/16 + 6 quantified SMART cost-reduction targets (~$180M, 20%) + NPV of Brave Blossom deferral ($508M). |
| [`vault/21_Complex_Valuation_Model.md`](vault/21_Complex_Valuation_Model.md) | Valuation model analysis | Full analysis of the Complex Valuation Model workbook — 16 sheets, 3 mining scenarios (Springbok $279M, Brave Blossom $1,320M, combined $1,670M), carbon pricing, decommissioned mine CH₄ emissions. |
| [`vault/22_Requirements_and_Deliverables_Outline.md`](vault/22_Requirements_and_Deliverables_Outline.md) | Master reference | Consolidated requirements & deliverables from all 21 vault files — scenario, tasks, presentation structure, SWOT areas, SMART targets, financial models, key numbers, gaps, timeline, cross-ref map. |
| [`vault/23_Phase1_Outputs.md`](vault/23_Phase1_Outputs.md) | Phase 1 handoff | Phase 1 outputs summary — 8 analysis docs, corrected NPV figures ($997M/$1,277M), Phase 2 task list, verification status. Handoff document for Phase 2. |

### `workbooks/` — Excel Financial Models

| File | Purpose |
|------|---------|
| [`workbooks/Complex_Valuation_Model.xlsx`](workbooks/Complex_Valuation_Model.xlsx) | **Original DCF valuation model** — 15 sheets as provided by challenge organisers. Contains the blank Brave Blossom Capital tab that produced the $0 capital / $1,320M NPV error. |
| [`workbooks/Complex_Valuation_Model_POPULATED.xlsx`](workbooks/Complex_Valuation_Model_POPULATED.xlsx) | **Populated DCF model** — Brave Blossom Capital tab completed with 12 items and year-by-year scheduling. Produces the corrected $997M NPV. Use this for all analysis. ⚠️ No cached formula values — use `_compute_npv_scenarios.py` for computed results. |
| [`workbooks/2024_Springbok_Planned_Closure_Costs.xlsx`](workbooks/2024_Springbok_Planned_Closure_Costs.xlsx) | **$900M ERC closure cost model** — 6 sheets, 9 rehabilitation domains, 284 line items, 296 unit rates. The third-party closure estimate that the SMART targets challenge. |

---

## Key Files

### Plain-English Guides (Read These First)

| File | Description |
|------|-------------|
| [`analysis/08_plain_english_evaluation.md`](analysis/08_plain_english_evaluation.md) | **Plain-English Evaluation & Recommendations** — Start here. Explains each of the five options (A–E) in plain English: what it is, what's good, what's bad, and why. No jargon, no spreadsheets. |
|| [`PLAIN_ENGLISH_GUIDE.md`](PLAIN_ENGLISH_GUIDE.md) | **Plain English Guide: How We Reached Our Conclusions** — Read second. Walks through how the recommendations were reached, how the financial numbers were calculated, what the risks are, and what happens to workers and the town. Includes a plain-English glossary of all acronyms. |

### Calculation Walkthroughs

These documents walk through the financial calculations cell-by-cell, referencing exact workbook cells so readers can follow the live Excel formulas. They are the most technically detailed guides in the project — read them after the plain-English guides above, with the workbook open alongside.

|| File | Description |
|------|-------------|
|| [`OPTION_A_CALCULATION_WALKTHROUGH.md`](OPTION_A_CALCULATION_WALKTHROUGH.md) | **Option A: Brave Blossom Underground Mine** — Full DCF walkthrough. Explains discount rate, NPV, IRR, tax shield, revenue/cost/capital chains, carbon liability, and how they combine to produce the corrected **$997M NPV** (~52% IRR). 826 lines with exact cell references to `Brave Blossom`, `Assumptions`, `Brave Blossom Capital`, `Carbon`, `Decommissioned Mine`, and `Analysis` sheets. |
|| [`OPTION_C_CALCULATION_WALKTHROUGH.md`](OPTION_C_CALCULATION_WALKTHROUGH.md) | **Option C: Phased Go — Stage-Gated Brave Blossom** — Focuses on what's unique to Option C: three stage gates (end 2028, end 2030, 2032/33), $20M initial PFS commitment, gate pass criteria, real option value, and the **closure-liability deferral NPV of $508M** (at 7% discount). Cross-references the shared DCF chain from Option A. 658 lines. |

### Brief & Planning

| File | Description |
|------|-------------|
| [`2026_Graduate_Challenge_Brief.pdf`](2026_Graduate_Challenge_Brief.pdf) | Original 18-page challenge brief + 8-page supplement by Shaun Booth |
| [`TASK_LIST.md`](TASK_LIST.md) | Master task list — 95 tasks across 8 workstreams, classified by readiness (READY/DRAFT-READY/BLOCKED) |

### Vault (Knowledge Base)

The vault is a 23-file structured knowledge base that decomposes the brief into navigable topics. Start with the index:

| File | Topic |
|------|-------|
| [`vault/INDEX.md`](vault/INDEX.md) | Navigation index — links to all 23 vault files with summaries |
| [`vault/01_Challenge_Overview.md`](vault/01_Challenge_Overview.md) | What the challenge is, who you represent, the two core tasks |
| [`vault/04_Brave_Blossom_Project.md`](vault/04_Brave_Blossom_Project.md) | Conceptual underground project — geology, design, capital |
| [`vault/07_Capital_Estimate.md`](vault/07_Capital_Estimate.md) | $389M concept capital estimate breakdown by year and item |
| [`vault/17_2024_Planned_Closure_Costs.md`](vault/17_2024_Planned_Closure_Costs.md) | Full $900M ERC closure cost workbook breakdown |
| [`vault/18_Audit_Results.md`](vault/18_Audit_Results.md) | 284 line items verified, 0 arithmetic errors |
| [`vault/19_Issues_and_Anomalies.md`](vault/19_Issues_and_Anomalies.md) | 8 issues/anomalies identified (duplicate TSF, contingency, etc.) |
| [`vault/21_Complex_Valuation_Model.md`](vault/21_Complex_Valuation_Model.md) | Full valuation model analysis — 15 sheets, 3 scenarios, carbon pricing |
| [`vault/22_Requirements_and_Deliverables_Outline.md`](vault/22_Requirements_and_Deliverables_Outline.md) | Master reference — consolidated requirements & deliverables |
| [`vault/23_Phase1_Outputs.md`](vault/23_Phase1_Outputs.md) | Phase 1 handoff — corrected NPV figures, Phase 2 task list |

### Analysis Documents

| File | Lines | Content |
|------|-------|---------|
| [`analysis/01_workbook_audit.md`](analysis/01_workbook_audit.md) | 550 | Full 15-sheet Complex Valuation Model audit. Cell counts, cross-sheet links, Brave Blossom Capital deep-dive. Unit costs verified at $387.9M. |
| [`analysis/01b_financial_model_analysis.md`](analysis/01b_financial_model_analysis.md) | 350 | NPV/IRR scenarios with populated capital schedule. Corrected NPV: $997M base case. Sensitivity analysis, DCF dashboard. |
| [`analysis/02_brave_blossom_swot.md`](analysis/02_brave_blossom_swot.md) | 344 | 11-area SWOT + valuation narrative + $389M capital review + 1:2.31 capital-to-liability ratio |
| [`analysis/03_closure_liability_review.md`](analysis/03_closure_liability_review.md) | 511 | $900M review, 8 anomalies, 6 SMART targets with compliance checks, cashflow/risk analysis |
| [`analysis/04_other_projects_and_workplan.md`](analysis/04_other_projects_and_workplan.md) | 787 | Brumby/Bronco/Option E SWOTs, cross-option matrix A–E, stage-gated workplan (2027–2033) |
| [`analysis/05_recommendation_gonogo.md`](analysis/05_recommendation_gonogo.md) | 361 | Do-nothing scenario, 3 recommendations with cashflow + risks, timeline |
| [`analysis/06_slide_structure_outline.md`](analysis/06_slide_structure_outline.md) | 564 | 14-slide outline (900s total) + 6 appendix slides, timing, speaker notes |
| [`analysis/07_qa_preparation.md`](analysis/07_qa_preparation.md) | 527 | 64 board questions, 10 model answers, 8 hostile question prep, solo presenter framework |
| [`analysis/08_plain_english_evaluation.md`](analysis/08_plain_english_evaluation.md) | 434 | Plain-English evaluation of all 5 options (A–E): what each means, pros/cons, and why. No jargon. **Start here.** |

### Excel Workbooks

| File | Description |
|------|-------------|
| [`workbooks/Complex_Valuation_Model.xlsx`](workbooks/Complex_Valuation_Model.xlsx) | Original 15-sheet DCF valuation model (as provided by challenge organisers) |
| [`workbooks/Complex_Valuation_Model_POPULATED.xlsx`](workbooks/Complex_Valuation_Model_POPULATED.xlsx) | Populated version — Brave Blossom Capital tab completed with 12 items and year-by-year scheduling |
| [`workbooks/2024_Springbok_Planned_Closure_Costs.xlsx`](workbooks/2024_Springbok_Planned_Closure_Costs.xlsx) | $900M ERC closure cost model — 6 sheets, 9 domains, 284 line items |

### Presentation Deliverables

| File | Description |
|------|-------------|
| [`slides/Wallaby_Mining_Board_Presentation.pptx`](slides/Wallaby_Mining_Board_Presentation.pptx) | 20-slide PowerPoint deck (126 KB) — native python-pptx charts |
| [`slides/deck.html`](slides/deck.html) | 20-slide HTML deck — presenter mode, CSS charts, keyboard nav, presenter notes panel (press N) |
| [`slides/speaker_notes.md`](slides/speaker_notes.md) | 900s timed speaker notes — 14 content slides + 6 appendix, per-slide notes with key phrases |

### Python Scripts

| File | Purpose |
|------|---------|
| [`analysis/_populate_capital.py`](analysis/_populate_capital.py) | Populates Brave Blossom Capital sheet from vault file 07 capital estimate |
| [`analysis/_compute_npv_scenarios.py`](analysis/_compute_npv_scenarios.py) | Computes NPV/IRR scenarios with populated capital schedule |
| [`analysis/_audit_*.py`](analysis/) | Four audit scripts (deep, raw, structure, targeted) for workbook verification |
| [`slides/_build_pptx.py`](slides/_build_pptx.py) | Builds the PPTX presentation deck programmatically |
| [`verify_phase2.py`](verify_phase2.py) | 55-check verification suite — all checks pass |

---

## The Brief

The [2026 Graduate Challenge Brief](2026_Graduate_Challenge_Brief.pdf) is an 18-page document (plus an 8-page supplement by Shaun Booth, 30 June 2026) that sets out a business strategy simulation for Wallaby Mining. The brief defines:

- **The scenario**: WM's acquisition of 60% of the Waratah Resources JV, including the Springbok mine, Brumby open cut, and Brave Blossom project
- **Task 1**: Strategic assessment of life-extension options — evaluate Brave Blossom, Brumby, Bronco acquisition, future exploration, and the no-go option. SWOT across 11 areas per option. Go/no-go recommendation with forward work plan.
- **Task 2**: Evaluate the $900M closure estimate and post-mining carbon liability. SMART-aligned recommendations to reduce both. Implementation by end of 2027.
- **Deliverable format**: 15-minute presentation + 5-minute Q&A to the WM Board. Four sections, top 3 recommendations, professional and creative.
- **Financial models provided**: Complex Valuation Model (15-sheet DCF workbook) and 2024 Planned Closure Cost workbook ($900M ERC model, 6 sheets)

The vault files ([`vault/01`](vault/01_Challenge_Overview.md) through [`vault/22`](vault/22_Requirements_and_Deliverables_Outline.md)) decompose the brief into 22 structured topic files. File 22 is the master consolidated reference that maps every requirement to its corresponding vault file and deliverable.

---

## Analysis Process & Methodology

The analysis was executed in a structured, phase-gated pipeline designed to ensure every brief requirement was addressed with traceable evidence from the source documents and Excel workbooks.

### Phase 1 — Analysis (Complete)

**Step 1: Brief Decomposition (Vault Construction)**

The 18-page brief + 8-page supplement was read, decomposed, and restructured into 22 vault files — one per topic (company profile, operations, project details, coal quality, capital estimate, closure costs, SWOT framework, deliverables checklist, etc.). This ensured no requirement was missed and every number could be traced back to its source. Vault file 22 ([`22_Requirements_and_Deliverables_Outline.md`](vault/22_Requirements_and_Deliverables_Outline.md)) consolidates all requirements into a single master reference with a cross-reference map.

**Step 2: Excel Workbook Audit**

Both provided workbooks were programmatically audited using Python (openpyxl 3.1.5, pandas):

- **Complex Valuation Model**: All 15 sheets catalogued with cell counts, formula counts, and cross-sheet link verification. The critical finding was that the **Brave Blossom Capital tab** (sheet 10) had a fully-built scheduling framework (882 cells, formulas, year headers) but **zero populated scheduling rows** — producing $0 capital and inflating NPV. Full audit in [`analysis/01_workbook_audit.md`](analysis/01_workbook_audit.md).
- **2024 Planned Closure Cost Workbook**: All 284 line items across 9 domains verified. **0 arithmetic errors** found. 8 issues/anomalies identified (duplicate TSF costing, duplicate 4I/4J, outdated rate escalations, etc.). Results in [`vault/18_Audit_Results.md`](vault/18_Audit_Results.md) and [`vault/19_Issues_and_Anomalies.md`](vault/19_Issues_and_Anomalies.md).

**Step 3: Capital Schedule Population & NPV Correction**

The Brave Blossom Capital tab was populated using the 12 capital items and year-by-year scheduling from vault file 07 (the $389M concept estimate). The script [`_populate_capital.py`](analysis/_populate_capital.py) wrote item names, unit costs, and fractional unit counts into the workbook, producing [`Complex_Valuation_Model_POPULATED.xlsx`](workbooks/Complex_Valuation_Model_POPULATED.xlsx).

With real capital costs flowing through the model, NPV was recomputed across 5 scenarios using [`_compute_npv_scenarios.py`](analysis/_compute_npv_scenarios.py):

| Scenario | NPV (AUD M) | vs Original |
|----------|------------|-------------|
| Original (no capital — workbook as-is) | $1,320M | — |
| With capital, no contingency, tax shield | $1,072M | -19% |
| **With capital, 30% contingency, tax shield (base case)** | **$997M** | **-24%** |
| With capital, no contingency, no tax shield | $965M | -27% |
| With capital, 30% contingency, no tax shield | $859M | -35% |

**Key conclusion**: Brave Blossom remains strongly NPV-positive in all scenarios. The original $1,320M was computed with $0 capital — a material overstatement. Full analysis in [`analysis/01b_financial_model_analysis.md`](analysis/01b_financial_model_analysis.md).

**Step 4: SWOT, Closure Review, and Recommendations**

With corrected financials in hand, the analysis documents were produced:

- **02_brave_blossom_swot.md** — Full 11-area SWOT for Brave Blossom (Option A: Full Go), plus valuation narrative and capital-to-liability ratio analysis (1:2.31 — every $1 invested defers $2.31 of closure liability)
- **03_closure_liability_review.md** — Formal review of the $900M estimate, all 8 anomalies documented, 6 SMART targets with compliance checks (total: $180.3M / 20% reduction), cashflow impact and implementation risk
- **04_other_projects_and_workplan.md** — SWOT for Brumby (Option C), Bronco (Option D), and future exploration (Option E). Cross-option comparison matrix A–E. Stage-gated forward workplan (2027–2033, $522.9M total, 3 gates)
- **05_recommendation_gonogo.md** — Do-nothing scenario (Option B), top 3 recommendations with cashflow analysis and implementation risks, proposed timeline
- **06_slide_structure_outline.md** — 14-slide presentation outline (900s total) + 6 appendix slides
- **07_qa_preparation.md** — 64 anticipated board questions, 10 model answers, 8 hostile question prep

### Phase 2 — Assembly (Complete)

All analysis documents were updated with the corrected NPV figures ($997M / $1,277M combined, IRR ~52%). The presentation decks were built:

- **PPTX**: 20 slides built programmatically with python-pptx — native charts (bar, waterfall, timeline)
- **HTML**: 20-slide deck with CSS charts, keyboard navigation, presenter notes panel (press N)
- **Speaker notes**: 900s timed, per-slide notes with key phrases
- **Q&A**: All [DRAFT] placeholders filled, solo presenter framework with topic-to-appendix quick reference

A 55-check verification suite ([`verify_phase2.py`](verify_phase2.py)) was run — all 55 checks pass.

### Phase 3 — Polish & Rehearsal (Ready)

All content is complete. In-person timing rehearsal with the HTML deck timer is the only remaining step.

---

## Excel Workbook Content

### Complex Valuation Model (15 sheets)

The master DCF valuation model covering three mining operations. Originally provided with the Brave Blossom Capital tab as a blank template shell.

| # | Sheet | Purpose | Status |
|---|-------|---------|--------|
| 1 | `Assumptions` | Master inputs (8% discount, FX 0.743, PHCC $240/t, carbon $45–200/t) | ✅ Populated |
| 2 | `Analysis` | DCF summary — Springbok $279M, Brave Blossom $1,320M, Combined $1,670M | ✅ Populated |
| 3 | `List` | Reference list | ✅ Populated |
| 4 | `Summary` | Revenue/cost dashboard | ✅ Populated |
| 5 | `Historical Performance` | Springbok actuals 2011–2025 | ✅ Populated |
| 6 | `Springbok Assumptions` | Scenario toggle (Current vs Proposed operating config) | ✅ Populated |
| 7 | `Springbok` | Full operational + financial model 2027–2053 | ✅ Populated |
| 8 | `SpringbokCapital` | Springbok capital schedule (fully populated) | ✅ Populated |
| 9 | `Brave Blossom ` | Full operational + financial model 2027–2071 | ✅ Populated |
| 10 | `Brave Blossom Capital` | Capital schedule — **was blank, now populated** | ✅ **Populated** |
| 11 | `Historical Performance BrumbyOC` | Brumby open cut actuals 2011–2015 | ✅ Populated |
| 12 | `OC Clean` | Open cut proposal model (template) | ⚠️ Template |
| 13 | `OC Clean Capital` | Open cut capital schedule (template) | ⚠️ Template |
| 14 | `Carbon` | Price forecast scenarios ($0 or $45–200/t) | ✅ Populated |
| 15 | `Decommissioned Mine` | Post-closure CH₄ emissions (NGER Method 3.32) | ✅ Populated |

**Critical correction**: The original workbook computed Brave Blossom NPV at $1,320M with $0 capital (the Capital tab was empty). After populating the 12 capital items with unit costs and year-by-year scheduling from the concept estimate, the corrected base-case NPV is **$997M** (30% contingency, tax shield). The Combined NPV (Springbok + Brave Blossom) corrected from $1,670M to **$1,277M**.

The populated workbook is at [`workbooks/Complex_Valuation_Model_POPULATED.xlsx`](workbooks/Complex_Valuation_Model_POPULATED.xlsx).

### 2024 Planned Closure Cost Workbook (6 sheets)

The $900M ERC (Estimated Rehabilitation Cost) closure model for the Springbok Coal Complex.

| Component | Detail |
|-----------|--------|
| **Total liability** | $900M ($494M direct works + $173M contingency + $233M holding costs over 20 years) |
| **Domains** | 9 rehabilitation domains |
| **Line items** | 284 individual cost items |
| **Unit rates** | 296 rates in a rate library |
| **Audit result** | 0 arithmetic errors. 8 issues/anomalies identified (see [`vault/19`](vault/19_Issues_and_Anomalies.md)) |

The 8 issues found during audit directly informed the 6 SMART closure-reduction targets:

| # | SMART Target | Saving | % of $900M |
|---|-------------|--------|------------|
| 1 | Remove duplicate TSF costing (Domain 2) | $39.5–43.8M | 4.4–4.9% |
| 2 | Reduce contingency 35%→25% | $49.4M | 5.5% |
| 3 | House sale vs demolition (505 houses) | $11.8M | 1.3% |
| 4 | Accelerate progressive rehab ($2.5M→$5M/yr) | $11.2M | 1.2% |
| 5 | Progressive lease relinquishment | ~$50M | 5.6% |
| 6 | Monetise gas drainage post-closure | $14.1M cost→revenue | 1.6%+ |
| | **Total** | **$180.3M** | **20%** |

---

## Key Financial Figures

| Metric | Original (Workbook) | Corrected (This Analysis) |
|--------|-------------------|--------------------------|
| Brave Blossom NPV | $1,320M | **$997M** (base case, 30% contingency + tax shield) |
| Combined NPV | $1,670M | **$1,277M** |
| Brave Blossom IRR | 77.4% | **~52%** |
| Scenario range | — | **$859M–$1,072M** (all positive) |
| Springbok NPV | $279M | $279M (unchanged) |
| Closure liability | $900M | $900M (unchanged) |
| SMART reductions | — | $180.3M (20% of $900M) |
| Closure deferral NPV | ~$495M | $508M (at 7% discount) |
| Capital-to-liability ratio | — | 1:2.31 ($1 defers $2.31 closure) |
| Total value with deferral | — | ~$1,505M (NPV + deferral benefit) |
| Capital (direct) | $389M | $388.9M (unit costs), $518.4M (scheduled) |
| Capital (risked, 30% contingency) | — | $673.9M |
| Coal price realisation | — | $211.20/t (88% of $240 PHCC benchmark) |
| Discount rate | 8% | 8% (unchanged) |
| Emissions above baseline | — | 20.1M tCO₂e (Brave Blossom, 20 years) |
| Decommissioned mine emissions | — | 4.68M tCO₂e (combined, 20 years) |

---

## Deliverables Mapping

How each deliverable from the brief maps to files in this repository:

| Brief Deliverable | Repository File(s) |
|-------------------|-------------------|
| **Section 1**: Brave Blossom valuation & SWOT | [`analysis/01b_financial_model_analysis.md`](analysis/01b_financial_model_analysis.md), [`analysis/02_brave_blossom_swot.md`](analysis/02_brave_blossom_swot.md), [`workbooks/Complex_Valuation_Model_POPULATED.xlsx`](workbooks/Complex_Valuation_Model_POPULATED.xlsx) |
| **Section 2**: Post-mining liability reduction | [`analysis/03_closure_liability_review.md`](analysis/03_closure_liability_review.md), [`vault/17`](vault/17_2024_Planned_Closure_Costs.md), [`vault/18`](vault/18_Audit_Results.md), [`vault/19`](vault/19_Issues_and_Anomalies.md) |
| **Section 3**: Other projects & forward workplan | [`analysis/04_other_projects_and_workplan.md`](analysis/04_other_projects_and_workplan.md) |
| **Section 4**: Go/no-go recommendation & timeline | [`analysis/05_recommendation_gonogo.md`](analysis/05_recommendation_gonogo.md) |
| **Brave Blossom Capital tab populated** | [`workbooks/Complex_Valuation_Model_POPULATED.xlsx`](workbooks/Complex_Valuation_Model_POPULATED.xlsx) |
| **15-min presentation deck** | [`slides/Wallaby_Mining_Board_Presentation.pptx`](slides/Wallaby_Mining_Board_Presentation.pptx), [`slides/deck.html`](slides/deck.html) |
| **Speaker notes** | [`slides/speaker_notes.md`](slides/speaker_notes.md) |
| **Q&A preparation** | [`analysis/07_qa_preparation.md`](analysis/07_qa_preparation.md) |
|| **SWOT (11 areas per option)** | [`analysis/02_brave_blossom_swot.md`](analysis/02_brave_blossom_swot.md) (Brave Blossom), [`analysis/04_other_projects_and_workplan.md`](analysis/04_other_projects_and_workplan.md) (Brumby/Bronco/Option E) |
|| **SMART recommendations** | [`analysis/03_closure_liability_review.md`](analysis/03_closure_liability_review.md) — 6 targets, $180.3M total |
|| **Option A calculation walkthrough** | [`OPTION_A_CALCULATION_WALKTHROUGH.md`](OPTION_A_CALCULATION_WALKTHROUGH.md) — cell-by-cell DCF walkthrough ($997M NPV) |
|| **Option C calculation walkthrough** | [`OPTION_C_CALCULATION_WALKTHROUGH.md`](OPTION_C_CALCULATION_WALKTHROUGH.md) — stage-gate framework + deferral NPV ($508M) |

---

## Status

| Phase | Status | Description |
|-------|--------|-------------|
| **Phase 1 — Analysis** | ✅ Complete | 8 analysis documents (3,994 lines) + populated workbook. 92/92 verification checks pass. |
| **Phase 2 — Assembly** | ✅ Complete | PPTX + HTML decks (20 slides each), speaker notes, Q&A finalised. 55/55 verification checks pass. |
| **Phase 3 — Rehearsal** | Ready | All content complete. In-person timing rehearsal only remaining step. |

### Blocked Items (3)

These items from the brief could not be completed due to missing source information:

1. **Conceptual Mining Study document** — page 8 of the brief is blank/cover only; may contain additional mine plan data not provided
2. **Seam thickness contours** (Figure 6) — only available as a contour map image, not extractable as text/data
3. **Depth of cover values** (Figure 7) — units unclear, not digitisable

These are noted as limitations in the presentation and do not affect the core analysis or recommendations.

---

## Calculation Walkthroughs

The project includes two detailed calculation walkthroughs that trace every financial figure back to specific Excel cells. These are the most technically detailed documents in the repository — designed for readers who want to verify the math themselves by opening the workbook alongside the guide.

### Option A — Brave Blossom Underground Mine

**File**: [`OPTION_A_CALCULATION_WALKTHROUGH.md`](OPTION_A_CALCULATION_WALKTHROUGH.md) (826 lines)

A complete cell-by-cell walkthrough of the Brave Blossom DCF model. It covers:

- **Financial terms explained**: discount rate (8%), NPV, IRR, tax shield, carbon liability — each with plain-English definitions and the workbook cell where it's defined
- **Revenue chain**: production volume → saleable coal → pricing (PHCC benchmark $240/t, 88% realisation) → FX conversion → gross revenue, with exact cell references across the `Brave Blossom` sheet (columns D through BO, years 2027–2052)
- **Cost chain**: operating costs, royalties (tiered), CPI escalation, carbon costs (two scenarios from the `Carbon` sheet)
- **Capital schedule**: 12 items from the `Brave Blossom Capital` sheet — unit costs, fractional unit counts, year-by-year phasing, 30% contingency, tax shield computation
- **Cashflow & NPV**: how net cashflow is built line by line, discounted at 8%, summed to the corrected **$997M base-case NPV** and **~52% IRR**
- **Sensitivity scenarios**: 5 scenarios ($859M–$1,072M), all NPV-positive

**How to use it**: Open `Complex_Valuation_Model_POPULATED.xlsx` in Excel. Each section of the walkthrough references specific sheets and cells (e.g., `Brave Blossom !D7`, `Assumptions!B2`, `Brave Blossom Capital!D57`). Navigate to the referenced cell to see the live formula and verify the calculation.

### Option C — Phased Go: Stage-Gated Brave Blossom

**File**: [`OPTION_C_CALCULATION_WALKTHROUGH.md`](OPTION_C_CALCULATION_WALKTHROUGH.md) (658 lines)

Covers what's **unique to Option C** — the stage-gated decision framework overlaid on the same DCF model as Option A. It does not repeat the shared DCF chain; instead it cross-references the Option A walkthrough. It covers:

- **Stage gates explained**: three gates (Gate 1: end 2028, Gate 2: end 2030, Gate 3: 2032/33) — what each gate reviews and the pass criteria
- **$20M initial commitment**: PFS and exploration drilling, funded upfront before Gate 1
- **Capital phasing**: how the 12 capital items in `Brave Blossom Capital` map to the three stages — which costs are deferred vs committed at each gate
- **Real option value**: the value of optionality — the right to stop at each gate without committing the full $389M–$674M
- **Closure-liability deferral NPV**: a separate $508M benefit (at 7% discount) from deferring the $900M closure liability by developing Brave Blossom instead of closing Springbok in 2032. This is quantified independently of the project NPV.
- **Decision tree**: the three-gate decision tree with branch probabilities and expected values

**How to use it**: Read the Option A walkthrough first (or have it open for reference). Option C's walkthrough assumes familiarity with the DCF chain and focuses on the incremental analysis: gate criteria, capital phasing differences, and the deferral NPV calculation.

---

*Last updated: 2026-08-30*
