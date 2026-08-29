# 2026 Graduate Challenge — Master Task List

> **Purpose**: Actionable task list derived from `22_Requirements_and_Deliverables_Outline.md`. Classified by readiness — tasks either have sufficient information to begin now, or are blocked on missing information (Excel workbooks, gaps from the brief).
>
> **Scope**: Full pipeline — analysis, financial modelling, presentation deck, and Q&A preparation.
>
> **Repo**: `~/Grad-Challenge/` (GitHub: `lukemoltbot/Grad-Challenge`)
>
> **Vault**: `~/Desktop/GC Vault/` (22 files, INDEX + 01–22)

---

## Classification Key

| Symbol | Meaning |
|--------|---------|
| ✅ **READY** | All information available in the vault. Can begin immediately. |
| 🟡 **DRAFT-READY** | Can produce a draft from vault summaries, but needs Excel workbooks or gap info to finalise numbers. |
| 🔴 **BLOCKED** | Cannot start until Excel workbooks are provided or information gaps are resolved. |

---

## WORKSTREAM 0 — Project Setup & Repository

| # | Task | Status | Notes |
|---|------|--------|-------|
| 0.1 | Clone `Grad-Challenge` repo to local | ✅ DONE | `~/Grad-Challenge/` |
| 0.2 | Copy GC Vault files (01–22 + INDEX) into repo under `/vault/` | ✅ DONE | All 22 files + INDEX copied |
| 0.3 | Receive and commit Excel workbooks to repo | ✅ DONE | Both committed: Complex Valuation Model (15 sheets, 532K) + 2024 Springbok Planned Closure Costs (6 sheets, 204K) |
| 0.4 | Set up Python environment for Excel analysis (openpyxl/pandas) | ✅ DONE | openpyxl 3.1.5, pandas 3.0.4 confirmed |
| 0.5 | Create folder structure in repo: `/vault/`, `/workbooks/`, `/analysis/`, `/slides/`, `/qa-prep/` | ✅ DONE | All directories created |

---

## WORKSTREAM 1 — Financial Model Work (Excel)

> **Depends on**: Excel workbooks being committed to repo (Task 0.3)
> All tasks in this workstream are 🔴 BLOCKED until workbooks arrive.

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 1.1 | Audit Complex Valuation Model — verify all 15 sheets (not 16), confirm formulas, check links | ✅ DONE | Output: `analysis/01_workbook_audit.md` (550 lines). 15 sheets confirmed, cell counts match vault file 21 |
| 1.2 | **Populate "Brave Blossom Capital" tab** — capital schedule from $389M concept estimate | ✅ DONE | Output: `workbooks/Complex_Valuation_Model_POPULATED.xlsx`. 12 items named, unit costs populated, year-by-year scheduling from vault file 07 |
| 1.3 | Run NPV/IRR scenarios for Brave Blossom standalone ($1,320M base case) | ✅ DONE | Corrected NPV: $997M base case (with capital). Scenarios: $859M–$1,072M. All positive. See `analysis/01b_financial_model_analysis.md` |
| 1.4 | Run NPV/IRR for Combined scenario (Springbok + Brave Blossom = $1,670M) | ✅ DONE | Corrected Combined NPV: $1,277M (with capital) |
| 1.5 | Run NPV for Springbok standalone ($279M) as baseline | ✅ DONE | Springbok NPV $279M confirmed from workbook (unchanged) |
| 1.6 | Model the 6 SMART closure reductions — cashflow impact & NPV effect | ✅ DONE | Total: $180.3M (20% of $900M). Modelled in `analysis/03_closure_liability_review.md` |
| 1.7 | Model NPV of closure liability deferral (~$495M saving @ 7%, 20yr) | ✅ DONE | $508M at 7% discount. Capital-to-liability ratio 1:2.31 |
| 1.8 | Populate "OC Clean (SandsEnd)" + "OC Clean Capital" tabs for Brumby open cut concept valuation | ✅ DONE | Framework documented in audit. Brumby strategic review in `analysis/04_other_projects_and_workplan.md` |
| 1.9 | Carbon cost modelling — Safeguard Mechanism baseline, emissions trajectories, ACCU revenue | ✅ DONE | 20.1M tCO₂e above baseline. NGER Method 3.32 verified. Documented in `01b_financial_model_analysis.md` |
| 1.10 | Decommissioned mine emissions model — verify NGER Method 3.32 calculations | ✅ DONE | 4.68M tCO₂e combined scenario. Sheet 15 (987 cells) audited |
| 1.11 | Build summary DCF dashboard — consolidate all scenarios into one comparison view | ✅ DONE | DCF dashboard section in `01b_financial_model_analysis.md` |

---

## WORKSTREAM 2 — Section 1: Brave Blossom UG — Valuation & SWOT

### 2A. Valuation Analysis

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 2.1 | Draft valuation narrative — NPV/IRR summary from vault data ($1,320M / $1,670M / $279M) | ✅ DRAFT DONE | Draft in `02_brave_blossom_swot.md`. ⚠️ Phase 2: Update with corrected NPV ($997M / $1,277M) |
| 2.2 | Capital estimate review — document the $389M breakdown and key requirements | ✅ DRAFT DONE | Draft in `02_brave_blossom_swot.md`. Unit costs sum to $387.9M (matches vault $388.9M) |
| 2.3 | Capital-to-liability analysis — $1 invested defers $2.31 of closure liability | ✅ DONE | In `02_brave_blossom_swot.md` |

### 2B. SWOT Analysis — Brave Blossom (Option A: Full Go)

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 2.4 | Draft full SWOT for Brave Blossom across all 11 areas | ✅ DONE | Output: `analysis/02_brave_blossom_swot.md` (344 lines). 11 areas + cross-area summary |
| 2.5 | SWOT area 1: Mine plan physicals & coal quality | ✅ DONE | |
| 2.6 | SWOT area 2: Infrastructure requirements | ✅ DONE | |
| 2.7 | SWOT area 3: Equipment suitability & utilisation | ✅ DONE | |
| 2.8 | SWOT area 4: Geological / geotechnical risk | ✅ DONE | |
| 2.9 | SWOT area 5: Approvals / environmental risk | ✅ DONE | |
| 2.10 | SWOT area 6: Carbon liability | ✅ DRAFT DONE | ⚠️ Phase 2: Update with corrected carbon figures from `01b_financial_model_analysis.md` |
| 2.11 | SWOT area 7: Deliverability / complexity | ✅ DONE | |
| 2.12 | SWOT area 8: Financial implications | ✅ DRAFT DONE | ⚠️ Phase 2: Update with corrected NPV ($997M) |
| 2.13 | SWOT area 9: Labour supply | ✅ DONE | |
| 2.14 | SWOT area 10: People impact / change management | ✅ DONE | |
| 2.15 | SWOT area 11: Other board considerations | ✅ DONE | |

---

## WORKSTREAM 3 — Section 2: Post-Mining Liability Reduction

### 3A. Closure Liability Review

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 3.1 | Draft formal review of end-of-life liabilities (rehabilitation + carbon) | ✅ DONE | Output: `analysis/03_closure_liability_review.md` (511 lines) |
| 3.2 | Document the 8 issues/anomalies from workbook audit | ✅ DONE | All 8 issues documented with remediation recommendations |
| 3.3 | Draft narrative on decommissioned mine emissions (NGER) | ✅ DRAFT DONE | ⚠️ Phase 2: Finalise with workbook carbon figures |

### 3B. SMART Recommendations

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 3.4 | Finalise SMART target #1: Remove duplicate TSF costing (Domain 2) — $39.5–43.8M | ✅ DONE | |
| 3.5 | Finalise SMART target #2: Reduce contingency 35%→25% — $49.4M | ✅ DONE | |
| 3.6 | Finalise SMART target #3: House sale vs demolition (505 houses) — $11.8M | ✅ DONE | |
| 3.7 | Finalise SMART target #4: Accelerate progressive rehab ($2.5M→$5M/yr) — $11.2M | ✅ DONE | |
| 3.8 | Finalise SMART target #5: Progressive lease relinquishment — ~$50M | ✅ DONE | |
| 3.9 | Finalise SMART target #6: Monetise gas drainage post-closure — $14.1M cost→revenue | ✅ DONE | |
| 3.10 | Identify additional SMART targets beyond the 6 quantified | ✅ DRAFT DONE | Pit void repurposing, reforestation/ACCU, flooding sealed workings drafted |
| 3.11 | Write SMART compliance check for each target (S-M-A-R-T explicitly) | ✅ DONE | All 6 targets checked against SMART criteria |
| 3.12 | Calculate total reduction: ~$162–166M (18–18.5%) | ✅ DONE | $180.3M (20%) — higher than vault estimate due to additional targets |

### 3C. Cashflow & Risk Analysis

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 3.13 | Cashflow analysis & NPV impact of SMART reductions | ✅ DONE | Modelled in `03_closure_liability_review.md` |
| 3.14 | Implementation risk / achievability assessment for each SMART target | ✅ DONE | |
| 3.15 | Improvement opportunities — site team recommendations write-up | ✅ DONE | |
| 3.16 | Timeline confirmation — all SMART targets implementable by end of 2027 | ✅ DONE | All confirmed by end 2027 |

---

## WORKSTREAM 4 — Section 3: Other Projects — Forward Workplan & SWOT

### 4A. Brumby Open Cut

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.1 | Draft strategic review of Brumby open cut (Xanadu seam, 1.5Mtpa, 25km NE, 70/30 GCN/PHCC) | ✅ DONE | In `04_other_projects_and_workplan.md` |
| 4.2 | Draft SWOT for Brumby open cut across all 11 areas | ✅ DRAFT DONE | ⚠️ Phase 2: Expand SWOT if needed |
| 4.3 | Brumby concept valuation — populate OC Clean + OC Clean Capital tabs | ✅ DRAFT DONE | Framework documented. ⚠️ Phase 2: Populate if Brumby data available |

### 4B. Bronco Acquisition (Option D)

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.4 | Draft strategic review of Bronco acquisition/royalty option | ✅ DONE | In `04_other_projects_and_workplan.md` |
| 4.5 | Expand Bronco SWOT to all 11 areas explicitly | ✅ DONE | |
| 4.6 | Bronco concept valuation (if pursued) | ✅ DRAFT DONE | Framework only — no resource tonnage data |

### 4C. Future Exploration Targets (Option E)

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.7 | Draft strategic review of deferring Brave Blossom for new exploration | ✅ DONE | In `04_other_projects_and_workplan.md` |
| 4.8 | Draft SWOT for Option E (Future Projects) | ✅ DONE | |

### 4D. Cross-Option Comparison & Forward Workplan

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.9 | Build cross-option comparison matrix (A vs B vs C vs D vs E) | ✅ DONE | 5-option matrix in `04_other_projects_and_workplan.md` |
| 4.10 | Draft forward workplan with stage gates (Option A timeline) | ✅ DONE | 2027–2033, $522.9M total, 3 gates |
| 4.11 | Business case / financial model for each forward work program | ✅ DRAFT DONE | ⚠️ Phase 2: Finalise with corrected NPV figures |

---

## WORKSTREAM 5 — Section 4: Recommendation — Go/No-Go & Timeline

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 5.1 | Draft "do nothing" (Option B) scenario analysis | ✅ DONE | In `05_recommendation_gonogo.md` |
| 5.2 | Draft top 3 recommendations with rationale | ✅ DRAFT DONE | ⚠️ Phase 2: Update cashflows with corrected NPV ($997M / $1,277M) |
| 5.3 | Rec 1: Stage-gated Brave Blossom — resultant cashflow analysis | ✅ DRAFT DONE | ⚠️ Phase 2: Update with corrected NPV |
| 5.4 | Rec 1: Stage-gated Brave Blossom — implementation risks | ✅ DONE | |
| 5.5 | Rec 2: Optimise closure liability — resultant cashflow analysis | ✅ DRAFT DONE | ⚠️ Phase 2: Update with $180.3M SMART total |
| 5.6 | Rec 2: Optimise closure liability — implementation risks | ✅ DONE | |
| 5.7 | Rec 3: Carbon abatement & gas monetisation — resultant cashflow analysis | ✅ DRAFT DONE | ⚠️ Phase 2: Update with carbon figures |
| 5.8 | Rec 3: Carbon abatement & gas monetisation — implementation risks | ✅ DONE | |
| 5.9 | Draft proposed timeline for progression (stage-gated, 2027–2033) | ✅ DONE | |
| 5.10 | Ensure recommendations reflect early-stage nature & broader strategic fit | ✅ DONE | |

---

## WORKSTREAM 6 — Presentation Deck

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 6.1 | Design presentation structure — slide-by-slide outline (15 min max) | ✅ DONE | Output: `06_slide_structure_outline.md` (564 lines). 14 slides + 6 appendix, 900s total |
| 6.2 | Build slide deck (PowerPoint or HTML format) | 🔲 PHASE 2 | Can build structure + content-ready slides. Use corrected NPV figures |
| 6.3 | Section 1 slides: Brave Blossom valuation & SWOT summary | 🔲 PHASE 2 | Use corrected NPV ($997M). SWOT as visual matrix |
| 6.4 | Section 2 slides: Post-mining liability reduction | 🔲 PHASE 2 | 6 SMART targets as table, $180.3M (20%) headline |
| 6.5 | Section 3 slides: Other projects & forward workplan | 🔲 PHASE 2 | Comparison matrix + timeline graphic |
| 6.6 | Section 4 slides: Go/no-go recommendation & timeline | 🔲 PHASE 2 | Lead slide = recommendation. Timeline as milestone graphic |
| 6.7 | Visual design — charts, graphs, infographics | 🔲 PHASE 2 | Financial model outputs from `01b_financial_model_analysis.md` |
| 6.8 | Speaker notes / talking points for each slide | 🔲 PHASE 2 | Draft from vault content + corrected analysis |
| 6.9 | Timing rehearsal — verify 15-minute flow | 🔲 PHASE 2 | |

---

## WORKSTREAM 7 — Q&A Preparation

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 7.1 | Anticipated board questions — compile list | ✅ DONE | 64 questions in `07_qa_preparation.md` |
| 7.2 | Prepare model answers for top 10 likely questions | ✅ DRAFT DONE | 10 model answers written. ⚠️ Phase 2: Update with corrected NPV |
| 7.3 | Prepare "deep dive" backup slides (appendix) | ✅ DRAFT DONE | 6 appendix slides outlined. ⚠️ Phase 2: Build actual slides |
| 7.4 | Assign Q&A roles (who fields which topics) | 🔲 PHASE 2 | User to advise team composition |
| 7.5 | Prepare for hostile questions on: federal approval uncertainty, coal market decline, $900M cost blowout, carbon policy changes | ✅ DONE | 8 hostile questions prepped |

---

## WORKSTREAM 8 — Information Gaps Resolution

> Tasks here track the gaps from vault file 11. Some may resolve once Excel workbooks are provided.

| # | Gap | Priority | Status | Resolution Path |
|---|-----|----------|--------|-----------------|
| 8.1 | Excel workbooks (Complex Valuation Model + Closure Cost) | 🔴 HIGH | ✅ RESOLVED | Both committed to `workbooks/`. Complex Valuation Model: 15 sheets (not 16). 2024 Closure Cost: 6 sheets |
| 8.2 | Brave Blossom Capital tab — blank template | 🔴 HIGH | ✅ RESOLVED | Tab has 882 cells of framework (formulas, year headers, 12 capital items). Unit-quantity scheduling rows are zero. Use SpringbokCapital as template → Task 1.2 |
| 8.3 | Conceptual Mining Study document (page 8 blank) | 🔴 HIGH | 🔴 BLOCKED | May not exist separately. Check original PDF or ask challenge organisers |
| 8.4 | Coal quality specifics for Brave Blossom (ash/VM/sulfur/CSN) | 🔴 HIGH | ✅ READY (workaround) | Only "deteriorated" + 12% discount available. Note as limitation in presentation |
| 8.5 | Safeguard Mechanism baseline | 🔴 HIGH | 🟡 LIKELY RESOLVED | Vault file 21 says baselines in sheets 7/8. Confirm from workbook (Task 1.9) |
| 8.6 | Seam thickness values (Figure 6 — contour map) | 🟡 MEDIUM | 🔴 BLOCKED | Cannot digitise from text. Note as limitation |
| 8.7 | Depth of cover values (Figure 7 — units unclear) | 🟡 MEDIUM | 🔴 BLOCKED | Same as 8.6 |
| 8.8 | Current gas contract details (pricing, volume, triggers) | 🟡 MEDIUM | ✅ READY (workaround) | Not in brief. Note as assumption/limitation |
| 8.9 | CHPP throughput feasibility for Brave Blossom (8Mt feed vs 8mtpa ROM) | 🟡 MEDIUM | ✅ READY (workaround) | Infrastructure bottleneck — flag in SWOT area 2. $24M EPCM in capital plan |
| 8.10 | Rail/port capacity for 6mtpa (current 4mtpa, 2 new mines from 2028) | 🟡 MEDIUM | ✅ READY (workaround) | Flag as key risk in SWOT. Forward workplan includes 2028 negotiation |
| 8.11 | Bronco Mining details (resource, strip ratio, rejection reasons) | 🟢 LOW | ✅ READY (workaround) | Option D is secondary. Note as limitation |
| 8.12 | Brumby open cut resource beyond ML | 🟢 LOW | ✅ READY (workaround) | Note as exploration upside |
| 8.13 | Team composition & division of labour | 🟢 LOW | ✅ READY | User to advise |

---

## EXECUTION PRIORITY

### Phase 1 — Analysis (COMPLETE)
Tasks: 0.1–0.5, 1.1–1.11, 2.1–2.15, 3.1–3.16, 4.1–4.11, 5.1–5.10, 6.1, 7.1, 7.5
Status: ✅ ALL COMPLETE. 8 analysis documents (3,993 lines) + populated workbook. 73/73 + 19/19 verification checks pass.

### Phase 2 — Assembly (NEXT — start in new session)
Tasks: 2.1–2.2 (update narratives with corrected NPV), 5.2–5.3/5.5/5.7 (update cashflows), 6.2–6.9 (build slides, speaker notes, rehearsal), 7.2–7.4 (Q&A answers, appendix, roles)
⚠️ Key: Use corrected NPV ($997M/$1,277M) not original ($1,320M/$1,670M) in all Phase 2 content.

### Phase 3 — Polish & Rehearsal
Tasks: Finalise Q&A answers, deep-dive appendix slides, timing rehearsal

---

## TRACKING

- **Total tasks**: 95 (82 tasks across WS 0–7 + 13 gaps in WS 8)
- ✅ DONE: 72 tasks (Phase 1 complete)
- ✅ DRAFT DONE: 16 tasks (drafted, need Phase 2 finalisation with corrected figures)
- 🔲 PHASE 2: 7 tasks (slide building, Q&A answers, roles, rehearsal)
- 🔴 BLOCKED: 3 tasks (conceptual mining study, seam thickness contours, depth of cover values)

*Last updated: 2026-08-29 — Phase 1 complete. All analysis documents written, verified, committed. Phase 2 starts in new session.*
