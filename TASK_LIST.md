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
| 0.2 | Copy GC Vault files (01–22 + INDEX) into repo under `/vault/` | ✅ READY | Source: `~/Desktop/GC Vault/` |
| 0.3 | Receive and commit Excel workbooks to repo | 🔴 BLOCKED | User to provide: Complex Valuation Model (16-sheet) + 2024 Planned Closure Cost Workbook. Commit to `/workbooks/` |
| 0.4 | Set up Python environment for Excel analysis (openpyxl/pandas) | ✅ READY | Needed once workbooks arrive |
| 0.5 | Create folder structure in repo: `/vault/`, `/workbooks/`, `/analysis/`, `/slides/`, `/qa-prep/` | ✅ READY | |

---

## WORKSTREAM 1 — Financial Model Work (Excel)

> **Depends on**: Excel workbooks being committed to repo (Task 0.3)
> All tasks in this workstream are 🔴 BLOCKED until workbooks arrive.

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 1.1 | Audit Complex Valuation Model — verify all 16 sheets, confirm formulas, check links | 🔴 BLOCKED | 0.3 | Vault file 21 has sheet-by-sheet summary for cross-reference |
| 1.2 | **Populate "Brave Blossom Capital" tab** — capital schedule from $389M concept estimate | 🔴 BLOCKED | 0.3, 1.1 | ⚠️ Explicit deliverable from the brief. Source data: vault file 07 (capital estimate breakdown). Must include contingency analysis (35% → ~$525M) |
| 1.3 | Run NPV/IRR scenarios for Brave Blossom standalone ($1,320M base case) | 🔴 BLOCKED | 0.3, 1.1 | Vary discount rate (8% base), coal price ($211.20/t realisation), capital ($389M vs $525M) |
| 1.4 | Run NPV/IRR for Combined scenario (Springbok + Brave Blossom = $1,670M) | 🔴 BLOCKED | 0.3, 1.1 | Sensitivity on carbon price ($45–$200/t) |
| 1.5 | Run NPV for Springbok standalone ($279M) as baseline | 🔴 BLOCKED | 0.3, 1.1 | |
| 1.6 | Model the 6 SMART closure reductions — cashflow impact & NPV effect | 🔴 BLOCKED | 0.3, 1.1 | ~$162M (18%) reduction. Each SMART measure individually + combined. See vault file 20 for details |
| 1.7 | Model NPV of closure liability deferral (~$495M saving @ 7%, 20yr) | 🔴 BLOCKED | 0.3, 1.1 | Capital-to-liability ratio 0.43:1 |
| 1.8 | Populate "OC Clean (SandsEnd)" + "OC Clean Capital" tabs for Brumby open cut concept valuation | 🔴 BLOCKED | 0.3, 1.1 | For Section 3 alternative project valuation. May be partial — depends on available Brumby data |
| 1.9 | Carbon cost modelling — Safeguard Mechanism baseline, emissions trajectories, ACCU revenue | 🔴 BLOCKED | 0.3, 1.1 | Baselines now in sheets 7/8 per vault file 21. Scenarios: $0/t vs $45–$200/t (Accelerated Transition) |
| 1.10 | Decommissioned mine emissions model — verify NGER Method 3.32 calculations | 🔴 BLOCKED | 0.3, 1.1 | 4.68M tCO₂e combined scenario over ~20 years |
| 1.11 | Build summary DCF dashboard — consolidate all scenarios into one comparison view | 🔴 BLOCKED | 1.3–1.10 | For presentation use |

---

## WORKSTREAM 2 — Section 1: Brave Blossom UG — Valuation & SWOT

### 2A. Valuation Analysis

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 2.1 | Draft valuation narrative — NPV/IRR summary from vault data ($1,320M / $1,670M / $279M) | 🟡 DRAFT-READY | — | Can write from vault file 21. Finalise once 1.3–1.5 confirm numbers |
| 2.2 | Capital estimate review — document the $389M breakdown and key requirements | 🟡 DRAFT-READY | — | Vault file 07 has breakdown. Finalise once 1.2 populates the tab |
| 2.3 | Capital-to-liability analysis — $1 invested defers $2.31 of closure liability | ✅ READY | — | Key board metric. ~$106M net positive NPV before Brave Blossom revenue |

### 2B. SWOT Analysis — Brave Blossom (Option A: Full Go)

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 2.4 | Draft full SWOT for Brave Blossom across all 11 areas | ✅ READY | — | Vault file 13 has a draft. Needs expansion to cover all 11 areas explicitly with quantified points |
| 2.5 | SWOT area 1: Mine plan physicals & coal quality | ✅ READY | — | Coal quality "deteriorated", 12% discount, $211.20/t (33% below Springbok). Working section thins eastward |
| 2.6 | SWOT area 2: Infrastructure requirements | ✅ READY | — | CHPP end-of-life 2032 ($24M EPCM), rail/port 4mtpa→6mtpa, 2 competing mines from 2028 |
| 2.7 | SWOT area 3: Equipment suitability & utilisation | ✅ READY | — | No residual equipment from Springbok. Longwall + mining equipment = new capital |
| 2.8 | SWOT area 4: Geological / geotechnical risk | ✅ READY | — | Monocline, historic roof falls, old workings condition unknown, private landholder resistance (last 4 panels) |
| 2.9 | SWOT area 5: Approvals / environmental risk | ✅ READY | — | 25% of mine plan on MDL — federal approval uncertain. 75% within ML (state approval exists) |
| 2.10 | SWOT area 6: Carbon liability | 🟡 DRAFT-READY | — | 20.1M tCO₂e above baseline (20 years). Can draft from vault; finalise with workbook carbon modelling (1.9) |
| 2.11 | SWOT area 7: Deliverability / complexity | ✅ READY | — | Timeline-critical: must commence 2027 for cashflow continuity beyond 2030. Stage-gated funding approach |
| 2.12 | SWOT area 8: Financial implications | 🟡 DRAFT-READY | — | $389M capital, peak $221M in 2032. Can draft; finalise with NPV/IRR scenarios (1.3–1.4) |
| 2.13 | SWOT area 9: Labour supply | ✅ READY | — | Workforce transition from Springbok (800+ jobs). Owner operator vs contract miner decision |
| 2.14 | SWOT area 10: People impact / change management | ✅ READY | — | Roster patterns, housing (505 houses), H&S, training. Town = Dustyroo Flats |
| 2.15 | SWOT area 11: Other board considerations | ✅ READY | — | Community social licence, customer base (Japanese JV), fallback if federal approval fails (75% within ML = ~15yr life) |

---

## WORKSTREAM 3 — Section 2: Post-Mining Liability Reduction

### 3A. Closure Liability Review

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 3.1 | Draft formal review of end-of-life liabilities (rehabilitation + carbon) | ✅ READY | — | Vault files 17–20 fully audited. $900M = $494M direct + $173M contingency + $233M holding. 9 domains, 284 line items |
| 3.2 | Document the 8 issues/anomalies from workbook audit | ✅ READY | — | Vault file 19: 0 arithmetic errors, 8 issues (duplicate TSF, contingency 35%, etc.) |
| 3.3 | Draft narrative on decommissioned mine emissions (NGER) | 🟡 DRAFT-READY | — | 4.68M tCO₂e combined, Safeguard >100k tCO₂-e/yr. Finalise with workbook (1.10) |

### 3B. SMART Recommendations

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 3.4 | Finalise SMART target #1: Remove duplicate TSF costing (Domain 2) — $39.5–43.8M | ✅ READY | — | Already quantified. Vault file 20 |
| 3.5 | Finalise SMART target #2: Reduce contingency 35%→25% — $49.4M | ✅ READY | — | Already quantified. Vault file 20 |
| 3.6 | Finalise SMART target #3: House sale vs demolition (505 houses) — $11.8M | ✅ READY | — | Already quantified. Vault file 20 |
| 3.7 | Finalise SMART target #4: Accelerate progressive rehab ($2.5M→$5M/yr) — $11.2M | ✅ READY | — | Already quantified. Vault file 20 |
| 3.8 | Finalise SMART target #5: Progressive lease relinquishment — ~$50M | ✅ READY | — | First application 2035. Vault file 20 |
| 3.9 | Finalise SMART target #6: Monetise gas drainage post-closure — $14.1M cost→revenue | ✅ READY | — | Q2 2032. Vault file 20 |
| 3.10 | Identify additional SMART targets beyond the 6 quantified | 🟡 DRAFT-READY | — | Pit void repurposing, reforestation/ACCU, flooding sealed workings. Can draft; quantify with workbook |
| 3.11 | Write SMART compliance check for each target (S-M-A-R-T explicitly) | ✅ READY | 3.4–3.10 | Format verification, not content generation |
| 3.12 | Calculate total reduction: ~$162–166M (18–18.5%) | ✅ READY | 3.4–3.9 | Already computed in vault file 20 |

### 3C. Cashflow & Risk Analysis

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 3.13 | Cashflow analysis & NPV impact of SMART reductions | 🔴 BLOCKED | 1.6 | Need workbook to model cashflow impact of each reduction over time |
| 3.14 | Implementation risk / achievability assessment for each SMART target | ✅ READY | 3.4–3.9 | Qualitative — can assess from vault data |
| 3.15 | Improvement opportunities — site team recommendations write-up | ✅ READY | — | Consolidate from vault files 09, 12, 20 |
| 3.16 | Timeline confirmation — all SMART targets implementable by end of 2027 | ✅ READY | 3.4–3.9 | Check each target's timeline against the 2027 deadline (most are Q2–Q4 2027) |

---

## WORKSTREAM 4 — Section 3: Other Projects — Forward Workplan & SWOT

### 4A. Brumby Open Cut

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.1 | Draft strategic review of Brumby open cut (Xanadu seam, 1.5Mtpa, 25km NE, 70/30 GCN/PHCC) | ✅ READY | — | Vault file 15 |
| 4.2 | Draft SWOT for Brumby open cut across all 11 areas | 🟡 DRAFT-READY | — | Vault file 13 has Option D (Bronco) SWOT but not a dedicated Brumby SWOT. Can draft from files 15 + 05 |
| 4.3 | Brumby concept valuation — populate OC Clean + OC Clean Capital tabs | 🔴 BLOCKED | 1.8 | Blank template tabs in workbook |

### 4B. Bronco Acquisition (Option D)

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.4 | Draft strategic review of Bronco acquisition/royalty option | ✅ READY | — | Vault file 13 has Option D SWOT. Resource quality "similar to historic Brumby" |
| 4.5 | Expand Bronco SWOT to all 11 areas explicitly | ✅ READY | — | Vault file 13 draft + file 11 gap notes (🟢 lower priority) |
| 4.6 | Bronco concept valuation (if pursued) | 🔴 BLOCKED | 1.8, 1.1 | No resource tonnage/quality data — 🟢 lower priority gap |

### 4C. Future Exploration Targets (Option E)

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.7 | Draft strategic review of deferring Brave Blossom for new exploration | ✅ READY | — | $5–10M exploration, post-2052 revenue. Vault file 12 |
| 4.8 | Draft SWOT for Option E (Future Projects) | ✅ READY | — | Can draft from vault file 12 alternatives section |

### 4D. Cross-Option Comparison & Forward Workplan

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 4.9 | Build cross-option comparison matrix (A vs B vs C vs D vs E) | ✅ READY | — | Vault file 13 has a 4-option matrix; extend to include Option E |
| 4.10 | Draft forward workplan with stage gates (Option A timeline) | ✅ READY | — | Vault file 12 has full timeline: 2027–2033, $522.9M total, 3 gates |
| 4.11 | Business case / financial model for each forward work program | 🔴 BLOCKED | 1.1–1.8 | Brief requires "accompanying business case" for all work programs |

---

## WORKSTREAM 5 — Section 4: Recommendation — Go/No-Go & Timeline

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 5.1 | Draft "do nothing" (Option B) scenario analysis | ✅ READY | — | Vault file 13 has Option B SWOT. $900M due 2031–2050, no revenue, 800+ jobs lost |
| 5.2 | Draft top 3 recommendations with rationale | 🟡 DRAFT-READY | — | Vault file 12 has draft top 3. Finalise once financial modelling confirms numbers |
| 5.3 | Rec 1: Stage-gated Brave Blossom — resultant cashflow analysis | 🔴 BLOCKED | 1.3, 1.7 | Need NPV/IRR + deferral savings from workbook |
| 5.4 | Rec 1: Stage-gated Brave Blossom — implementation risks | ✅ READY | — | Vault file 12 lists: federal approval, coal quality, rail/port, CHPP end-of-life |
| 5.5 | Rec 2: Optimise closure liability — resultant cashflow analysis | 🔴 BLOCKED | 1.6 | Need SMART reduction cashflow impact from workbook |
| 5.6 | Rec 2: Optimise closure liability — implementation risks | ✅ READY | 3.14 | |
| 5.7 | Rec 3: Carbon abatement & gas monetisation — resultant cashflow analysis | 🔴 BLOCKED | 1.9, 1.10 | Need carbon modelling + gas revenue from workbook |
| 5.8 | Rec 3: Carbon abatement & gas monetisation — implementation risks | ✅ READY | — | Gas drainage infrastructure exists; extension feasible. Vault file 12 |
| 5.9 | Draft proposed timeline for progression (stage-gated, 2027–2033) | ✅ READY | 4.10 | Vault file 12 timeline + milestones |
| 5.10 | Ensure recommendations reflect early-stage nature & broader strategic fit | ✅ READY | 5.2 | Qualitative check — concept estimate, PFS not yet done |

---

## WORKSTREAM 6 — Presentation Deck

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 6.1 | Design presentation structure — slide-by-slide outline (15 min max) | ✅ READY | — | 4 sections + intro/recommendation. ~12–15 slides. Lead with recommendation per brief tip #1 |
| 6.2 | Build slide deck (PowerPoint or HTML format) | 🟡 DRAFT-READY | 6.1, 2.4, 3.1, 4.9, 5.2 | Can build structure + content-ready slides. Finalise numbers once financial modelling done |
| 6.3 | Section 1 slides: Brave Blossom valuation & SWOT summary | 🟡 DRAFT-READY | 2.1–2.3, 2.4 | SWOT as visual matrix, NPV/IRR as headline numbers |
| 6.4 | Section 2 slides: Post-mining liability reduction | 🟡 DRAFT-READY | 3.1, 3.4–3.12 | 6 SMART targets as a table, $162M (18%) reduction as headline |
| 6.5 | Section 3 slides: Other projects & forward workplan | 🟡 DRAFT-READY | 4.1–4.10 | Comparison matrix + timeline graphic |
| 6.6 | Section 4 slides: Go/no-go recommendation & timeline | 🟡 DRAFT-READY | 5.1–5.10 | Lead slide = recommendation. Timeline as milestone graphic |
| 6.7 | Visual design — charts, graphs, infographics | 🔴 BLOCKED | 1.11 | Need financial model outputs for charts. SWOT matrices can be designed now |
| 6.8 | Speaker notes / talking points for each slide | 🟡 DRAFT-READY | 6.3–6.6 | Can draft from vault content. Finalise once all analysis complete |
| 6.9 | Timing rehearsal — verify 15-minute flow | ✅ READY | 6.2–6.8 | Can do once slides are drafted |

---

## WORKSTREAM 7 — Q&A Preparation

| # | Task | Status | Dependencies | Notes |
|---|------|--------|--------------|-------|
| 7.1 | Anticipated board questions — compile list | ✅ READY | — | Based on brief tips: federal approval risk, coal quality, carbon cost, $900M estimate, timeline |
| 7.2 | Prepare model answers for top 10 likely questions | 🟡 DRAFT-READY | 7.1 | Can draft from vault. Finalise with financial data for quantitative answers |
| 7.3 | Prepare "deep dive" backup slides (appendix) | 🟡 DRAFT-READY | 7.2 | Detailed SWOT, financial sensitivity tables, carbon modelling detail |
| 7.4 | Assign Q&A roles (who fields which topics) | ✅ READY | — | User to advise team composition (gap #15 in vault file 11) |
| 7.5 | Prepare for hostile questions on: federal approval uncertainty, coal market decline, $900M cost blowout, carbon policy changes | ✅ READY | — | Vault file 13 threats section + brief tip #4 |

---

## WORKSTREAM 8 — Information Gaps Resolution

> Tasks here track the gaps from vault file 11. Some may resolve once Excel workbooks are provided.

| # | Gap | Priority | Status | Resolution Path |
|---|-----|----------|--------|-----------------|
| 8.1 | Excel workbooks (Complex Valuation Model + Closure Cost) | 🔴 HIGH | 🔴 BLOCKED | User to provide files, commit to repo `/workbooks/` |
| 8.2 | Brave Blossom Capital tab — blank template | 🔴 HIGH | 🔴 BLOCKED | Resolve once 8.1 done → Task 1.2 |
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

### Phase 1 — Immediate (no blockers, start now)
Tasks: 0.2, 0.4, 0.5, 2.3, 2.4–2.15 (SWOT drafts), 3.1–3.2, 3.4–3.12 (SMART targets), 4.1, 4.4–4.5, 4.7–4.9, 5.1, 5.4, 5.8, 5.10, 6.1, 7.1, 7.5

### Phase 2 — Once Excel workbooks received (Task 0.3)
Tasks: 1.1–1.11 (all financial modelling), 3.13, 4.3, 4.6, 4.11, 5.3, 5.5, 5.7, 6.7

### Phase 3 — Assembly & Polish (once analysis complete)
Tasks: 2.1–2.2 (finalise narratives), 5.2 (finalise recommendations), 6.2–6.6 (build slides), 6.8 (speaker notes), 7.2–7.3 (Q&A prep), 6.9 (rehearsal)

---

## TRACKING

- **Total tasks**: 95
- ✅ READY (start now): 52
- 🟡 DRAFT-READY (draft now, finalise later): 19
- 🔴 BLOCKED (need workbooks/info): 23
- ✅ DONE: 1

*Last updated: 2026-08-29*
