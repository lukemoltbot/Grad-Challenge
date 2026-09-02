# Brave Blossom Capital Sheet — Review & Assessment

**Source:** `Complex_Valuation_Model_-_Team_Green_1.xlsx` (Team Green submission)
**Sheet:** `Brave Blossom Capital` (66 rows × 30 cols)
**Comparison baseline:** Previous populated workbook (`Complex_Valuation_Model_POPULATED.xlsx`)
**Date:** 2026-09-02
**Status:** Ready for incorporation into revised analysis

---

## 1. Executive Summary

The Team Green workbook **completely replaces** the Brave Blossom capital schedule with a granular, equipment-level breakdown. The old schedule had 12 project-level line items; the new schedule has **24 equipment-level items** with fundamentally different capital philosophy and timing.

**Headline numbers:**

| Metric | New (Team Green) | Old (Populated) | Delta |
|--------|-----------------:|----------------:|------:|
| Line items | 24 | 12 | +12 |
| Nominal subtotal | $521.2M | $518.4M | +$2.8M |
| Nominal w/ 30% contingency | $677.6M | $673.9M | +$3.6M |
| PV of capex (base, 8% DF) | $398.6M | $341.4M | +$57.2M |
| PV of capex (w/ 30% cont, 8% DF) | $518.2M | $443.8M | +$74.4M |
| Capital spread (active years) | Y1–Y8 (2027–2034) | Y1–Y7 (2027–2033) | +1 yr |
| Peak year | Y1: $132.1M | Y6: $221.5M | Shifted forward |

**Critical finding:** The new schedule is **nearly identical nominally** (+$2.8M, +0.5%) but **$74.4M more expensive in present-value terms** (+16.8%). This divergence is driven entirely by **front-loading**: the new schedule deploys $132.1M in Year 1 versus $10M in the old schedule. The old schedule peaked in Year 6 ($221.5M for the longwall package), which discounts heavily over 6.5 years; the new schedule front-loads equipment acquisition into Years 1–2.

**Estimated NPV impact:** The $74.4M PV increase in capital would reduce the corrected Brave Blossom project NPV from ~$997M to ~**$923M** (a ~7.5% reduction), assuming revenue and operating costs are unchanged.

---

## 2. Structural Overview

### Key structural changes:

- **Old approach:** Project-level / phase-based (e.g., "Longwall" = $190M single line, "Mining Equipment" = $69.9M single line, "Infrastructure + CHPP EPCM" = $48M)
- **New approach:** Equipment-level / unit-based (e.g., Shearer $11M, x300 Roof Supports $300M, x3 Continuous Miners $27M, CHPP components broken out individually)
- **Year columns:** Both workbooks use 27 year columns (D:AD), starting calendar year 2027 (referencing `Assumptions!J1` → `I1+1` where `I1=2026`). Old schedule uses 7 active years (2027–2033); new uses 8 (2027–2034).
- **Contingency:** Both use 30% (cell `C56=0.3`)
- **Formula structure:** Identical computation framework — rows 4–27 are inputs (unit cost in column C, unit counts in year columns D+), rows 31–54 compute costs (`=$C{row}*{col}{row}`), row 55 sums, row 56 applies contingency (`=$C$56*D55`), row 57 totals (`=D56+D55`)
- **Cached values:** The new workbook **has cached formula values** (opened/saved in Excel). All computed figures below are from the workbook's own cached results, cross-verified by independent recomputation.

---

## 3. Complete Line Item Inventory (New Workbook)

| # | Item | Unit Cost (AUD$k) | Total Units | Total Cost (AUDM) | % of Base | Year Spread |
|---|------|------------------:|------------:|-------------------:|----------:|-------------|
| 1 | Shearer | 11,000 | 1.00 | 11.00 | 2.1% | Y1–Y4 (0.25/yr) |
| 2 | CHPP Structural Repairs | 5,000 | 1.00 | 5.00 | 1.0% | Y2–Y3 (0.5/yr) |
| 3 | CHPP Conveyor | 5,000 | 1.00 | 5.00 | 1.0% | Y4–Y5 (0.5/yr) |
| 4 | CHPP Raw Coal Bin | 9,000 | 1.00 | 9.00 | 1.7% | Y3–Y6 (0.25/yr) |
| 5 | Ventilation Network | 20,000 | 1.00 | 20.00 | 3.8% | Y1 (1.0) |
| 6 | Gas Drainage | 12,000 | 1.00 | 12.00 | 2.3% | Y1–Y2 (0.5/yr) |
| 7 | x3 Continuous Miners | 9,000 | 3.00 | 27.00 | 5.2% | Y1–Y6 (0.5/yr) |
| 8 | **x300 Roof Supports** | **1,000** | **300.00** | **300.00** | **57.6%** | Y1–Y6 (50/yr) |
| 9 | x10 DCBs (AFC, BSL, Dewater, CMEs) | 1,000 | 10.00 | 10.00 | 1.9% | Y1–Y5 (2/yr) |
| 10 | x4 U/G Substations | 3,000 | 4.00 | 12.00 | 2.3% | Y1–Y4 (1/yr) |
| 11 | x10 HT Cables | 1,000 | 10.00 | 10.00 | 1.9% | Y1:5, Y2:2, Y3:2, Y4:1 |
| 12 | AFC (Panline & Chain) | 6,000 | 1.00 | 6.00 | 1.2% | Y1 (1.0) |
| 13 | BSL | 6,000 | 1.00 | 6.00 | 1.2% | Y1 (1.0) |
| 14 | Drift | 12,000 | 1.00 | 12.00 | 2.3% | Y1–Y4 (0.25/yr) |
| 15 | x4 Shuttle Cars | 1,000 | 4.00 | 4.00 | 0.8% | Y1–Y4 (1/yr) |
| 16 | x2 Breaker Feeder | 2,000 | 2.00 | 4.00 | 0.8% | Y1–Y2 (1/yr) |
| 17 | x12 Driftrunners (Overhauled) | 300 | 12.00 | 3.60 | 0.7% | Y1–Y4 (3/yr) |
| 18 | ROM Stockpile | 5,000 | 1.00 | 5.00 | 1.0% | Y1–Y2 (0.5/yr) |
| 19 | x8 Jugs | 700 | 8.00 | 5.60 | 1.1% | Y1–Y4 (2/yr) |
| 20 | Pump Infrastructure | 20,000 | 1.00 | 20.00 | 3.8% | Y1–Y8 (0.125/yr) |
| 21 | Monorail | 10,000 | 1.00 | 10.00 | 1.9% | Y1–Y4 (0.25/yr) |
| 22 | x4 Mobile Fans | 1,000 | 4.00 | 4.00 | 0.8% | Y1–Y4 (1/yr) |
| 23 | Geo Studies | 5,000 | 3.00 | 15.00 | 2.9% | Y1:1, Y2:1, Y3:0.5, Y4:0.5 |
| 24 | Exploration Drilling | 5,000 | 1.00 | 5.00 | 1.0% | Y1 (1.0) |
| | **SUBTOTAL** | | | **521.20** | **100%** | |
| | **Contingency (30%)** | | | **156.36** | | |
| | **TOTAL WITH CONTINGENCY** | | | **677.56** | | |

---

## 4. Old Schedule Line Items (for comparison)

| # | Item | Unit Cost (AUD$k) | Total Units | Total Cost (AUDM) | % of Base |
|---|------|------------------:|------------:|-------------------:|----------:|
| 1 | Projects / Studies | 5,000 | 3.00 | 15.00 | 2.9% |
| 2 | Exploration Drilling | 5,000 | 3.00 | 15.00 | 2.9% |
| 3 | ROM Bin Upgrade | 15,000 | 1.50 | 22.50 | 4.3% |
| 4 | Ventilation Shafts | 20,000 | 2.00 | 40.00 | 7.7% |
| 5 | Mine Infrastructure | 5,000 | 2.00 | 10.00 | 1.9% |
| 6 | Drift | 12,000 | 2.00 | 24.00 | 4.6% |
| 7 | Drift Conveyor to ROM Stockpile | 17,000 | 2.00 | 34.00 | 6.6% |
| 8 | ROM Stockpile | 5,000 | 2.00 | 10.00 | 1.9% |
| 9 | UG Conveyors Relocation & Reuse | 20,000 | 2.00 | 40.00 | 7.7% |
| 10 | **Longwall** | **190,000** | **1.00** | **190.00** | **36.7%** |
| 11 | Mining Equipment | 69,900 | 1.00 | 69.90 | 13.5% |
| 12 | Infrastructure + CHPP EPCM | 24,000 | 2.00 | 48.00 | 9.3% |
| | **SUBTOTAL** | | | **518.40** | **100%** |
| | **Contingency (30%)** | | | **155.52** | |
| | **TOTAL WITH CONTINGENCY** | | | **673.92** | |

---

## 5. Year-by-Year Capital Profile

| Year | Calendar | New Base (AUDM) | New Total w/Cont (AUDM) | Old Base (AUDM) | Old Total w/Cont (AUDM) | Delta (Total) |
|------|---------|-----------------:|------------------------:|----------------:|------------------------:|--------------:|
| Y1 | 2027 | 132.05 | 171.67 | 10.00 | 13.00 | +158.67 |
| Y2 | 2028 | 94.55 | 122.92 | 10.00 | 13.00 | +109.92 |
| Y3 | 2029 | 83.80 | 108.94 | 59.00 | 76.70 | +32.24 |
| Y4 | 2030 | 82.80 | 107.64 | 84.00 | 109.20 | −1.56 |
| Y5 | 2031 | 63.75 | 82.88 | 109.90 | 142.87 | −59.99 |
| Y6 | 2032 | 59.25 | 77.03 | 221.50 | 287.95 | −210.92 |
| Y7 | 2033 | 2.50 | 3.25 | 24.00 | 31.20 | −27.95 |
| Y8 | 2034 | 2.50 | 3.25 | 0.00 | 0.00 | +3.25 |
| | **TOTAL** | **521.20** | **677.56** | **518.40** | **673.92** | **+3.64** |

**New schedule peak:** Y1 ($171.7M w/ contingency) — driven by ventilation network ($20M), roof supports ($50M), continuous miners ($4.5M), substations ($3M), AFC/BSL ($12M), drift ($3M), and exploration/studies all deploying simultaneously.

**Old schedule peak:** Y6 ($288.0M w/ contingency) — longwall package ($190M) + mining equipment ($69.9M) + infrastructure/CHPP EPCM ($24M) concentrated in a single year.

### Front-loading comparison:
- **New Y1–Y2 total:** $294.6M (w/ contingency) = **43.5%** of all capital in first 2 years
- **Old Y1–Y2 total:** $26.0M (w/ contingency) = **3.9%** of all capital in first 2 years
- **New Y1 alone ($171.7M)** exceeds the old schedule's first 5 years combined ($354.8M → no, old Y1–Y5 = $354.8M; actually old Y1–Y5 base = $272.9M, w/cont = $354.8M). New Y1 is 48% of old Y1–Y5 combined.

---

## 6. Present-Value Analysis

**Parameters:** Discount rate = 8% (from `Assumptions!B2`), base year = 2026, mid-year convention (`DF = 1/(1.08)^(year - 2026 + 0.5)`).

| Year | Calendar | DF | New Total w/Cont (AUDM) | PV (AUDM) | Old Total w/Cont (AUDM) | PV (AUDM) |
|------|---------|---|------------------------:|----------:|------------------------:|----------:|
| Y1 | 2027 | 0.9246 | 171.67 | 158.78 | 13.00 | 12.02 |
| Y2 | 2028 | 0.8561 | 122.92 | 105.23 | 13.00 | 11.13 |
| Y3 | 2029 | 0.7927 | 108.94 | 86.36 | 76.70 | 60.81 |
| Y4 | 2030 | 0.7340 | 107.64 | 79.01 | 109.20 | 80.15 |
| Y5 | 2031 | 0.6796 | 82.88 | 56.33 | 142.87 | 97.11 |
| Y6 | 2032 | 0.6293 | 77.03 | 48.48 | 287.95 | 181.21 |
| Y7 | 2033 | 0.5827 | 3.25 | 1.89 | 31.20 | 18.18 |
| Y8 | 2034 | 0.5395 | 3.25 | 1.75 | 0.00 | 0.00 |
| | | | **677.56** | **$518.2M** | **673.92** | **$443.8M** |

**PV delta (w/ contingency): +$74.4M** — the new schedule costs $74.4M more in today's dollars despite costing only $3.6M more nominally.

---

## 7. Top-5 Cost Drivers

### New Schedule:

| Rank | Item | Cost (AUDM) | % of Base |
|------|------|------------:|----------:|
| 1 | **x300 Roof Supports** | **300.00** | **57.6%** |
| 2 | x3 Continuous Miners | 27.00 | 5.2% |
| 3 | Ventilation Network | 20.00 | 3.8% |
| 4 | Pump Infrastructure | 20.00 | 3.8% |
| 5 | Geo Studies | 15.00 | 2.9% |

**Roof supports alone are 57.6% of all base capital.** At $1,000k ($1M) per unit × 300 units = $300M, this is the single most consequential assumption in the entire sheet. For context, the old workbook's largest item was Longwall at $190M (36.7%).

### Old Schedule:

| Rank | Item | Cost (AUDM) | % of Base |
|------|------|------------:|----------:|
| 1 | **Longwall** | **190.00** | **36.7%** |
| 2 | Mining Equipment | 69.90 | 13.5% |
| 3 | Infrastructure + CHPP EPCM | 48.00 | 9.3% |
| 4 | Ventilation Shafts | 40.00 | 7.7% |
| 5 | UG Conveyors Relocation & Reuse | 40.00 | 7.7% |

---

## 8. Analytical Assessment

### 8.1 What the new schedule does well:

1. **Granularity:** 24 equipment-level items provide far better visibility than 12 project-level buckets. This is the correct granularity for a pre-feasibility capital estimate and allows for individual component sensitivity analysis.

2. **Phasing realism:** Equipment is spread across 6–8 years with fractional phasing (0.25/yr, 0.5/yr, 0.125/yr), reflecting staged mine development and seasonal/quarterly deployment rather than single mega-purchases.

3. **Decomposes the old "Longwall" mega-item:** The $190M single-line longwall is broken into shearer ($11M), AFC ($6M), BSL ($6M), roof supports ($300M), etc. This is more defensible at the component level, though the total is higher (see concerns below).

4. **Includes exploration and studies:** Geo Studies ($15M) and Exploration Drilling ($5M) are present, addressing a gap noted in earlier reviews. The old schedule had $30M combined for these; the new schedule has $20M — a modest reduction.

5. **Pump infrastructure horizon:** The 8-year spread (0.125/yr) for pump infrastructure reflects ongoing water management — a realistic operational consideration for a gassy underground mine.

6. **CHPP component visibility:** Three separate CHPP items (structural repairs, conveyor, raw coal bin) totaling $19M provide much better granularity than the old schedule's implicit inclusion in "Infrastructure + CHPP EPCM" ($48M).

### 8.2 Concerns and Issues:

**A. Roof support cost assumption ($300M / 57.6% of total) — CRITICAL**
- 300 roof supports at $1,000k ($1M) each is an enormous line item — larger than the entire old longwall package ($190M).
- For reference, typical longwall shield costs are $200k–$500k per unit for standard shields; $1M per shield is at the very top of the range (high-capacity, 2m+ wide shields for thick seams). This should be sense-checked against supplier quotes.
- The 50/yr deployment over 6 years is consistent with multiple longwall panels, but the unit cost is the single most sensitive assumption in the model.
- **This single item drives the entire NPV delta.** Without roof supports, the new schedule is $221.2M (vs. $328.4M old ex-longwall) — the new schedule would be $107M cheaper nominally without this one item.

**B. Longwall components total exceeds old package**
- Old "Longwall" = $190M as a single package.
- New longwall components (shearer $11M + AFC $6M + BSL $6M + roof supports $300M = $323M) total **70% more** than the old package.
- The breakout is more granular, but the total is significantly higher. This suggests either: (a) the old package underpriced roof supports, (b) the new schedule overprices them, or (c) the new schedule includes more roof supports than the old package assumed.

**C. Front-loading risk — MATERIAL**
- New Y1–Y2: $294.6M (43.5% of total capital w/ contingency)
- Old Y1–Y2: $26.0M (3.9% of total capital w/ contingency)
- This creates a $268.6M shift of capital forward by 4–5 years, which is the sole driver of the $74.4M PV penalty.
- If the mine is still in development during Y1–Y2 (pre-production), this capital is at risk — significant early-stage financing required before revenue offsets costs.
- The old schedule's back-loaded profile was more PV-efficient: the $221.5M Y6 longwall purchase discounts to $181.2M, while the new schedule's $171.7M Y1 capital discounts only to $158.8M.

**D. Old schedule items partially or fully absent from new schedule:**
- "ROM Bin Upgrade" ($15M old → $9M "CHPP Raw Coal Bin" new) — scope reduction or different item
- "Drift Conveyor to ROM Stockpile" ($17M old) — **not present in new schedule**; the new "Drift" ($12M) may partially cover this but is a different line item
- "UG Conveyors Relocation & Reuse" ($20M old) — **not present**; partially captured in "Monorail" ($10M) but not equivalent
- "Mine Infrastructure" ($5M old) — partially captured across substations ($12M) and monorail ($10M) but not equivalent
- "Mining Equipment" ($69.9M old) — partially captured across shuttle cars ($4M), breaker feeder ($4M), driftrunners ($3.6M), jugs ($5.6M), mobile fans ($4M) = $21.2M new — **significantly lower** ($48.7M reduction)
- "Infrastructure + CHPP EPCM" ($48M old) — not present as a single item; CHPP items total $19M, but EPCM/management contingency is not separately captured

**E. The "missing" conveyor and infrastructure items** suggest the new schedule may be equipment-only and omits some civil/structural development capital. The old schedule had $17M (drift conveyor) + $20M (UG conveyors relocation) + $5M (mine infrastructure) = $42M of development capital that has no direct equivalent in the new schedule. However, the new schedule's Drift ($12M), Monorail ($10M), and Pump Infrastructure ($20M) may partially substitute.

**F. Mining Equipment reduction is large** — the old schedule's $69.9M "Mining Equipment" is replaced by ~$21.2M of individual equipment items. This is either: (a) the old schedule over-estimated equipment costs, (b) the new schedule under-estimates (missing equipment categories), or (c) different scope definitions. Worth confirming.

### 8.3 Nominal vs. PV Divergence — Key Analytical Insight:

| | Nominal | PV (8%, mid-year) |
|---|--------:|------------------:|
| New total w/ contingency | $677.6M | $518.2M |
| Old total w/ contingency | $673.9M | $443.8M |
| **Delta** | **+$3.6M (+0.5%)** | **+$74.4M (+16.8%)** |

The new schedule is **nearly identical in nominal terms** but **16.8% more expensive in present-value terms**. This is the classic "front-loading penalty" — capital brought forward costs more in PV terms even at lower or equal nominal amounts. The old schedule's Year 6 longwall purchase ($221.5M base) discounts to $139.5M; the new schedule's Year 1–2 capital ($226.6M base) discounts to $205.8M — a $66.3M PV difference from timing alone.

---

## 9. NPV Impact on Project Valuation

Using the workbook's own discount rate (8%) and base year (2026):

| Metric | New Schedule | Old Schedule | Delta |
|--------|-------------:|-------------:|------:|
| PV of capital (pre-contingency) | $398.6M | $341.4M | +$57.2M |
| PV of capital (w/ 30% contingency) | $518.2M | $443.8M | +$74.4M |

**Impact on project NPV:**

- Previous corrected Brave Blossom NPV: **$997M**
- PV capital increase (w/ contingency): +$74.4M
- **Estimated new Brave Blossom NPV: ~$923M** (a ~7.5% reduction)

This assumes revenue and operating cost schedules are unchanged. If the new equipment list implies different production profiles (e.g., continuous miners for development galleries + longwall for production), the revenue side would also shift and the net NPV impact could be different.

---

## 10. Recommendations for Revised Analysis

### Must-verify before incorporation:

1. **Roof support unit cost ($1,000k/unit):** Sense-check against supplier quotes or industry benchmarks. This is 57.6% of all capital and the most sensitive assumption. Test $500k, $750k, and $1,000k per unit as a sensitivity range. At $500k/unit, the schedule would be $371.2M base ($482.6M w/ contingency) — cheaper than the old schedule.

2. **Missing development capital:** Confirm whether drift conveyor to ROM stockpile ($17M old), UG conveyors relocation ($20M old), and mine infrastructure ($5M old) are captured elsewhere or genuinely omitted. If omitted, add ~$42M to the new schedule.

3. **Mining equipment scope reduction:** The old schedule's $69.9M mining equipment is replaced by ~$21.2M of individual items. Confirm this is a scope definition change, not an omission.

4. **Production method:** The new schedule includes both continuous miners (3 units, $27M) and longwall equipment (shearer, AFC, BSL, roof supports). Confirm this is a dual-method mine (CM development galleries + longwall production panels). If so, the revenue schedule should reflect both phases.

5. **Front-loading feasibility:** Can the schedule be re-phased? Moving $50M of roof supports from Y1–Y2 to Y3–Y4 would save ~$8M in PV. Evaluate whether 50/yr deployment is a hard constraint or can be optimized.

### For the revised analysis session:

- **Adopt the new capital schedule** as the primary basis — it is more defensible at the line-item level
- **Flag the roof support cost** as the #1 sensitivity: test $500k, $750k, and $1,000k per unit
- **Re-compute NPV** with the new capital schedule replacing the old one. Expected result: NPV drops from ~$997M to ~$923M (assuming unchanged revenue/opex)
- **Re-examine the revenue side** to ensure consistency with the equipment list (CM development + longwall production dual-method)
- **Consider adding back** missing development items (drift conveyor, UG conveyors) if confirmed absent — this would add ~$42M nominal (~$35M PV)
- **Test front-loading sensitivity:** The PV penalty of early capital is $74.4M; evaluate whether a more staged deployment is feasible
- **Update the Option A walkthrough** (`OPTION_A_CALCULATION_WALKTHROUGH.md`) with the revised capital figures
- **Update analysis docs** (02_brave_blossom_swot.md, 05_recommendation_gonogo.md, 07_qa_preparation.md) to reflect the new capital basis

---

## 11. Data Quality Notes

- **Cached values present:** Unlike the old populated workbook (which had no cached values), the new Team Green workbook **has cached formula values** — it was opened and saved in Excel. All computed figures in this review are from the workbook's own cached results, cross-verified by independent Python recomputation. All values match exactly.
- **Formula structure is clean:** All cost rows (31–54) use the consistent pattern `=$C{row}*{col}{row}`, with column C referencing the unit cost in the input section. The total row (55) uses `=SUM(D31:D54)`, contingency (56) uses `=$C$56*D55`, and total (57) uses `=D56+D55`. No formula inconsistencies detected.
- **Year header chain:** Row 1 years are formula-driven: `D1="='Brave Blossom '!D2"` → `Assumptions!J1` = 2027, then `E1=D1+1`, etc. through AD1=2053. Clean chain.
- **Template rows:** Rows 43–54 in the cost section reference input rows 16–27 via `=A{row}` formulas. Rows 16–21 in the old workbook had "Other Items (as required)" placeholders with no costs; the new workbook uses all rows 4–27 for real items.
- **Typo in item name:** "Pump Infrastruture" (row 23) — should be "Infrastructure". Minor but should be corrected for presentation materials.
- **No data validation rules** on the input section (rows 4–27) — the workbook relies on manual entry correctness.

---

## 12. Quick Reference Cell Map

| Cell | Sheet | Content | Notes |
|------|-------|---------|-------|
| B2 | Assumptions | 0.08 | Discount rate (8%) |
| J1 | Assumptions | 2027 | Base year for Brave Blossom (I1+1) |
| C56 | BB Capital | 0.3 | Contingency rate (30%) |
| D1:AD1 | BB Capital | 2027–2053 | Year headers (formula chain from Assumptions) |
| A4:A27 | BB Capital | Item names | 24 equipment-level line items |
| C4:C27 | BB Capital | Unit costs (AUD$k) | Range: $300k–$20,000k |
| D4:AD27 | BB Capital | Unit counts by year | Fractional values (0.125–300) |
| D31:AD54 | BB Capital | `=$C{r}*{col}{r}` | Computed costs per item per year |
| D55:AD55 | BB Capital | `=SUM(D31:D54)` | Annual totals (cached: $132,050 to $0) |
| D56:AD56 | BB Capital | `=$C$56*D55` | Annual contingency (cached: $39,615 to $0) |
| D57:AD57 | BB Capital | `=D56+D55` | Annual total w/ contingency (cached: $171,665 to $0) |
