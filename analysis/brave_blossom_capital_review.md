# Brave Blossom Capital Sheet — Review & Assessment

**Source:** `Complex_Valuation_Model_-_Team_Green_1_1.xlsx` (Team Green submission)
**Sheet:** `Brave Blossom Capital` (66 rows × 30 cols)
**Comparison baseline:** Previous populated workbook (`Complex_Valuation_Model_POPULATED.xlsx`)
**Date:** 2026-09-02

---

## 1. Structural Overview

The new Team Green workbook **completely replaces** the capital items in the Brave Blossom Capital sheet. The old schedule had 12 high-level line items (projects, exploration, ROM bin, ventilation, drift, longwall package, etc.). The new schedule has **22 granular equipment-level line items** with a fundamentally different capital philosophy.

### Key structural changes:
- **Old approach:** Project-level / phase-based (e.g., "Longwall" = $190M single line, "Mining Equipment" = $69.9M single line)
- **New approach:** Equipment-level / unit-based (e.g., Shearer $11M, x300 Roof Supports $300M, x3 Continuous Miners $27M, CHPP components broken out individually)
- **Year columns:** Both workbooks use 27 year columns (D:AD), starting calendar year 2027 (referencing `Assumptions!J1` → `I1+1` where I1=2026)
- **Contingency:** Both use 30% (cell C56=0.3)
- **Formula structure:** Identical computation framework — rows 4-25 are inputs, rows 31-52 compute costs (`=$C{row}*{col}`), row 55 sums, row 56 applies contingency, row 57 totals

---

## 2. Complete Line Item Inventory (New Workbook)

| # | Item | Unit Cost (AUD$k) | Total Units | Total Cost (AUDM) | Year Spread |
|---|------|------------------:|------------:|-------------------:|-------------|
| 1 | Shearer | 11,000 | 1.00 | 11.00 | Y1–Y4 (0.25/yr) |
| 2 | CHPP Structural Repairs | 5,000 | 1.00 | 5.00 | Y2–Y3 (0.5/yr) |
| 3 | CHPP Conveyor | 5,000 | 1.00 | 5.00 | Y4–Y5 (0.5/yr) |
| 4 | CHPP Raw Coal Bin | 9,000 | 1.00 | 9.00 | Y3–Y6 (0.25/yr) |
| 5 | Ventilation Network | 20,000 | 1.00 | 20.00 | Y1 (1.0) |
| 6 | Gas Drainage | 12,000 | 1.00 | 12.00 | Y1–Y2 (0.5/yr) |
| 7 | x3 Continuous Miners | 9,000 | 3.00 | 27.00 | Y1–Y6 (0.5/yr) |
| 8 | **x300 Roof Supports** | **1,000** | **300.00** | **300.00** | Y1–Y6 (50/yr) |
| 9 | x10 DCBs (AFC, BSL, Dewater, CMEs) | 1,000 | 10.00 | 10.00 | Y1–Y5 (2/yr) |
| 10 | x4 U/G Substations | 3,000 | 4.00 | 12.00 | Y1–Y4 (1/yr) |
| 11 | x10 HT Cables | 1,000 | 10.00 | 10.00 | Y1:5, Y2:2, Y3:2, Y4:1 |
| 12 | AFC (Panline & Chain) | 6,000 | 1.00 | 6.00 | Y1 (1.0) |
| 13 | BSL | 6,000 | 1.00 | 6.00 | Y1 (1.0) |
| 14 | Pipes (Water, Air, Nitrogen) | 10,000 | 1.00 | 10.00 | Y1–Y4 (0.25/yr) |
| 15 | x4 Shuttle Cars | 1,000 | 4.00 | 4.00 | Y1–Y4 (1/yr) |
| 16 | x2 Breaker Feeder | 2,000 | 2.00 | 4.00 | Y1–Y2 (1/yr) |
| 17 | x12 Driftrunners (Overhauled) | 300 | 12.00 | 3.60 | Y1–Y4 (3/yr) |
| 18 | Shearer Carrier | 1,000 | 1.00 | 1.00 | Y1–Y2 (0.5/yr) |
| 19 | x8 Jugs | 700 | 8.00 | 5.60 | Y1–Y4 (2/yr) |
| 20 | Pump Infrastructure | 20,000 | 1.00 | 20.00 | Y1–Y8 (0.125/yr) |
| 21 | Monorail | 10,000 | 1.00 | 10.00 | Y1–Y4 (0.25/yr) |
| 22 | x4 Mobile Fans | 1,000 | 4.00 | 4.00 | Y1–Y4 (1/yr) |
| | **SUBTOTAL** | | | **495.20** | |
| | **Contingency (30%)** | | | **148.56** | |
| | **TOTAL WITH CONTINGENCY** | | | **643.76** | |

---

## 3. Head-to-Head Comparison: New vs. Old

| Metric | New (Team Green) | Old (Populated) | Delta |
|--------|-----------------:|----------------:|------:|
| Line items | 22 | 12 | +10 |
| Nominal subtotal | $495.20M | $518.40M | −$23.20M |
| Nominal w/ 30% contingency | $643.76M | $673.92M | −$30.16M |
| PV of capex (pre-cont, 8% DF) | $376.87M | $341.40M | +$35.47M |
| PV of capex (w/ 30% cont, 8% DF) | $489.93M | $443.82M | +$46.11M |
| Capital spread (years) | Y1–Y8 | Y1–Y7 | +1 yr |

### Critical observation — nominal vs. PV divergence:
The new schedule is **$23M cheaper nominally** but **$46M more expensive in present-value terms**. This is because the new schedule front-loads capital: $119.6M in Year 1 alone (vs. $10M old), and $287.4M across Y1–Y2 (vs. $20M old). The old schedule peaked in Year 6 ($221.5M for the longwall package), which discounts heavily. **Front-loading destroys PV even at lower nominal cost.**

---

## 4. Year-by-Year Capital Profile

| Year | Calendar | New Capex (AUDM) | Old Capex (AUDM) | Delta |
|------|---------|-----------------:|-----------------:|------:|
| Y1 | 2027 | 119.55 | 10.00 | +109.55 |
| Y2 | 2028 | 87.05 | 10.00 | +77.05 |
| Y3 | 2029 | 80.80 | 59.00 | +21.80 |
| Y4 | 2030 | 79.80 | 84.00 | −4.20 |
| Y5 | 2031 | 63.75 | 109.90 | −46.15 |
| Y6 | 2032 | 59.25 | 221.50 | −162.25 |
| Y7 | 2033 | 2.50 | 24.00 | −21.50 |
| Y8 | 2034 | 2.50 | 0.00 | +2.50 |

**New schedule peak:** Y1 ($119.6M) — driven by ventilation, roof supports, continuous miners, and infrastructure all deploying simultaneously.

**Old schedule peak:** Y6 ($221.5M) — longwall package + mining equipment concentrated in a single year.

---

## 5. Top-5 Cost Drivers (New Schedule)

| Rank | Item | Cost (AUDM) | % of Subtotal |
|------|------|------------:|--------------:|
| 1 | x300 Roof Supports | 300.00 | 60.6% |
| 2 | x3 Continuous Miners | 27.00 | 5.5% |
| 3 | Ventilation Network | 20.00 | 4.0% |
| 4 | Pump Infrastructure | 20.00 | 4.0% |
| 5 | Gas Drainage | 12.00 | 2.4% |

**Roof supports alone are 60.6% of all capital.** At $1,000k each × 300 units = $300M, this is the single most consequential assumption in the entire sheet. For context, the old workbook's largest item was the Longwall at $190M (36.7% of old subtotal).

---

## 6. Analytical Assessment

### 6.1 What the new schedule does well:
1. **Granularity:** 22 equipment-level items provide far better visibility than 12 project-level buckets. This is the correct granularity for a pre-feasibility capital estimate.
2. **Phasing realism:** Equipment is spread across 6–8 years, reflecting staged mine development rather than a single mega-purchase. The 0.25/yr fractional phasing on major items (shearer, CHPP bin, pipes, monorail, pump infrastructure) suggests quarterly or seasonal deployment.
3. **Removes old "Longwall" mega-item:** The $190M single-year longwall is replaced with component-level breakdown (shearer $11M, AFC $6M, BSL $6M, roof supports $300M, etc.), which is more defensible.
4. **Extends capital horizon to Y8:** Pump infrastructure's 8-year spread reflects ongoing water management — a realistic operational consideration.

### 6.2 Concerns and issues:

**A. Roof support cost assumption ($300M / 60.6% of total)**
- 300 roof supports at $1,000k ($1M) each is an enormous line item. For reference, typical longwall shield costs are $200k–$500k per unit for standard shields; $1M per shield is at the very top of the range (high-capacity, 2m+ wide shields for thick seams).
- The 50/yr deployment over 6 years is consistent with multiple longwall panels, but the unit cost should be sense-checked against actual quotes.
- **This single assumption drives the entire NPV delta.** Without roof supports, the new schedule is $195.2M (vs. $328.4M old ex-longwall).

**B. No "Projects / Studies" or "Exploration" lines**
- The old schedule had $30M for pre-mining studies and exploration. The new schedule has zero. If this is a producing mine (Brave Blossom UG), this may be appropriate — but if it's a greenfield/development project, the absence of pre-feasibility and exploration capex is a gap.

**C. No "Longwall" as a packaged item**
- The old workbook had a $190M "Longwall" line. The new workbook breaks this into shearer ($11M), AFC ($6M), BSL ($6M), roof supports ($300M), shearer carrier ($1M), etc. The components sum to ~$324M — **70% more than the old $190M package.** The breakout is more granular but the total is significantly higher.

**D. Missing items from old schedule:**
- ROM Bin Upgrade ($15M old → $9M "CHPP Raw Coal Bin" new) — possible reduction
- Drift ($12M old) — not present
- Drift Conveyor to ROM Stockpile ($17M old) — not present
- ROM Stockpile ($5M old) — not present
- UG Conveyors Relocation & Reuse ($20M old) — partially captured in "Pipes" ($10M new) but not equivalent
- Mine Infrastructure ($5M old) — partially captured in substations/monorail but not equivalent
- Mining Equipment ($69.9M old) — partially captured across shuttle cars, breaker feeder, driftrunners, jugs, mobile fans (~$21.2M new) — significantly lower

**E. The "missing" drift, ROM stockpile, and conveyor infrastructure items** suggest the new schedule may be equipment-only and omits civil/structural development capital. If so, the $495.2M subtotal **understates** total project capital.

**F. CHPP scope reduced:** Old workbook had no explicit CHPP breakdown. New workbook has 3 CHPP items totaling $19M (structural repairs, conveyor, raw coal bin). This is reasonable for a brownfield CHPP upgrade but low if a new CHPP is required.

### 6.3 Front-loading risk:
The new schedule deploys $206.6M in Y1–Y2 (42% of subtotal) vs. $20M old (3.9%). This creates a significantly worse PV profile and implies heavy early-stage financing requirements. If the mine is still in development during Y1–Y2, this capital is at risk (pre-production capex with no offsetting revenue).

---

## 7. NPV Impact on Project Valuation

Using the workbook's own discount rate (8%) and base year (2026):

| Metric | New Schedule | Old Schedule | Delta |
|--------|-------------:|-------------:|------:|
| PV of capital (pre-contingency) | $376.87M | $341.40M | +$35.47M |
| PV of capital (w/ 30% contingency) | $489.93M | $443.82M | +$46.11M |

**Impact on project NPV:** The new capital schedule would **reduce Brave Blossom project NPV by ~$46M** (with contingency) compared to the old schedule. Applied to the previous corrected NPV of $997M, this would bring Brave Blossom NPV to approximately **$951M** (a ~4.6% reduction).

However, this assumes the revenue/operating cost schedules are unchanged. If the new equipment list implies different production profiles (e.g., continuous miners vs. longwall), the revenue side would also shift.

---

## 8. Recommendations for Revised Analysis

### Must-verify before incorporation:
1. **Roof support unit cost ($1,000k/unit):** Sense-check against supplier quotes or industry benchmarks. This is 60% of all capital and the most sensitive assumption.
2. **Missing civil/development capital:** Confirm whether drift development, ROM stockpile construction, and surface infrastructure are captured elsewhere (Springbok sheet? A different capital sheet?) or genuinely omitted.
3. **Production method:** The new schedule includes both continuous miners (3 units) and longwall equipment (shearer, AFC, BSL, roof supports). Is this a dual-method mine (CM development + longwall production)? If so, the revenue schedule should reflect both phases.
4. **Exploration/studies gap:** Confirm whether Brave Blossom is a producing mine (no exploration capex needed) or development-stage (requires $20-30M pre-feasibility capex).

### For the revised analysis session:
- **Adopt the new capital schedule** as the primary basis — it is more defensible at the line-item level
- **Flag the roof support cost** as a key sensitivity: test $500k, $750k, and $1,000k per unit
- **Re-compute NPV** with the new capital schedule replacing the old one
- **Re-examine the revenue side** to ensure consistency with the equipment list (continuous miner + longwall dual-method)
- **Consider adding back** missing development items (drift, ROM stockpile) if they are confirmed absent
- **Test front-loading sensitivity:** The PV penalty of early capital is $46M; evaluate whether a more staged deployment is feasible

---

## 9. Data Quality Notes

- Workbook has **no cached formula values** (populated programmatically, never opened in Excel). All values above were computed from the formula structure.
- The `data_only=False` flag was used to read formulas; a LibreOffice headless recalculation pass may be needed to embed cached values for downstream use.
- Cell A53 references `=A26` and row 53 formulas reference row 26, but row 26 is empty — this creates a phantom row of zeros in the cost section. Not a material issue but indicates the template was built for more items than were populated.
- No data validation rules on the input section (rows 4–25) — the new workbook relies on manual entry correctness.
