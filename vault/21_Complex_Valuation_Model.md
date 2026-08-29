# Complex Valuation Model — Workbook Analysis

> **Source**: `Complex_Valuation_Model.xlsx` (provided as supplement to the 2026 Graduate Challenge)
> **Sheets**: 15 (corrected from initial estimate of 16 — "Brave Blossom (2)" does not exist in the actual workbook) | **Time span**: 2011 (historical) → 2075 (carbon/decommissioning)
> **Created**: 2026-08-29 | **Updated**: 2026-08-29 (audited against actual workbook file)
> **Cross-refs**: Files 03 (Springbok), 04 (Brave Blossom), 05 (Key Considerations), 07 (Capital Estimate), 15 (Brumby OC), 17 (Closure Costs)

---

## Overview

This workbook is the **master financial model** for the Wallaby Mining / Waratah Resources JV. It contains full DCF valuations for three operating scenarios — Springbok (existing UG mine), Brave Blossom (proposed UG extension), and an open-cut proposal ("OC Clean" / SandsEnd) — plus historical performance data, corporate assumptions, carbon pricing, and post-closure methane emissions modelling.

### Key NPV Results (from Analysis sheet, 8% discount rate)

| Scenario | NPV (AUD M) |
|---|---|
| **Springbok standalone (base case)** | **279** |
| **Brave Blossom UG standalone** | **1,320** |
| **Combined (Springbok + Brave Blossom)** | **1,670** |
| OC Clean (SandsEnd open cut) | 0 (template unpopulated) |

The combined cumulative DCF turns positive in 2032 (~AUD 295M) after absorbing the AUD 895M Springbok closure hit in 2032 and Brave Blossom capex ramp. It reaches **AUD 1.67B by 2052**.

---

## Sheet-by-Sheet Breakdown

> **Corrected from initial vault analysis**. The actual workbook has 15 sheets (not 16). "Brave Blossom (2)" does not exist. "SpringbokCapital" is a separate sheet (was previously described as part of the Springbok sheet). Sheet name "Brave Blossom " has a trailing space in the actual file.

### 1. Assumptions (58 rows × 74 cols, 2019 populated cells)
**Purpose**: Master input sheet driving all calculations. Year columns 2021–2091 (71 years).

| Parameter | Value |
|---|---|
| Discount rate | 8% |
| FX rate (USD:AUD) | 0.69 (flat from 2026+) |
| PHCC price (USD/t) | $295–360 historical (2021–2025), $240 flat forecast (2026+) |
| 88% PHCC benchmark | $211.20/t (= 88% × $240) |
| Thermal GCN (6000NAR) | $118/t |
| Thermal 5500NAR | $105/t |
| GAR→NAR adjustment | 260 kcal/kg |
| Price basis | 6000 kcal/kg (NAR) |
| Carbon scenario | "Accelerated Transition" ($45–200/t, see Carbon sheet) |
| QLD royalty | Tiered: $0–100/t @ 7%, $100–150/t @ 12.5%, $150–175/t @ 15%, $175–225/t @ 20%, $225–300/t @ 30%, >$300/t @ 40% |
| Royalty on first threshold | AUD 7.00/t (= $100 × 7%) |

### 2. Analysis (24 rows × 37 cols, 404 populated cells)
**Purpose**: High-level DCF summary and comparison across scenarios. Years 2021–2052.

Key outputs:
- **FOR cash cost/t**: Springbok ~AUD 135–169/t (historical), ~AUD 144–156/t (forecast). Brave Blossom: AUD 1088/t in ramp year 2031 (low volume), then ~AUD 120/t steady state
- **Net cashflow before tax**: Springbok generates AUD 350–546k p.a. (2021–2031), then −AUD 895M in 2032 (closure). Brave Blossom: −AUD 76k–174k during construction (2031–2032), then AUD 323–453k p.a. from 2033
- **Attributable (after-tax, 60% JV share) cashflow**: Combined peaks at ~AUD 210k (2022), dips to −AUD 641k (2032 closure year), recovers to ~AUD 135–182k p.a.
- **Cumulative DCF**: Combined reaches AUD 1.67B by 2052. Springbok-only peaks at AUD 279M (2027) then declines to AUD 279M flat post-closure
- **Brave Blossom standalone cumulative**: AUD 1,390M by 2052

### 3. List (4 rows × 8 cols, 20 populated cells)
**Purpose**: Dropdown validation lists for the Springbok Assumptions sheet.
- Owner/operate vs Contract
- Roster: 7on7off, 5on2off, Other
- Priority: Max Quality (4), Max Quantity (3), (2), (1)
- Gas drainage: SIS, UIS
- Progressive rehab: Yes/No
- Maintenance: Preventative, Corrective, Increase 10%, Decrease 10%

### 4. Summary (57 rows × 146 cols, 546 populated cells)
**Purpose**: Dashboard sheet for revenue and cost graphs. Two sections:

**Section 1 — Springbok Historical + Forecast (2022–2031)**:
- Sales revenue: AUD 1,596–2,073M (historical), declining to AUD 1,210–1,345M (forecast)
- Average realisation: AUD 431–505/t (historical), AUD 348/t (forecast)
- Cost breakdown (AUD/t): Direct mining ~105–137, CHPP ~19–28, Overhead ~17–21, Distribution ~24–128, Carbon $0–8, Royalty AUD 52–98

**Section 2 — Brave Blossom Project Margin (2031–2050)**:
- Revenue: ramps from AUD 25M (2031) to AUD 1,652M p.a. (2033–2048), AUD 1,582M (2049–2050)
- Realisation: AUD 306/t (88% PHCC at $211.20 USD × FX 0.69)
- Direct mining: AUD 90/t steady state (AUD 676–823/t during ramp)
- CHPP: AUD 17.7/t, Overhead: AUD 12.8/t, Distribution: AUD 22/t
- Carbon cost rises from AUD 12/t (2033) to AUD 27/t (2050)
- Royalty: AUD 55–80/t

### 5. Historical Performance (170 rows × 9 cols, 851 populated cells)
**Purpose**: Springbok Complex actuals 2011–2025 (15 years).

Covers: ROM production, strip ratio, CHPP feed/yield/product, marketing splits (PHCC, GCN), coal quality (energy, ash), costs per tonne, revenue. Provides the basis for forecast calibration.

### 6. Springbok Assumptions (10 rows × 3 cols, 26 populated cells)
**Purpose**: Scenario toggle — Current vs Proposed operating configuration.

| Parameter | Current | Proposed |
|---|---|---|
| Roster — Longwall | 7on7off | 7on7off |
| Roster — Development | 7on7off | **5on2off** |
| Roster — Outbye | 7on7off | 7on7off |
| Development miners | 3 | **4** |
| Gas drainage strategy | SIS | **UIS** |
| Wash strategy | Max Quality | Max Quality |
| Progressive rehabilitation | Yes | **No** |
| Maintenance strategy | Preventative | **Corrective** |

> The "Proposed" column appears to model a **cost-optimised / life-extension** scenario: faster development (5-on-2-off + extra miner), upgraded gas drainage (UIS), deferred rehabilitation, and shift from preventive to corrective maintenance. These levers directly affect the closure liability and life-of-mine economics.

### 7. Springbok (161 rows × 35 cols, 2751 populated cells)
**Purpose**: Full operational + financial model for Springbok Complex, 2027–2053.

Key parameters:
- **Reserves**: 20,000 kt
- **Total ROM**: 27,857 kt over 5 years (2027–2031)
- **CHPP yield**: 67%
- **Saleable**: 18,664 kt
- **Product**: 100% PHCC (no thermal)
- **Realised price**: USD 240/t (= AUD 348/t at 0.69 FX)
- **Scope 1 emissions**: 3.44M tCO₂e (coal mine waste gas @ 0.12 t/t ROM)
- **Safeguard baseline**: Hybrid declining, starts ~426k tCO₂e (2027) declining to ~293k (2031)
- **Emissions above baseline**: 217–357k tCO₂e p.a. — generates carbon liability
- **NPV**: AUD 279M
- **Mine life**: 5 years (2027–2031), closure in 2032

Capital schedule is on the separate **SpringbokCapital** sheet (see sheet 8 below). AUD 122,800k base + 20% contingency = **AUD 147,360k total** over 2027–2031. Items: replacement miner, ventilation shaft, SIS drill program, roof supports, conveyor belt, community program, TSF, exploration.

### 8. SpringbokCapital (48 rows × 22 cols, 465 populated cells)
**Purpose**: Capital schedule for Springbok Complex, 2027–2031. Separate sheet (was previously described as part of sheet 7). Fully populated with capital items, unit costs, and scheduling across years. Items: replacement miner, ventilation shaft, SIS drill program, roof supports, conveyor belt, community program, TSF, exploration.

### 9. Brave Blossom (151 rows × 72 cols, 6145 populated cells)
**Purpose**: Full operational + financial model for Brave Blossom UG project, 2027–2071.

Key parameters:
- **Development**: 255 km total (15 km/yr × 17 years, ramp 2031–2032)
- **Total ROM**: 142,568 kt (dev 5,768 + longwall 136,800)
- **CHPP yield**: 68% (48% during ramp)
- **Saleable**: 96,844 kt (88% PHCC only)
- **Realised price**: USD 211.20/t (88% of PHCC benchmark)
- **Scope 1 emissions**: 23.6M tCO₂e total (operational 18.9M + decommissioned mine 4.7M)
- **Safeguard baseline**: Declining from ~9,475 tCO₂e (2031) to ~zero (2050)
- **Emissions above baseline**: 20.1M tCO₂e — significant carbon liability
- **NPV**: AUD 1,320M
- **Mine life**: ~20 years production (2033–2052), post-closure emissions to 2071

### 10. Brave Blossom Capital (66 rows × 30 cols, 882 populated cells)
**Purpose**: Capital schedule template for Brave Blossom, 2027–2053.

**Status**: Has structural framework (year headers via formula, 12 capital items defined with unit costs AUD 5k–190k) but **unit-number scheduling rows are all zero** — the capital items are listed with their unit costs, but the number of units per year has not been entered. Needs population of unit quantities by year to complete the capital schedule.

12 capital items defined with unit costs (AUD 5k–190k):
| Item | Unit Cost (AUDk) |
|---|---|
| (5 items @ $5k each) | 5,000 |
| (1 item @ $15k) | 15,000 |
| (1 item @ $20k) | 20,000 |
| (1 item @ $12k) | 12,000 |
| (1 item @ $17k) | 17,000 |
| (1 item @ $19k) | 19,000 |
| Major item | 190,000 |
| (1 item @ $69.9k) | 69,900 |
| (1 item @ $24k) | 24,000 |

Contingency: 30%. **Unit-number scheduling rows are zero** — template framework exists (882 cells with formulas/labels) but capital quantities by year not yet assigned.

### 11. Historical Performance BrumbyOC (162 rows × 8 cols, 107 populated cells)
**Purpose**: Brumby Open Cut historical actuals, 2011–2015 (5 years).

Key data:
- ROM: 1,000–1,349 kt/yr
- Strip ratio: 4.5–4.9 bcm/t
- CHPP yield: 69–75%
- Product: PHCC (270–310 kt) + GCN (450–697 kt)
- GCN energy: 6,212–6,290 kcal/kg (gar), ash 9.2–10.1%
- Provides benchmark data for the OC Clean proposal

### 12. OC Clean (154 rows × 35 cols, 2684 populated cells)
**Purpose**: Open cut proposal model — "OC - Proposal" / SandsEnd. 2027–2052.

Structure mirrors Springbok sheet but for an open cut operation:
- Reserves: 20,000 kt
- Mining: T&S + Dragline waste removal, blasted prime waste
- Carbon: Scope 1 @ 0.1 tCO₂e/t ROM, safeguard declining baseline, 21-year mine life
- Products: Coking Product A + Thermal Product A
- **Has formula structure (2684 populated cells)** but operational values are zero/template — formula framework ready for input.

### 13. OC Clean Capital (71 rows × 25 cols, 801 populated cells)
**Purpose**: Capital schedule for SandsEnd open cut, 2027–2048.

26 capital items defined:
| Key Items | Unit Cost (AUDk) |
|---|---|
| Project cost (feasibility/exploration) | 8,000 |
| Excavators (300t–800t) | 5k–20k each |
| Trucks (130t–320t) | 3k–8k each |
| Dragline | 100,000 |
| Dozers (D10, D11) | 2k–3k |
| ROM Pad | 20,000 |
| CHPP Upgrade | 50,000 |
| 25 km Powerline | 20,000 |
| Substation | 10,000 |
| Dry Processing Plant | 20,000 |
| Creek Crossing Civil Work | 20,000 |
| Water Management Dams | 7,500 |

**Cost scheduling rows are zero** — template framework exists (801 cells with formulas/labels) but quantities by year not yet assigned.

### 14. Carbon (6 rows × 51 cols, 153 populated cells)
**Purpose**: Carbon price forecast scenarios, 2026–2075.

Two scenarios:
| Scenario | Price Range |
|---|---|
| **Zero Carbon Liability** | $0/t (all years) |
| **Accelerated Transition** | $45/t (2026) → $200/t (2075) |

The "Accelerated Transition" curve: ramps steeply to $120/t by 2038, plateaus, then climbs again to $200/t from 2064+. This is the active scenario (referenced in Assumptions cell C20).

### 15. Decommissioned Mine (36 rows × 134 cols, 987 populated cells)
**Purpose**: Post-closure methane (CH₄) emissions model per NGER Method 3.32 (Part 2). Calculates fugitive emissions from sealed/abandoned underground mines.

Three closure scenarios modelled:

| Scenario | Mine Closed | Stop Reporting | CH₄ pre-closure (tCO₂e) | Mine Void (m³) | Flood Rate |
|---|---|---|---|---|---|
| Springbok only | 31 Dec 2031 | 2051 | 649,512 | 21,159,588 | 74,000 m³/yr |
| Springbok + Brave Blossom | 31 Dec 2050 | 2070 | 1,008,000 | 42,319,176 | 74,000 m³/yr |
| Springbok Delayed 12 Months | 31 Dec 2032 | 2052 | 324,756 | 21,159,588 | 74,000 m³/yr |

All classified "Gassy". Emissions decline exponentially post-closure as the mine void floods (74,000 m³/yr), reducing the unflooded volume available for CH₄ release. Springbok+Brave Blossom generates **4.68M tCO₂e** of decommissioned mine emissions over ~20 years.

Key inputs: mine void volume, flood rate, gassy classification, CH₄ emissions in 12 months pre-closure, NGER emission factor formula (3.32). The "Months (T)" row tracks elapsed months since sealing (0 pre-closure, 12–252 post-closure).

---

## Structural Relationships

```
Assumptions ──drives──→ All operational sheets
    │
    ├── Springbok (NPV $279M) ←── Springbok Assumptions (scenario toggle)
    │       └── SpringbokCapital ($147M capex, fully populated)
    │
    ├── Brave Blossom (NPV $1,320M) — sheet name has trailing space
    │       └── Brave Blossom Capital (framework exists, scheduling rows zero)
    │
    ├── OC Clean (formula structure, values zero) ←── OC Clean Capital (framework, scheduling zero)
    │
    ├── Carbon → feeds carbon cost into all models
    │
    ├── Decommissioned Mine → feeds post-closure CH₄ into Brave Blossom & Springbok
    │
    ├── Analysis ← aggregates Springbok + Brave Blossom DCFs
    │       └── Summary ← dashboard graphs
    │
    ├── List → dropdown validation for Springbok Assumptions
    ├── Historical Performance (Springbok 2011–2025)
    └── Historical Performance BrumbyOC (Brumby 2011–2015)
```

> **Note**: "Brave Blossom (2)" was listed in the initial vault analysis but does NOT exist in the actual workbook. The alternative decommissioning scenario described may have been removed or never existed in the final file version provided.

---

## Key Observations

1. **The model is partially populated**: Springbok and Brave Blossom (sheet 9) are fully built with formulas. OC Clean, OC Clean Capital, and Brave Blossom Capital have structural frameworks (formula templates, item lists, year headers) but the actual input quantities are zero — awaiting scheduling.

2. **"Brave Blossom (2)" does not exist** in the actual workbook. The initial vault analysis (based on pre-workbook information) described an alternative decommissioning scenario sheet. This sheet is not present in the provided file. The Decommissioned Mine sheet (15) still models three closure scenarios independently.

3. **The AUD 895M closure hit** appears in 2032 in the Analysis sheet (row 7, column N = −895,000). This aligns with the $900M ERC closure cost estimate from the 2024 closure workbook (see File 17).

4. **Carbon cost is material**: At the "Accelerated Transition" price curve, carbon adds AUD 2.8–8.1/t to Springbok costs (2026–2031) and AUD 12–27/t to Brave Blossom (2033–2050). The emissions-above-baseline is substantial: 1.6M tCO₂e for Springbok (5 years) and 20.1M tCO₂e for Brave Blossom (20 years).

5. **Decommissioned mine emissions** are a hidden liability: 4.68M tCO₂e over ~20 years post-closure for the combined Springbok + Brave Blossom scenario. At $150/t carbon price (mid-range), this represents ~AUD 700M in carbon liability — **a material amount not reflected in the headline NPV** unless the model already captures it via the emissions-above-baseline rows.

6. **The "Proposed" Springbok scenario** (deferred rehab, corrective maintenance, UIS gas drainage, extra development miner) appears designed to model cost reductions for the closure liability task. The shift from SIS to UIS gas drainage and from preventive to corrective maintenance are direct levers on operating cost and closure provision.

7. **Royalty structure** is the QLD tiered ad valorem system: 7% on first $100/t, 12.5% to $150, 15% to $175, 20% to $225, 30% to $300, 40% above $300. At $240/t PHCC, the effective royalty rate is significant (blended ~25-29%).

8. **Brave Blossom product is 88% PHCC** at $211.20/t (88% of the $240/t benchmark), not full-spec PHCC. This is a **33% lower price realisation** than Springbok's $240/t PHCC — a critical assumption affecting the AUD 1.32B NPV. Sensitivity analysis on this parameter would be valuable.

9. **SpringbokCapital is a separate sheet** (sheet 8 in the actual workbook), not embedded in the Springbok sheet. It is fully populated with 465 cells covering capital items, unit costs, and year-by-year scheduling. This is a correctly built capital schedule that can serve as a template for the Brave Blossom Capital tab.

10. **OC Clean and OC Clean Capital have formula frameworks** (2684 and 801 populated cells respectively) but operational/cost values are zero. They are not truly "blank" — the formula structure exists, just needs input data.
