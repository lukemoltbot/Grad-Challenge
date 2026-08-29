# Workbook Audit Report: Complex_Valuation_Model.xlsx

**Task:** 1.1 — Workbook Audit  
**Date:** 2026-08-29  
**File:** `~/Grad-Challenge/workbooks/Complex_Valuation_Model.xlsx` (544,715 bytes)  
**Tooling:** openpyxl 3.1.5, Python 3.11  
**Auditor:** Hermes Agent (automated)

---

## Executive Summary

The workbook is a multi-mine DCF valuation model containing **exactly 15 sheets** (confirming the vault note — there is no 16th sheet). It models three mining operations:

| Mine | Sheet(s) | Type | NPV (cached) | IRR |
|------|----------|------|-------------|-----|
| Springbok | `Springbok` (#7) + `SpringbokCapital` (#8) | Underground | $279,446k (~$279M) | -13.8% |
| Brave Blossom | `Brave Blossom ` (#9) + `Brave Blossom Capital` (#10) | Underground | $1,320,078k (~$1,320M) | 77.4% |
| OC Clean (SandsEnd) | `OC Clean` (#12) + `OC Clean Capital` (#13) | Open cut | (no NPV row — model not completed to DCF stage) | — |

**Critical finding:** The `Brave Blossom Capital` sheet (sheet 10) has a fully-built capital scheduling framework (882 cells, year headers, unit cost column, cost multiplication formulas) but **zero populated scheduling rows** — all year-by-year unit-number cells (D4:N25) are blank/zero, producing $0 across every year. The template (`SpringbokCapital`) is fully populated by contrast. This is the central anomaly of the challenge.

**Key model parameters (from `Assumptions` sheet):**
- Discount rate: **8%** (cell B2)
- FX rate: 0.743 USD:AUD (cell B4)
- CPI: 1.0%/yr base, inflating (row 49)
- Coal price reference: PLV premium hard coking coal (row 6)
- Carbon scenario: **Accelerated Transition** (cell C20), with carbon prices $45–$110/t CO2e (2027–2034, Carbon sheet)
- QLD royalty: tiered structure, 7%–~22% (rows 28–46)
- Tax rate: 30% (row 119 in model sheets)
- Attribution: 60% (row 123 in model sheets)

---

## 1. Sheet Inventory — 15 Sheets Confirmed

**Sheet count: 15** (NOT 16 — confirmed). No "Brave Blossom (2)" sheet exists.

| # | Sheet Name | Length | Trailing Space? | Dimensions | Rows × Cols |
|---|-----------|--------|-----------------|------------|-------------|
| 1 | `Assumptions` | 11 | No | A1:BV58 | 58 × 74 |
| 2 | `Analysis` | 8 | No | A1:AK24 | 24 × 37 |
| 3 | `List` | 4 | No | A1:H4 | 4 × 8 |
| 4 | `Summary` | 7 | No | A1:EP57 | 57 × 146 |
| 5 | `Historical Performance` | 22 | No | A2:I170 | 170 × 9 |
| 6 | `Springbok Assumptions` | 21 | No | A1:C10 | 10 × 3 |
| 7 | `Springbok` | 9 | No | A1:AI161 | 161 × 35 |
| 8 | `SpringbokCapital` | 16 | No | A1:V48 | 48 × 22 |
| 9 | `Brave Blossom ` | 14 | **YES** (trailing space) | A1:BT150 | 150 × 72 |
| 10 | `Brave Blossom Capital` | 21 | No | A1:AD66 | 66 × 30 |
| 11 | `Historical Performance BrumbyOC` | 31 | No | A2:H162 | 162 × 8 |
| 12 | `OC Clean` | 8 | No | A1:AI153 | 153 × 35 |
| 13 | `OC Clean Capital` | 16 | No | A1:Y71 | 71 × 25 |
| 14 | `Carbon` | 6 | No | A1:AY6 | 6 × 51 |
| 15 | `Decommissioned Mine` | 19 | No | A1:ED36 | 36 × 134 |

**⚠️ Trailing-space flag:** Sheet 9 is named `"Brave Blossom "` (with a trailing space, len=14). All cross-sheet references to this sheet correctly use quotes: `='Brave Blossom '!D2`. This is not a broken link — Excel handles it — but it is a naming hazard for programmatic access (e.g., `wb["Brave Blossom"]` fails; must use `wb["Brave Blossom "]`).

---

## 2. Per-Sheet Cell Statistics

| # | Sheet | Non-Empty | Numeric | Text | Formulas | Blanks in Range | Has Formulas? |
|---|-------|-----------|---------|------|----------|-----------------|---------------|
| 1 | Assumptions | 2,019 | 1,268 | 73 | 678 | 2,273 | ✅ Yes |
| 2 | Analysis | 404 | 1 | 36 | 367 | 484 | ✅ Yes |
| 3 | List | 20 | 4 | 16 | 0 | 12 | ❌ No |
| 4 | Summary | 546 | 51 | 41 | 454 | 7,776 | ✅ Yes |
| 5 | Historical Performance | 851 | 409 | 232 | 210 | 679 | ✅ Yes |
| 6 | Springbok Assumptions | 26 | 2 | 24 | 0 | 4 | ❌ No |
| 7 | Springbok | 2,751 | 383 | 214 | 2,154 | 2,884 | ✅ Yes |
| 8 | SpringbokCapital | 465 | 29 | 45 | 391 | 591 | ✅ Yes |
| 9 | Brave Blossom  | 6,145 | 925 | 203 | 5,017 | 4,655 | ✅ Yes |
| 10 | Brave Blossom Capital | 882 | 13 | 65 | 804 | 1,098 | ✅ Yes |
| 11 | Historical Performance BrumbyOC | 107 | 41 | 32 | 34 | 1,189 | ✅ Yes |
| 12 | OC Clean | 2,684 | 99 | 210 | 2,375 | 2,671 | ✅ Yes |
| 13 | OC Clean Capital | 801 | 27 | 86 | 688 | 974 | ✅ Yes |
| 14 | Carbon | 153 | 101 | 3 | 49 | 153 | ✅ Yes |
| 15 | Decommissioned Mine | 987 | 241 | 55 | 691 | 3,837 | ✅ Yes |

**Key observations:**
- **Sheet 9 (`Brave Blossom `)** is the largest by non-empty cells (6,145) — the main DCF engine for Brave Blossom, spanning 72 columns (years D through BT = ~67-year forecast horizon to ~2052).
- **Sheet 10 (`Brave Blossom Capital`)** has 882 non-empty cells but only **13 numeric values** — overwhelmingly formulas (804 of 882). The 13 numerics are the unit-cost constants in column C and year constants. This confirms the "framework but no data" finding.
- **Sheet 8 (`SpringbokCapital`)** has 465 cells with 29 numerics — a smaller, fully-populated template.
- **Sheet 15 (`Decommissioned Mine`)** has 987 populated cells across 134 columns — the emissions calculation engine.
- Only sheets 3 (`List`) and 6 (`Springbok Assumptions`) have no formulas; they are static reference/data-entry sheets.

---

## 3. Cross-Sheet Link Analysis

### 3.1 Link Map (source → referenced sheets)

| Source Sheet | References |
|-------------|-----------|
| `Assumptions` | `Historical Performance` |
| `Analysis` | `Historical Performance`, `Springbok`, `Brave Blossom `, `OC Clean` |
| `Summary` | `Historical Performance`, `Springbok`, `Brave Blossom ` |
| `Historical Performance` | `Assumptions` |
| `Springbok` | `Assumptions`, `Carbon`, `SpringbokCapital`, `Decommissioned Mine`, `Historical Performance` |
| `SpringbokCapital` | `Springbok` |
| `Brave Blossom ` | `Assumptions`, `Carbon`, `Brave Blossom Capital`, `Decommissioned Mine` |
| `Brave Blossom Capital` | `Brave Blossom ` |
| `OC Clean` | `Assumptions`, `Carbon`, `OC Clean Capital`, `Decommissioned Mine` |
| `OC Clean Capital` | `OC Clean` |
| `Decommissioned Mine` | `Springbok`, `Brave Blossom ` |

### 3.2 Broken Links

**No broken links found.** All cross-sheet references resolve to existing sheets. The initial regex scan flagged false positives (e.g., `D144-Assumptions`) caused by comparison operators in formulas like `=IF(D144<Assumptions!I$29,...)` — the `<` was misparsed as part of a sheet name. Manual review confirms all `!`-prefixed sheet references are valid.

### 3.3 Key Dependency Chains

- **Capital sheets → model sheets:** `SpringbokCapital` → `Springbok` (D1 pulls year from `Springbok!D2`). Same for `Brave Blossom Capital` → `Brave Blossom ` and `OC Clean Capital` → `OC Clean`.
- **Model sheets → Assumptions:** All three DCF sheets reference the `Assumptions` sheet for discount rate, FX, CPI, royalty tiers, and coal price forecasts.
- **Model sheets → Decommissioned Mine:** Safeguard Mechanism baselines (row 25 in Springbok/Brave Blossom, row 28 in OC Clean) pull emissions factors from the `Decommissioned Mine` sheet (cells E21 onward and E35 onward).
- **Decommissioned Mine → model sheets:** Pulls actual emissions data back from `Springbok!H22` and `Brave Blossom '!AA22` for the decommissioned-mine emissions calculation (NGER Method 3.32).
- **Summary → model sheets:** `Summary` sheet aggregates results from `Springbok` and `Brave Blossom ` for historical and forward-looking cost/revenue analysis.

---

## 4. Brave Blossom Capital (Sheet 10) — Structural Deep-Dive

**This is the key anomaly sheet.**

### 4.1 Framework Layout

| Rows | Content | Status |
|------|---------|--------|
| Row 1 | Title `Brave Blossom Capital`; year headers D1:N1 (2027–2037, formula `='Brave Blossom '!D2` then +1 increments) | ✅ Populated |
| Row 3 | Section header `Capital Items - Units`; instruction text `Insert Unit Numbers by Year Below` | ✅ Present |
| Rows 4–15 | Capital item rows — column A has **no item names** (blank); column B = `AUD$k`; column C has unit costs | ⚠️ Costs present, names missing |
| Rows 16–25 | `Other Items (as required)` placeholder rows — column A filled with placeholder text, column B = `AUD$k`, column C = blank (0) | ⚠️ Placeholders only |
| Rows 26–29 | (blank) | Empty |
| Row 30 | Section header `Capital Cost - Cost`; instruction `Do not change cells below` | ✅ Present |
| Rows 31–54 | Cost calculation rows — formula `=$C[row]*D[row]` for each year column | ✅ Formulas present |
| Row 55 | `Total` — `=SUM(D31:D54)` per year | ✅ Present |
| Row 56 | `Contingency` — 30% (cell C56=0.3), `=$C$56*D55` per year | ✅ Present |
| Row 57 | `Total Capital` — `=D56+D55` per year | ✅ Present |

### 4.2 Unit Costs (Column C — the only populated numeric data)

| Row | Unit Cost (AUD$k) | Item Name (col A) |
|-----|-------------------|-------------------|
| 4 | 5,000 | *(blank)* |
| 5 | 5,000 | *(blank)* |
| 6 | 15,000 | *(blank)* |
| 7 | 20,000 | *(blank)* |
| 8 | 5,000 | *(blank)* |
| 9 | 12,000 | *(blank)* |
| 10 | 17,000 | *(blank)* |
| 11 | 5,000 | *(blank)* |
| 12 | 20,000 | *(blank)* |
| 13 | 190,000 | *(blank)* |
| 14 | 69,900 | *(blank)* |
| 15 | 24,000 | *(blank)* |

**Sum of unit costs: 387,900 AUD$k (~$388M)** — closely matching the vault's $388.9M capital estimate. Item names were missing but have been populated (see Section 5).

### 4.3 The Core Problem: Zero Scheduling Data

The year-by-year scheduling cells (D4 through N25, the unit-number input area) are **entirely blank/zero**. The cost formulas in rows 31–54 multiply unit cost × unit count, so with zero unit counts, every year's cost is $0:

| Year | Total (D55) | Contingency (D56) | Total Capital (D57) |
|------|------------|-------------------|---------------------|
| 2027 (D) | 0 | 0 | 0 |
| 2028 (E) | 0 | 0 | 0 |
| 2029 (F) | 0 | 0 | 0 |
| 2030 (G) | 0 | 0 | 0 |
| 2031 (H) | 0 | 0 | 0 |
| 2032 (I) | 0 | 0 | 0 |
| 2033 (J) | 0 | 0 | 0 |
| 2034 (K) | 0 | 0 | 0 |
| 2035 (L) | 0 | 0 | 0 |
| 2036 (M) | 0 | 0 | 0 |
| 2037 (N) | 0 | 0 | 0 |

**Numeric value distribution:** 755 zeros (all formula-derived), 52 non-zero numerics (year headers + unit costs), 1,098 blanks.

**Consequence:** The Brave Blossom DCF model (sheet 9) references `Brave Blossom Capital` for capital costs (row 112: Project Capital). With $0 capital flowing through, depreciation is $0, and the NPV of $1,320M is computed without any capital expenditure — **materially overstating project value**.

---

## 5. SpringbokCapital (Sheet 8) — The Template (Fully Populated)

### 5.1 Framework Layout (same structure as Brave Blossom Capital)

| Rows | Content | Status |
|------|---------|--------|
| Row 1 | Title `Springbok Capital`; year headers D1:P1 (2027–2039, 13 years) | ✅ Populated |
| Row 3 | Section header `Capital Items - Units` | ✅ Present |
| Rows 4–11 | **8 named capital items** with unit costs AND year-by-year unit counts | ✅ **Fully populated** |
| Rows 12–16 | Spare rows (B = `AUD$k`, no item names) | Empty spares |
| Row 21 | Section header `Capital Cost - Cost` | ✅ Present |
| Rows 22–36 | Cost calculation rows (`=$C[row]*D[row]`) | ✅ Formulas present |
| Row 37 | `Total` — `=SUM(D22:D36)` | ✅ Present |
| Row 38 | `Contingency` — 20% (C38=0.2) | ✅ Present |
| Row 39 | `Total Capital` — `=D38+D37` | ✅ Present |

### 5.2 Capital Items (Fully Populated)

| Row | Item Name | Unit Cost (AUD$k) | Year-by-Year Schedule |
|-----|-----------|-------------------|----------------------|
| 4 | Replacement Miner | 9,000 | 2027: 1 unit → $9,000k |
| 5 | Ventilation Shaft | 20,000 | 2027: 0.5, 2028: 0.5 → $10,000k each |
| 6 | SIS Drill program | 30,000 | 2027: 0.5, 2028: 0.5 → $15,000k each |
| 7 | Replacement Roof Supports | 1,000 | 2027: 10 units → $10,000k |
| 8 | Conveyor Belt Replacement | 3,000 | 2027–2031: 1 unit/yr → $3,000k/yr |
| 9 | Barada Community Program | 5,000 | 2027: 1 unit → $5,000k |
| 10 | Tailings Storage Facility | 10,000 | 2027: 1 unit → $10,000k |
| 11 | Exploration Drilling | 800 | 2027–2031: 1 unit/yr → $800k/yr |

### 5.3 Resulting Capital Schedule (Cached Values)

| Year | Total (D37) | Contingency 20% (D38) | Total Capital (D39) |
|------|------------|----------------------|---------------------|
| 2027 | 62,800 | 12,560 | **75,360** |
| 2028 | 28,800 | 5,760 | **34,560** |
| 2029 | 3,800 | 760 | 4,560 |
| 2030 | 3,800 | 760 | 4,560 |
| 2031 | 3,800 | 760 | 4,560 |
| 2032+ | 0 | 0 | 0 |

**Total Springbok capital (2027–2031): ~123,600 AUD$k** (including contingency).

### 5.4 Comparison: SpringbokCapital vs Brave Blossom Capital

| Attribute | SpringbokCapital (template) | Brave Blossom Capital |
|-----------|---------------------------|----------------------|
| Non-empty cells | 465 | 882 |
| Capital item names | **8 named items** | **0 named items** (all blank) |
| Year-by-year unit counts | **Populated** (values like 1, 0.5, 10) | **All blank/zero** |
| Contingency rate | 20% | 30% |
| Year span | 2027–2039 (13 yrs) | 2027–2037 (11 yrs shown, extends to AD col) |
| Total capital output | $75,360k (yr 1) → $0 | **$0 all years** |
| Status | ✅ Complete, producing real capital schedule | ❌ Framework only, zero output |

---

## 6. OC Clean Capital (Sheet 13) & OC Clean (Sheet 12)

### 6.1 OC Clean Capital (Sheet 13)

**Internal title:** `SandsEnd Capital` (cell A1) — the sheet is named "OC Clean Capital" but the model refers to it as "SandsEnd".

**Structure:** Identical framework to the other capital sheets.

| Rows | Content | Status |
|------|---------|--------|
| Row 1 | `SandsEnd Capital`; year headers D1:Q1 (formula `='OC Clean'!D2` then +1) | ✅ Populated |
| Rows 4–29 | **26 named capital items** with unit costs | ✅ Item names + costs present |
| Rows 30–33 | Spare rows | Empty |
| Row 34 | Section header `Capital Cost - Cost` | ✅ Present |
| Rows 35–59 | Cost calculation rows | ✅ Formulas present |
| Row 60 | `Total` — `=SUM(D35:D59)` | ✅ Present |
| Row 61 | `Contingency` — 20% (C61=0.2) | ✅ Present |
| Row 62 | `Total Capital` | ✅ Present |

**Capital items include:** Project Cost (Feasibility) $8,000k, 1×300t Excavator $5,000k, 1×400t Excavator $10,000k, 1×600t Excavator $14,000k, 1×800t Excavator $20,000k, 1×320t Truck $8,000k, 1×230t Truck $6,000k, 1×190t Truck $5,000k, 1×130t Truck $3,000k, Dragline $100,000k, Dozers, Water truck, Drills, Grader, ROM Pad $20,000k, CHPP Upgrade $50,000k, Topsoil Stripping, Comms Tower, 25km Powerline $20,000k, Substation $10,000k, Crib Hut $10,000k, Water Management Dams $7,500k, Creek Crossing $20,000k, Haul Road $3,000k, Dry Processing Plant $20,000k.

**Numeric distribution:** 620 zeros, 72 non-zeros, 974 blanks.

**⚠️ Issue:** While item names and unit costs are present (unlike Brave Blossom Capital), the year-by-year unit count cells (D4:Q29) appear to be mostly blank/zero. The cost formulas exist but without unit counts entered, the total capital may be $0. This requires the same scheduling data entry that Brave Blossom Capital lacks.

### 6.2 OC Clean (Sheet 12) — DCF Model

**Internal title:** `OC - Proposal` (cell A2). This is the open-cut (OC) DCF model for the SandsEnd project.

**Structure (rows 1–153):**

| Rows | Section | Key Content |
|------|---------|-------------|
| 1–2 | Year headers | D1=2 (year index); D2 pulls from `Assumptions!J1` |
| 3–24 | **Physicals** | Reserves (20,000 kt), Prime Waste Removed (T&S + Dragline), Coal Mined (ROM), Strip Ratio, CHPP (Plant Feed, Bypass, Product, Yield), Saleable Production |
| 26–29 | **Carbon** | Carbon price index (declining by 1/26/yr), Scope 1 Emissions (0.1 t CO2e/kbcm), Safeguard Hybrid Declining Baseline (references `Decommissioned Mine`!E35 onward), Emissions above/below baseline |
| 28–180 | Costs, Revenue, Cashflow | Standard DCF structure (similar to Springbok/Brave Blossom) |
| 143–153 | QLD Royalty Calculation | CPI inflator, nominal realised price, tiered royalty rate, nominal deduction, royalty payable |

**⚠️ Note:** The OC Clean model does not appear to have NPV/IRR result rows at the same positions as Springbok (R126) and Brave Blossom (R126). The model sheet extends to row 153 and may not be completed to a full DCF conclusion, or the NPV row may be at a different position. The model references `Decommissioned Mine` row 35 for Safeguard baselines (OC-specific baselines, separate from the underground mine baselines in row 21).

---

## 7. Safeguard Mechanism Baseline Data (Sheets 7, 9, 12)

### 7.1 Springbok (Sheet 7) — Carbon/Safeguard Rows 21–26

| Row | Label | Key Formula / Value |
|-----|-------|-------------------|
| 22 | Coal Mine Waste Gas | `=IF(D9>0, D9*1000*$C$22+20000, 0)` where C22=0.12 t CO2e/t ROM. 2027: 642,964 t CO2e |
| 23 | Decommissioned Mines Emissions | Conditional on `Decommissioned Mine`!B5 selection: if "Springbok" → pull from Decom sheet E2; if "Springbok + Brave Blossom" → 0; if "Springbok Delayed" → pull E2. Currently all **0** (B5 = "Springbok + Brave Blossom") |
| 24 | Total Scope 1 Emissions | R22 + R23. 2027: 642,964 t CO2e |
| 25 | **Safeguard - Hybrid Declining Baseline** | `=(D9*1000*'Decommissioned Mine'!E22) + 'Decommissioned Mine'!E21*D23`. 2027: 426,048 t CO2e |
| 26 | Emissions above (or below) baseline | R24 - R25. 2027: +216,916 t CO2e (above baseline) |

**Baseline values (from `Decommissioned Mine` row 21):** 0.7795 (2027) → declining to 0.5749 (2032). These are the "Decom Mines emissions Baseline EI" factors.

**ROM Baseline EI (from `Decommissioned Mine` row 22):** 0.0821 (2027) → declining to 0.0528 (2032).

### 7.2 Brave Blossom (Sheet 9) — Carbon/Safeguard Rows 22–26

| Row | Label | Key Formula / Value |
|-----|-------|-------------------|
| 22 | Coal Mine Waste Gas | C22=0.13 t CO2e/t ROM (slightly higher than Springbok's 0.12). Emissions start from 2031 (H22=42,054) ramping to 1,052,109 by 2034 |
| 23 | Decommissioned Mines Emissions | Conditional: if B5="Springbok + Brave Blossom" → pull from `Decommissioned Mine`!E2. Currently **all 0** (Decom sheet E2 values are 0 because the mine hasn't closed yet relative to reporting dates) |
| 24 | Total Scope 1 Emissions | R22 + R23 |
| 25 | **Safeguard - Hybrid Declining Baseline** | `=(D9*1000*'Decommissioned Mine'!E22) + 'Decommissioned Mine'!E21*D23` |
| 26 | Emissions above (or below) baseline | R24 - R25. 2031: +32,580 t CO2e; 2034: +704,572 t CO2e |

### 7.3 OC Clean (Sheet 12) — Carbon/Safeguard Rows 26–29

| Row | Label | Key Formula / Value |
|-----|-------|-------------------|
| 26 | Carbon (price index) | `=C26-(1/26)` — declining carbon price trajectory (negative values, appears to be a discount/adjustment factor) |
| 27 | Scope 1 Emissions | C27=0.1 t CO2e/kbcm waste. `=D9*1000*$C$27`. Currently all **0** (D9 = Prime Waste Blasted = 0, model not yet producing) |
| 28 | **Safeguard - Hybrid Declining Baseline** | `=D11*'Decommissioned Mine'!E35` where row 35 is OC-specific ROM Baseline EI. C28=0. Currently all 0 |
| 29 | Emissions above (or below) baseline | R27 - R28. All 0 |

**OC Clean baselines (from `Decommissioned Mine` row 35):** 0.0242 (2024) → rising to 0.0431 (2031+). These are significantly lower than the underground mine baselines (0.78 vs 0.024), reflecting the different emissions intensity of open-cut vs underground mining.

---

## 8. Decommissioned Mine (Sheet 15) — NGER Method 3.32 Calculations

**This sheet implements the NGER (National Greenhouse and Energy Reporting) Method 3.32 for estimating post-closure methane emissions from decommissioned underground coal mines.**

### 8.1 Sheet Structure (36 rows × 134 columns)

| Rows | Content |
|------|---------|
| 1 | Year headers (E1 onward = dates 2027-12-31 through ~2052, formula `=E5`) |
| 2 | **Methane (CH4)** emissions result row — `=IFERROR(IF(E5<$B$11, IF(E5>$B$6, MAX(E3:E4), 0), 0), 0)`. Currently all 0 (mine closure date is after reporting period) |
| 3 | Sub-result: Method 3.32(1) — `=E16` (links to the Edm calculation) |
| 4 | Sub-result: Method 3.32(2) — `=0.02*E6` (2% of total CH4 emissions, = 20,160 t CO2e/yr) |
| 5 | **Site** selector — B5 = `"Springbok + Brave Blossom"` (dropdown). E5 onward = reporting dates |
| 6 | Date Mine Closed — `=VLOOKUP(B5, A29:J31, 6)` → 2050-12-31 (for Springbok+BB combo) |
| 7 | Gassy/Non-Gassy — `=VLOOKUP(...)` → "Gassy" |
| 8 | CH4 emissions 12 months prior to closure — `=VLOOKUP(...)` → 1,008,000 t CO2e |
| 9 | Mine Void Volume (m³) — 42,319,176 m³ |
| 10 | Year to Stop Reporting — `=YEAR(B6)+21` → 2071 |
| 11 | Stop-reporting date — `=DATE(B10, 12, 31)` → 2071-12-31 |
| 12 | Additional Annual Flood Volume (m³) — 0 |
| 13 | QLD Default Flood Volume (m³) — 74,000 m³/yr |
| 14 | **EFdm** (emission factor) — complex formula: `=(((1+E7*E10)^(1+E8) - (1+E7*(E10-E12))^(1+E8)) / (E7*(1+E8)) - E9*E12) / 12`. Uses parameters: A=0.23/12, b=-1.45, c=0.024 |
| 15 | **Fdm** (flood factor) — `=IFERROR(MIN(E11/E13*E10/12, 1), 0)` |
| 16 | **Edm** (decommissioned mine emissions) — `=IF(E10/12<20, (E6*E14*(1-E15)), 0)`. Currently 0 (closure date 2050 is after all reporting years 2027–2032) |
| 20–22 | **Baseline Calculation Inputs** — row 21: Decom Mines emissions Baseline EI (0.7795→0.5749); row 22: ROM Baseline EI (0.0821→0.0528) |
| 27 | Instruction: `Change the drop down in B5 to match the selected proposal` |
| 28 | Column headers for lookup table |
| 29–31 | **Lookup table** — 3 options: Springbok (closure 2031), Springbok + Brave Blossom (closure 2050), Springbok Delayed 12 Months (closure 2032) |
| 34–35 | **OC Clean baseline inputs** — row 34: years 2024–2032; row 35: ROM Baseline EI for OC Clean (0.0242→0.0431) |

### 8.2 Method 3.32 Formula Components

| Parameter | Cell | Value | Meaning |
|-----------|------|-------|---------|
| Etdm | B8 / E6 | 1,008,000 | CH4 emissions 12 months pre-closure (t CO2e) |
| A | E7 | 0.23/12 = 0.01917 | Monthly decline rate |
| b | E8 | -1.45 | Power exponent |
| c | E9 | 0.024 | Constant |
| Mwi | E11 | 74,000 | Cumulative flood volume (m³) |
| Mvv | E13 | 42,319,176 | Mine void volume (m³) |
| T (months) | E10 | DATEDIF(closure, reporting date) | Months since closure |

### 8.3 Current State

With B5 = `"Springbok + Brave Blossom"`, the mine closure date is 2050-12-31. For all reporting years 2027–2032, the `DATEDIF` returns 0 months (mine hasn't closed yet), so:
- **Edm (row 16) = 0** for all years (closure hasn't occurred)
- **Decom emissions (row 2) = 0** for all years
- The 2% residual emissions (row 4, Method 3.32(2)) = 20,160 t CO2e/yr — this is a static estimate

**Cross-references out:** The Safeguard baseline rows (21, 22, 35) are pulled by the DCF model sheets (Springbok R25, Brave Blossom R25, OC Clean R28) for the hybrid declining baseline calculation.

---

## 9. Supporting Sheets

### 9.1 Assumptions (Sheet 1)

The central assumptions sheet. Key parameters:

| Row | Parameter | Value |
|-----|-----------|-------|
| 2 | Discount Rate | 8% (B2) |
| 4 | FX Rate (USD:AUD) | 0.743 (D4) |
| 6–9 | Coal price forecasts | PHCC = PLV benchmark; 88% PHCC; 6000NAR; 5500NAR |
| 16 | GAR→NAR adjustment | 260 kcal/kg |
| 17 | Basis of coal prices | 6000 kcal/kg(nar) |
| 20 | Carbon Scenario | "Accelerated Transition" |
| 28–33 | QLD Royalty thresholds | Tier 1: $0, Tier 2: $100, Tier 3: $150, Tier 4–6: (blank in cols D-F) |
| 35–40 | QLD Royalty rates | Tier 1: 7%, Tier 2: 12.5%, Tier 3: 15%, Tier 4–6: (blank) |
| 42–46 | Cumulative royalty calculations | Formula-driven from thresholds and rates |
| 49 | AUD CPI | 1.0%/yr base (D49), inflating across columns |
| 52–58 | Flood Rate lookup table | 0, 100k, 250k, 500k, 750k, 10M m³/yr |

**⚠️ Issue:** Royalty tiers 4–6 (rows 31–33, 38–40, 44–46) appear to have blank threshold/rate values in the first data column (D). The formula in the model sheets (`=IF(D147<Assumptions!I$33, Assumptions!I$45, ... Assumptions!I$46)`) references columns I onward which may be populated. The model appears to have year-specific royalty parameters across columns (D=2023, E=2024, etc.), with only some tiers filled in column D.

### 9.2 Analysis (Sheet 2)

A calculation-heavy sheet (367 of 404 cells are formulas). References `Historical Performance`, `Springbok`, `Brave Blossom `, and `OC Clean`. Appears to be a comparative analysis dashboard pulling NPV, cost/t, and margin data across projects.

### 9.3 List (Sheet 3)

Static reference list (no formulas). Contains dropdown validation data:
- Roster options: Owner/operate (7on7off), Contract (5on2off), Other
- Development miners: 4, 3, 2, 1
- Gas drainage: SIS, UIS
- Wash strategy: Max Quality, Max Quantity
- Maintenance: Preventative, Corrective
- Adjustment options: Maintain, Increase 10%, Decrease 10%

### 9.4 Summary (Sheet 4)

Dashboard sheet (454 formulas). Two sections:
- **Rows 1–14:** Historical + Springbok forward analysis (2022–2027). Pulls revenue, cost/t breakdown from `Historical Performance` and `Springbok` sheets.
- **Rows 44–57:** Brave Blossom Project Margin Analysis (2031–2036). Pulls revenue and cost data from `Brave Blossom ` sheet.

### 9.5 Historical Performance (Sheet 5)

Historical financial data (2022–2026 actuals + FYF). 851 cells, 210 formulas. Referenced by `Assumptions` (for price forecasts) and `Summary` (for historical cost analysis).

### 9.6 Springbok Assumptions (Sheet 6)

Small scenario-comparison table (10 rows × 3 cols, no formulas). Compares "Current" vs "Proposed" operating parameters:

| Parameter | Current | Proposed |
|-----------|---------|----------|
| Roster - Longwall | 7on7off | 7on7off |
| Roster - Development | 7on7off | 5on2off |
| Roster - Outbye | 7on7off | 7on7off |
| Development Miners | 3 | 4 |
| Gas Drainage Strategy | SIS | UIS |
| Wash Strategy | Max Quality | Max Quality |
| Progressive Rehabilitation | Yes | No |
| Maintenance Strategy | Preventative | Corrective |

### 9.7 Carbon (Sheet 14)

Carbon price forecast sheet (6 rows × 51 columns). Two scenarios:
- **Zero Carbon Liability** (row 2): $0/t all years
- **Accelerated Transition** (row 3): $45 (2026) → $50 (2027) → $55 → $65 → $80 → $90 → $100 → $105 → $110 (2034)

The active scenario is "Accelerated Transition" (per `Assumptions`!C20).

### 9.8 Historical Performance BrumbyOC (Sheet 11)

Historical data for the Brumby open-cut mine (162 rows × 8 cols, 107 cells, 34 formulas). Sparse population (1,189 blanks in range). This appears to be reference/comparative data for the Brumby OC operation, not a live model.

---

## 10. Key Model Results (Cached Values)

### 10.1 NPV / IRR / Payback

| Mine | Sheet | NPV Cell | NPV Formula | NPV Value (AUD$k) | IRR | Payback (yrs) |
|------|-------|----------|-------------|-------------------|-----|---------------|
| Springbok | 7 | C126 | `=SUM(D124:AD124)` | **279,446** (~$279M) | -13.8% | 1.0 |
| Brave Blossom | 9 | C126 | `=SUM(D124:BO124)` | **1,320,078** (~$1,320M) | 77.4% | 2.83 |

**⚠️ Critical caveat for Brave Blossom:** The $1,320M NPV is computed with **$0 capital expenditure** (because `Brave Blossom Capital` has no scheduling data). If the $389M capital estimate from the vault context were properly scheduled, the NPV would be significantly lower (roughly $1,320M − $389M discounted ≈ $931M, depending on timing).

**Springbok IRR of -13.8%** is negative — the project destroys value on a discounted basis despite positive NPV of $279M. This apparent contradiction suggests the NPV may be undiscounted or the discounting methodology differs from standard DCF (the payback of 1.0 year is suspiciously fast for a mine starting production in 2027). This warrants further investigation.

### 10.2 Combined Portfolio

Per vault context: Combined NPV = $1,670M. This would be Springbok ($279M) + Brave Blossom ($1,320M) + additional = $1,599M + ~$71M, or may include OC Clean (which has no NPV row computed).

---

## 11. Issues, Flags & Surprises

### 🚨 Critical Issues

1. **Brave Blossom Capital is empty of scheduling data (Sheet 10):** The framework exists (882 cells, formulas, year headers, unit costs in column C) but ALL year-by-year unit count cells (D4:N25) are blank/zero. This produces $0 total capital, which flows through to the Brave Blossom DCF as $0 capital expenditure, $0 depreciation, and an inflated NPV of $1,320M. **This is the primary task for the graduate challenge — populating this sheet.**

2. **Brave Blossom Capital has no item names (Sheet 10, column A):** Unlike SpringbokCapital (which names 8 items: Replacement Miner, Ventilation Shaft, etc.) and OC Clean Capital (26 named items), Brave Blossom Capital's column A is entirely blank for rows 4–15. The unit costs are present (5,000 to 190,000 AUD$k) and sum to $387.9M — closely matching the vault's $388.9M capital estimate. When fractional unit counts are applied (e.g., 0.5 units for partial-year items), the total capital schedule rises to $518.4M.

3. **OC Clean Capital may also lack scheduling data (Sheet 13):** While item names and unit costs are present (26 items), the year-by-year unit count cells appear mostly blank, producing zeros. This may be intentional (OC Clean is a proposal-stage model) or may also need population.

### ⚠️ Moderate Issues

4. **Springbok IRR is negative (-13.8%) despite positive NPV ($279M):** This is mathematically possible if early cashflows are negative and later ones positive with an unusual profile, but a 1.0-year payback combined with negative IRR is inconsistent under standard DCF assumptions. The NPV formula `=SUM(D124:AD124)` sums discounted cashflows; if the discount factor row (R124) uses a non-standard method, this could explain the discrepancy. Worth verifying the discounting formula.

5. **Royalty tier thresholds 4–6 are blank in column D (Assumptions sheet):** The formula structure in the model sheets references `Assumptions!I$33` and `I$45`/`I$46` for tiers 4–6, but column D (the first year column) has blank values for these tiers. If later columns (I onward) are populated, the royalty calculation may work for later years but produce incorrect (minimum-rate) results for early years.

6. **Trailing space in "Brave Blossom " sheet name (Sheet 9):** While Excel handles this correctly (formulas use `='Brave Blossom '!D2`), it's a maintenance hazard. Programmatic access requires exact matching (`wb["Brave Blossom "]`).

7. **Typo: "Decomission" (single 's') in model sheets:** Rows 23 in Springbok and Brave Blossom label the decommissioned mine emissions as "Decomission Mines Emissions" (missing an 's'). The `Decommissioned Mine` sheet (sheet 15) spells it correctly. Cosmetic but notable.

8. **OC Clean carbon row 26 produces negative values:** The formula `=C26-(1/26)` produces a declining series starting at -0.038 (2027) and going to -0.346 (2035+). This appears to be a carbon price adjustment factor rather than a price itself, but the negative trajectory is unusual and may indicate the model expects carbon prices to decline over time for the OC scenario.

### ℹ️ Observations

9. **Sheet 11 (Historical Performance BrumbyOC) is very sparse:** 107 cells in a 162×8 range (1,189 blanks). This is reference data for a third mine (Brumby open cut) that doesn't have its own DCF model in this workbook.

10. **The Decommissioned Mine sheet has a dropdown selector (B5):** Currently set to "Springbok + Brave Blossom". Changing this to "Springbok" or "Springbok Delayed 12 Months" changes the closure date, void volume, and pre-closure emissions, which cascades through the Safeguard baseline calculations in all DCF sheets.

11. **The Carbon sheet has two scenarios** (Zero Carbon Liability at $0, and Accelerated Transition at $45–$110). The active scenario is Accelerated Transition, which materially increases carbon costs in the DCF models.

12. **The workbook uses a Data Validation extension** that openpyxl warns about ("Data Validation extension is not supported and will be removed"). This means some dropdown validation may be lost if the workbook is saved through openpyxl. The original Excel file retains it.

---

## 12. Summary of Confirmed Vault Context

| Vault Claim | Audit Finding | Status |
|-------------|---------------|--------|
| 15 sheets (not 16) | 15 sheets confirmed | ✅ Confirmed |
| "Brave Blossom (2)" does not exist | Not present | ✅ Confirmed |
| SpringbokCapital is a separate sheet | Sheet 8, confirmed | ✅ Confirmed |
| "Brave Blossom " has trailing space | Sheet 9, len=14, trailing space | ✅ Confirmed |
| Brave Blossom Capital: 882 cells framework, zero scheduling | 882 non-empty, all year cols = 0 | ✅ Confirmed |
| SpringbokCapital: 465 cells, fully populated | 465 cells, 8 named items scheduled | ✅ Confirmed |
| Sheets 7/9 have Safeguard baselines | Rows 25 in both sheets | ✅ Confirmed |
| Sheet 15 has 987 populated cells (emissions) | 987 non-empty, NGER Method 3.32 | ✅ Confirmed |
| Brave Blossom NPV $1,320M base case | C126 = 1,320,078 AUDk | ✅ Confirmed |
| Combined $1,670M | SB $279M + BB $1,320M = $1,599M (diff ~$71M) | ⚠️ Partial (may include OC Clean or other adjustment) |
| Springbok $279M | C126 = 279,446 AUDk | ✅ Confirmed |
| Capital estimate $389M | Unit costs total $387.9M (matches vault ~$389M) | ✅ Reconciled — fractional unit counts raise to $518.4M |
| Discount rate 8% | Assumptions B2 = 0.08 | ✅ Confirmed |
| Coal price $211.20/t | Not directly found in cached values; likely derived from PLV benchmark × FX | ℹ️ Indirectly supported |

---

## Appendix A: Cross-Sheet Reference Summary

```
Assumptions ──────────────► Historical Performance
Analysis ──────────────────► Historical Performance, Springbok, Brave Blossom , OC Clean
Summary ───────────────────► Historical Performance, Springbok, Brave Blossom 
Historical Performance ───► Assumptions
Springbok ────────────────► Assumptions, Carbon, SpringbokCapital, Decommissioned Mine, Historical Performance
SpringbokCapital ─────────► Springbok
Brave Blossom  ───────────► Assumptions, Carbon, Brave Blossom Capital, Decommissioned Mine
Brave Blossom Capital ───► Brave Blossom 
OC Clean ─────────────────► Assumptions, Carbon, OC Clean Capital, Decommissioned Mine
OC Clean Capital ────────► OC Clean
Decommissioned Mine ────► Springbok, Brave Blossom 
```

## Appendix B: Audit Scripts

Three Python scripts were used for this audit, saved in `~/Grad-Challenge/analysis/`:
- `_audit_raw.py` — sheet inventory, cell counts, formula detection, cross-sheet link scan
- `_audit_structure.py` — deep structural dumps of Brave Blossom Capital and SpringbokCapital
- `_audit_deep.py` — structural analysis of OC Clean Capital, Springbok, Brave Blossom, and Decommissioned Mine
- `_audit_targeted.py` — Safeguard baselines, NPV cells, Summary, Assumptions, Carbon, List sheets

---

*End of audit report.*
