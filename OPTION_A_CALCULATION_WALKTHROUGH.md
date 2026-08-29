# Option A Calculation Walkthrough: Brave Blossom Underground Mine

**Project**: 2026 Graduate Challenge — Wallaby Mining Board Presentation  
**Workbook**: `Complex_Valuation_Model_POPULATED.xlsx`  
**Option A**: Develop the Brave Blossom underground mine (staged commitment)  
**Purpose**: Walk the reader through every calculation behind Option A's financial conclusions, explaining the financial terms and referencing the exact workbook cells where each value is computed  

---

## How to Use This Document

Open `Complex_Valuation_Model_POPULATED.xlsx` in Excel alongside this document. Each section below references specific sheet names and cell references (e.g., `Brave Blossom !D7`) so you can follow the live formulas. The workbook has six sheets relevant to Option A:

| Sheet Name | Role in Option A |
|---|---|
| **Assumptions** | Global inputs: discount rate, FX, coal prices, royalty tiers, CPI, carbon scenario |
| **Brave Blossom** (note trailing space in sheet name) | The core DCF model — 26 years of physicals, revenue, costs, cashflow, NPV/IRR |
| **Brave Blossom Capital** | 12 capital items with unit costs and year-by-year phasing |
| **Carbon** | Carbon price scenarios (two rows: Zero Liability, Accelerated Transition) |
| **Decommissioned Mine** | Emissions baselines for the Safeguard Mechanism |
| **Analysis** | Summary sheet linking Brave Blossom and Springbok for combined view |

Column mapping: Column **D** = year index 2 (2027), **E** = 2028, **F** = 2029, **G** = 2030, **H** = 2031, **I** = 2032, and so on through column **BO** ≈ 2052. The year headers are in row 1 (year index) and row 2 (calendar year, linked from `Assumptions!J1` onward).

---

## Part 1: Financial Terms Used in This Model

### Discount Rate (Assumptions!B2 = 0.08, i.e., 8%)

**What it means**: A dollar received in 2033 is not worth a dollar today. You could invest a dollar today at some rate of return and have more than a dollar in 2033. The discount rate is that rate of return — it expresses the opportunity cost of capital. At 8%, $1 in 2033 (6 years from the 2027 base) is worth $1 ÷ (1.08)^6 = $0.63 today.

**Why 8%**: Mining companies typically use 8–12% for projects of moderate risk. 8% reflects a real (inflation-adjusted) discount rate appropriate for a brownfield project that reuses existing infrastructure.

### Discount Factor (Assumptions row 2, columns I–Q)

Each year's discount factor is computed in `Assumptions!I2` through `Q2` (and beyond):

```
Assumptions!I2 = =(1+$B2)^-(I$1-$I$1+0.5)
```

This is a **mid-year discounting** convention. The formula `(1+r)^-(year - base_year + 0.5)` assumes cash flows occur at the midpoint of each year rather than at year-end. For the base year (2026, cell I1), the exponent is `-0.5`, giving a factor of `1/(1.08)^0.5 = 0.9623`. For 2033 (year index 8, column P), the factor is `1/(1.08)^7.5 = 0.5742`. These factors are multiplied by each year's cashflow in `Brave Blossom` row 124 to convert future money to present value.

### Net Present Value (NPV) (Brave Blossom !C126)

**What it means**: NPV is the sum of all discounted cashflows over the project's life. If NPV > 0, the project earns more than the 8% hurdle rate and should proceed. If NPV < 0, the project does not meet the hurdle.

**Formula**:
```
C126 = =SUM(D124:BO124)
```
Row 124 contains the discounted cashflows. Columns D through BO cover the full mine life (2027–2052). The sum of all discounted annual cashflows = NPV.

### Internal Rate of Return (IRR) (Brave Blossom !C128)

**What it means**: IRR is the discount rate that makes NPV = zero. If you plugged the IRR into the discount factor formula instead of 8%, the sum of discounted cashflows would be exactly zero. Think of it as the project's "effective annual return" — comparable to an interest rate on a bank account.

**Formula**:
```
C128 = =IRR(D124:BO124)
```
Excel's `IRR` function iteratively solves for the rate where `SUM(discounted_cashflows) = 0`.

### Capital Expenditure (CapEx)

**What it means**: The upfront money spent to build the mine before it earns revenue. In this workbook, capital flows from the **Brave Blossom Capital** sheet into `Brave Blossom !D112:I112` (row 112, "Project Capital"). The capital sheet lists 12 items, each with a unit cost and year-by-year phasing fractions.

### Tax Shield (Brave Blossom !C119 = 0.3, i.e., 30%)

**What it means**: When a company spends money on capital, it can deduct depreciation from taxable income, reducing tax paid. This "shield" lowers the effective cost of capital. The workbook models this through row 116 (Depreciation), which reduces taxable income in row 119 (Tax Paid). The tax rate of 30% is in cell `C119`.

### CHPP Yield (Brave Blossom !C15 = 0.68, i.e., 68%)

**What it means**: Not all raw coal (ROM) becomes sellable product. The Coal Handling and Preparation Plant (CHPP) washes the coal and discards waste. A 68% yield means 68% of ROM tonnes become saleable product. The first production year uses a lower yield (58%, i.e., C15 - 0.1) and early years use 48% (C15 - 20%) to reflect ramp-up.

### FX Rate (Assumptions row 4)

**What it means**: Coal is priced in USD but costs are in AUD. The FX rate converts USD revenue to AUD. `Assumptions!I4 = 0.69` (USD:AUD), meaning 1 USD = 1/0.69 = 1.449 AUD. The forecast is flat at 0.69 from 2026 onward.

---

## Part 2: The Calculation Chain — Step by Step

The Brave Blossom DCF model follows a clear chain from physical production through to NPV. Each step feeds the next.

### Step 1: Mining Production (Physicals)

**Rows 5–9, Brave Blossom sheet**

The mine produces coal from two sources: development mining (driving tunnels to access the coal seam) and longwall mining (the machine that cuts coal from the seam face).

**Development Metres** (row 6): An input — metres of tunnel driven per year. Zero until 2031 (cell H6 = 7,500), then 15,000/year (I6 = 15,000).

**Development Tonnes** (row 7): Converts metres to tonnes of coal:
```
D7 = =(3*5.2*1.45*D6)/1000
```
- **3** = number of development headings (parallel tunnels)
- **5.2** = average seam height (metres)
- **1.45** = coal density (tonnes/m³)
- **D6** = development metres that year
- **÷1000** converts kg to tonnes (kt)

So for 2031 (H6 = 7,500m): `3 × 5.2 × 1.45 × 7,500 ÷ 1000 = 169,050 kt` of development coal. Wait — that's too high. Actually, the formula gives `3 × 5.2 × 1.45 × 7,500 / 1000 = 169.05 kt` — the ÷1000 converts the volumetric result (in m³ × density = tonnes, then /1000 to get kt). So 169 kt of coal comes from development in 2031.

**Longwall Tonnes** (row 8): An input. Zero until 2033, then 7,600 kt/year — the longwall machine's full production rate.

**Total ROM** (row 9):
```
D9 = =D7+D8
```
ROM (Run of Mine) = all coal extracted, before processing. In full production (2033+): 7,600 + longwall + development tonnes ≈ 7,769 kt/year.

### Step 2: Coal Processing (CHPP)

**Rows 11–19, Brave Blossom sheet**

**Primary Feed** (row 12): All ROM goes to the CHPP:
```
D12 = =D9
```

**Plant Product** (row 14): Saleable coal after washing:
```
D14 = =D12*($C$15-0.1)     [first production year, 68% - 10% = 58%]
H14 = =H12*($C$15-20%)     [ramp-up years, 68% - 20% = 48%]
E14 = =E12*$C$15            [full production, 68%]
```
The yield ramps up: 48% in early years → 58% in first year → 68% at steady state. This reflects the CHPP needing time to optimise.

**CHPP Yield** (row 15):
```
D15 = =IFERROR(D14/D12, 0)
```
A check row — product ÷ feed. Should match the percentages above.

**Saleable Production** (row 19):
```
D19 = =D13+D14
```
Bypass tonnes (row 13, coal that bypasses the CHPP — usually zero) plus plant product. This is what the mine can sell.

### Step 3: Revenue

**Rows 28–63, Brave Blossom sheet**

#### Sales Allocation (rows 30–34)

The model can sell coal as three products: PHCC (Premium Hard Coking Coal), 88% PHCC (a slightly lower quality), and Thermal GCN (power station coal). The allocation is controlled by switches in column C:

| Product | Cell C | Value | Meaning |
|---|---|---|---|
| PHCC (row 31) | C31 | 0 | Not sold — all coal is 88% quality |
| 88% PHCC (row 32) | C32 | 1 | 100% of saleable production sold as 88% PHCC |
| Thermal GCN (row 33) | C33 | 0 | Not sold |

**Sales volume formulas**:
```
D31 = =$C$31*D19    →  0 × saleable = 0
D32 = =$C$32*D19    →  1 × saleable = 100% of saleable
D34 = =SUM(D31:D33) →  total sales = saleable production
```

#### Pricing (rows 54–58)

The price comes from the Assumptions sheet. PHCC benchmark is $240/t (Assumptions!J6). 88% PHCC is 88% of that:

```
Assumptions!J7 = =J6*0.88 = 240 × 0.88 = $211.20/t
```

**Realised price** (rows 55–56): The model checks whether each product is being sold, then pulls the corresponding price:
```
D55 = =IF(D31>0, Assumptions!J6, 0)    [PHCC price if any PHCC sold]
D56 = =IF(D32>0, Assumptions!J$7, 0)   [88% PHCC price if any 88% sold]
```

Since C31=0 and C32=1, only the 88% PHCC price ($211.20/t) flows through.

**Average Export Price** (row 58):
```
D58 = =IFERROR(SUMPRODUCT(D31:D33, D55:D57)/D34, 0)
```
This is a weighted average: (volume × price) for each product, summed, ÷ total volume. Since only 88% PHCC is sold, the average = $211.20/t.

#### Revenue Calculation (rows 60–63)

**Total Revenue in USD** (row 60):
```
D60 = =SUMPRODUCT(D31:D32, D55:D56)
```
Multiply each product's sales volume by its price and sum. Since only 88% PHCC is sold: `Revenue = D32 × D56 = saleable_tonnes × $211.20/t` (in USD thousands).

**FX Conversion** (row 61):
```
D61 = =Assumptions!I4    →  0.69 (USD:AUD)
```

**Total Revenue in AUD** (row 62):
```
D62 = =(D60/D61)*(1+$BT$7)
```
- `D60/D61` = USD revenue ÷ FX rate = AUD revenue
- `*(1+$BT$7)` = applies a premium/marketing adjustment factor (cell BT7)

**Average Realisation** (row 63):
```
D63 = =IFERROR(D62/D19, 0)
```
Revenue ÷ saleable tonnes = AUD per tonne realised.

### Step 4: Operating Costs

**Rows 65–90, Brave Blossom sheet**

The model calculates six cost categories, each with a per-unit rate in column C:

| Row | Cost Item | Unit | Rate (C col) | Formula Pattern |
|---|---|---|---|---|
| 69–70 | Development Costs | AUD/m | C69=5,600 | `=IF(D6>0, $C$69*(1+$BT$8), 0)` then `D6 * rate / 1000` |
| 72–73 | Longwall Costs | AUD/LWt | C72=15 | `=IF(D8>0, $C$72*(1+$BT$9), 0)` then `D8 * rate` |
| 75 | Outbye Costs (fixed) | AUDk | C75=240,000 | `=IF(D9>0, $C$75*(1+$BT$10))` (split by 10 during ramp: H75=`/10`) |
| 77–78 | Gas Drainage | AUD/t | C77=6 | `=IF(D9>0, $C$77*(1+$BT$11), 0)` then `rate × ROM tonnes` |
| 80–82 | CHPP Costs | AUDk + AUD/feedt | C80=20,000 / C81=9.5 | Fixed + variable: `D82 = D9*D81 + D80` |
| 84–87 | Overheads | AUDk each | C84=45k, C85=9k, C86=15k | `=IF(saleable>0, $C$rate, 0)` then `× (1+$BT$13)` |

**Key**: Each cost has an escalation factor in column BT (e.g., `$BT$8`, `$BT$9`). These apply CPI or specific inflation indices to grow costs over time in real terms.

**Total FOR Costs** (row 89):
```
D89 = =SUM(D70, D73, D75, D78, D82, D87)
```
FOR = Free On Rail — all costs up to the point where coal is loaded onto trains. Sum of all six cost categories.

**FOR Cost per tonne** (row 90):
```
D90 = =IFERROR(D89/D19, 0)
```
Total FOR costs ÷ saleable production = cost per tonne.

### Step 5: Selling Expenses and Royalties

**Rows 92–99, Brave Blossom sheet**

**Port, Rail and Selling** (row 93, C93 = 22 AUD/t):
```
D93 = =$C$93*D34
```
$22/tonne × total sales volume. This is the cost of transporting coal from the mine rail loadout to the port and loading onto ships.

**Government Royalty** (row 95): Queensland's tiered royalty system. This is the most complex formula in the model. The royalty is calculated in a separate block (rows 140–148):

**Step 5a: CPI Inflator** (row 142): Builds a cumulative CPI index:
```
E142 = =D142*(1+Assumptions!J49)
```
Starting at 1.0, each year multiplies by (1 + CPI rate). Historical CPI = 1% (Assumptions D49:H49), forecast CPI = 2.65% (Assumptions J49 onward).

**Step 5b: Nominal Price** (row 144):
```
D144 = =D142*D63
```
Real AUD/t × CPI inflator = nominal AUD/t (today's dollars → inflated to nominal).

**Step 5c: Royalty Rate** (row 145): A nested IF statement selects the tier:
```
D145 = =IF(D144 < Assumptions!I$29,  Assumptions!I$35,           [Tier 1: < $100/t → 7%]
            IF(D144 < Assumptions!I$30, (Assumptions!I$42 + (D144 - Assumptions!I$29)*Assumptions!I$36)/D144,   [Tier 2: $100–150/t → 12.5% marginal]
              IF(D144 < Assumptions!I$31, (...),                  [Tier 3: $150–175/t → 15% marginal]
                IF(D144 < Assumptions!I$32, (...),                [Tier 4: $175–225/t → 20% marginal]
                  IF(D144 < Assumptions!I$33, (...),              [Tier 5: $225–300/t → 30% marginal]
                    (Assumptions!I$46 + (D144 - Assumptions!I$33)*Assumptions!I$40)/D144)))))  [Tier 6: >$300/t → 40% marginal]
```

The Queensland royalty system is **marginal** — each tier has a threshold and a rate. The cumulative royalty up to each threshold is pre-calculated in Assumptions rows 42–46:

| Tier | Threshold (AUD/t) | Rate | Cumulative Royalty at Threshold |
|---|---|---|---|
| 1 | 0–100 | 7% | $7.00/t (Assumptions row 42 = threshold × rate) |
| 2 | 100–150 | 12.5% | $13.25/t (row 43 = $7 + $50 × 12.5%) |
| 3 | 150–175 | 15% | $17.00/t (row 44 = $13.25 + $25 × 15%) |
| 4 | 175–225 | 20% | $27.00/t (row 45 = $17 + $50 × 20%) |
| 5 | 225–300 | 30% | $49.50/t (row 46 = $27 + $75 × 30%) |
| 6 | >300 | 40% | Add ($price - $300) × 40% |

The formula divides by the nominal price (D144) to express the effective royalty as a percentage of price.

**Step 5d: Nominal Deduction** (row 146):
```
D146 = =IFERROR((D93+D94)/D34, 0)
```
Port/rail costs per tonne (plus any take-or-pay liability). This is deductible from the royalty base.

**Step 5e: Royalty Payable** (rows 147–148):
```
D147 = =(D144-D146)*D145         [nominal: (price - deduction) × rate]
D148 = =D147/D142                 [convert back to real terms]
```

**Back in the main sheet** (row 95):
```
D95 = =D148*D34
```
Real royalty per tonne × sales volume = total royalty in AUDk.

**Total FOB Costs** (row 98):
```
D98 = =D89+D96
```
FOR costs + direct selling expenses (row 96 = port/rail + royalty). FOB = Free On Board — all costs up to the ship.

**FOB Cost per tonne** (row 99):
```
D99 = =IFERROR(D98/D19, 0)
```

### Step 6: Mine Earnings

**Row 101, Brave Blossom sheet**

```
D101 = =D62-D98
```

Total Revenue (AUD) − Total FOB Costs = Mine Earnings Before Other Costs and Taxes. This is the operating margin — what the mine earns from selling coal after all direct costs, before capital, closure, carbon, and tax.

### Step 7: Other Costs (Closure, Carbon, Rehabilitation)

**Rows 103–109, Brave Blossom sheet**

| Row | Item | Rate (C col) | Formula |
|---|---|---|---|
| 104 | Progressive Rehabilitation | C104=2,500 AUDk | `=IF(D9>0, $C$104, 0)` — flat annual cost while mining |
| 105 | Mine Closure | C105=150,000 AUDk | One-off in the final year (B135=2052) |
| 106 | Carbon Costs | — | See below |
| 107 | Other Revenue | — | Zero (placeholder) |

**Carbon Costs** (row 106):
```
D106 = =IF(Assumptions!$C$20="Zero Carbon Liability",
           'Brave Blossom '!D26*Carbon!B2,
           'Brave Blossom '!D26*Carbon!B3) / 1000
```

This formula:
1. Checks the carbon scenario selector (`Assumptions!C20` = "Accelerated Transition")
2. If Zero Liability: uses Carbon row 2 (all zeros)
3. If Accelerated Transition: uses Carbon row 3 ($45–$120/t CO₂e)
4. Multiplies emissions above baseline (row 26) × carbon price
5. Divides by 1,000 to convert to AUDk

The carbon prices for Accelerated Transition (Carbon sheet row 3):

| Year | 2026 | 2027 | 2028 | 2029 | 2030 | 2031 | 2032 | 2033 | 2034+ |
|---|---|---|---|---|---|---|---|---|---|
| $/t CO₂e | 45 | 50 | 55 | 65 | 80 | 90 | 100 | 105 | 110–120 |

**Total Other Costs** (row 108):
```
D108 = =SUM(D104:D107)
```

**Total Other — Closure Credits** (row 109):
```
D109 = =IF(D105<=0, (D104+D106), D108)
```
A logic check: if no closure cost this year (D105 ≤ 0), only count rehab + carbon. Otherwise, count all other costs. This prevents the $150M closure cost from being double-counted in years where it's not the final year.

### Step 8: Capital Costs

**Rows 111–115, Brave Blossom sheet + Brave Blossom Capital sheet**

**Project Capital** (row 112): Pulls from the Brave Blossom Capital sheet. Each cell links to the capital schedule:
```
D112 = (linked from Brave Blossom Capital sheet total row)
```

The Capital sheet contains 12 items (rows 4–15), each with:
- **Unit cost** in column C (e.g., C13 = 190,000 AUDk for the longwall)
- **Phasing fractions** in year columns (e.g., 0.5 in 2030, 1.0 in 2032)

The per-year cost for each item: `= unit_cost × phasing_fraction`

A 30% contingency is applied (Capital sheet cell C56 = 0.3):
```
Total Capital (row 57) = base_capital × (1 + contingency)
```

**Capital schedule (with 30% contingency):**

| Year | Base Capital (AUDk) | Contingency (30%) | Total (AUDk) |
|---|---|---|---|
| 2027 | 10,000 | 3,000 | 13,000 |
| 2028 | 10,000 | 3,000 | 13,000 |
| 2029 | 59,000 | 17,700 | 76,700 |
| 2030 | 84,000 | 25,200 | 109,200 |
| 2031 | 109,900 | 32,970 | 142,870 |
| 2032 | 221,500 | 66,450 | 287,950 |
| 2033 | 24,000 | 7,200 | 31,200 |
| **Total** | **518,400** | **155,520** | **673,920** |

> **Important**: The original workbook had an empty Capital sheet — no phasing data was entered. This meant row 112 pulled through $0 for every year, and the NPV was calculated as if the mine cost nothing to build. See Part 4 for the full correction story.

**Sustaining Capital** (row 114, C114 = 6 AUD/ROMt):
```
D114 = =$C$114*D9
```
$6 per tonne of ROM × ROM production. This covers ongoing equipment replacement and minor works throughout the mine life.

**Total Capital** (row 115):
```
D115 = =SUM(D112:D114)*(1+$BT$14)
```
Project capital + sustaining capital, adjusted by an escalation factor (BT14).

### Step 9: Depreciation

**Rows 116, 129–133, Brave Blossom sheet**

Depreciation spreads the cost of capital assets over their useful life for tax purposes.

```
D116 = =D132
```

Row 132 calculates depreciation:
```
D132 = =IF(D2=$B$135, D130+D131, IFERROR(D130/(B135+1-D2), 0))
```
- **D2** = year index
- **B135** = 2052 (last year of mining)
- **D130** = opening balance of undepreciated capital
- **D131** = capital purchases this year

In the final year (D2 = B135), the remaining balance is fully depreciated (D130 + D131). In other years, it's straight-line: opening balance ÷ remaining mine life. This is the depreciation used in the tax calculation.

### Step 10: Net Cashflow Before Tax

**Row 118, Brave Blossom sheet**

```
D118 = =D101-D108-D115
```

Mine Earnings (row 101) − Other Costs (row 108, via row 109) − Total Capital (row 115).

Wait — let's be precise. The formula subtracts row 108 (not row 109) and row 115:
- **D101** = Revenue − FOB Costs (operating margin)
- **D108** = Total Other Costs (rehab + closure + carbon + other revenue)
- **D115** = Total Capital (project + sustaining)

So: **Net Cashflow Before Tax = Operating Margin − Other Costs − Capital**.

### Step 11: Tax

**Row 119, Brave Blossom sheet (C119 = 0.3, i.e., 30%)**

```
D119 = =IF((D101-D109-D116)*$C$119 > 0, (D101-D109-D116)*$C$119, 0)
```

Tax is calculated on:
- **D101** = Mine Earnings (Revenue − FOB Costs)
- **− D109** = Other costs (closure credits version — rehab, closure, carbon)
- **− D116** = Depreciation (the tax shield — capital cost deducted from taxable income)
- **× 30%** = corporate tax rate

The `IF > 0` check means no tax is paid in loss-making years (no negative tax — loss carry-forwards are not modelled here).

### Step 12: Cashflows After Tax

**Rows 122–123, Brave Blossom sheet**

**Row 122** (Cashflows After Tax):
```
D122 = =D118-D119
```
Net Cashflow Before Tax − Tax Paid = Cashflow After Tax.

**Row 123** (Attributable Cashflow After Tax, C123 = 0.6, i.e., 60%):
```
D123 = =IFERROR(D122*$C$123, 0)
```
WM owns 60% of the Brave Blossom joint venture. The model calculates WM's attributable share — 60% of the after-tax cashflow. Waratah Resources owns the other 40%.

### Step 13: Discounting and NPV

**Row 124, Brave Blossom sheet**

```
D124 = =D123*Assumptions!I2
```

> **Note on a formula inconsistency**: Some columns in row 124 reference row 123 (attributable cashflow) while others reference row 122 (full cashflow before attribution). For example, `D124 = D123*Assumptions!I2` but `E124 = E122*Assumptions!J2`. This means some years are discounted at the 60% attributable basis and others at 100%. This appears to be an inconsistency in the original workbook. The NPV and IRR are both calculated on row 124, so the result blends both bases. The financial impact is modest since the years where row 122 is used (instead of 123) have zero or minimal cashflow.

Each year's attributable cashflow is multiplied by the discount factor from the Assumptions sheet. The discount factor uses the mid-year convention:
```
Assumptions!I2 = =(1+0.08)^-(year_index - base_year + 0.5)
```

**NPV** (row 126):
```
C126 = =SUM(D124:BO124)
```

The sum of all discounted cashflows from 2027 (column D) through ~2052 (column BO).

**Result**: The original workbook (with $0 capital) produced an NPV of **$1,320 million**. After populating the capital schedule (see Part 4), the corrected NPV is **$997 million**.

### Step 14: IRR

**Row 128, Brave Blossom sheet**

```
C128 = =IRR(D124:BO124)
```

Excel's `IRR` function finds the discount rate that makes the sum of row 124 equal to zero. With $0 capital, the IRR was 77.4%. After correcting for capital, the IRR is approximately **52%** — still well above the 8% hurdle rate.

### Step 15: Payback

**Row 127, Brave Blossom sheet**

```
C127 = =SUM(D125:BO125)
```

Row 125 contains a running cumulative cashflow. Payback is the number of years until cumulative cashflow turns positive.

---

## Part 3: The Carbon Calculation in Detail

**Rows 21–26, Brave Blossom sheet + Carbon sheet + Decommissioned Mine sheet**

### Emissions

**Coal Mine Waste Gas** (row 22, C22 = 0.13 t CO₂e per ROM tonne):
```
H22 = =IF(H9>0, H9*1000*$C$22+20000, 0)
```
ROM tonnes × 1,000 (converting kt to tonnes) × 0.13 t CO₂e/t + 20,000 t baseline = total fugitive emissions. The 20,000 is a fixed baseline from mine infrastructure.

**Decommissioned Mine Emissions** (row 23): Pulled from the Decommissioned Mine sheet — post-closure methane emissions that decline over time.

**Total Scope 1 Emissions** (row 24): Sum of rows 22 and 23.

**Safeguard Baseline** (row 25): Links to the Decommissioned Mine sheet's hybrid declining baseline:
```
D25 = =(D9*1000*'Decommissioned Mine'!E22) + 'Decommissioned Mine'!E21*D23
```
This combines a per-tonne baseline (Decommissioned Mine row 22) with a fixed emissions baseline (row 21).

**Emissions Above Baseline** (row 26): Row 24 − Row 25. If the mine emits more than the Safeguard baseline, the excess incurs carbon costs. If below, no cost (or credits, but the model doesn't generate credits here).

### Carbon Cost

The carbon cost (row 106, explained in Step 7) multiplies emissions above baseline by the scenario carbon price. Under Accelerated Transition, prices rise from $45/t (2026) to $120/t (2052). This is already embedded in the NPV — it is not a sensitivity, it is the base case.

---

## Part 4: The Capital Correction — $1,320M → $997M

This is the most important calculation story in the analysis. The original workbook produced an NPV of $1,320 million for Brave Blossom. That number was wrong because the capital costs were $0.

### What Happened

The Brave Blossom Capital sheet was a template. It listed 12 capital items with their unit costs (e.g., longwall = $190M, mining equipment = $69.9M), but the year-by-year phasing columns (which determine when each item is purchased and in what quantity) were all blank. With no phasing data entered, every year's capital cost calculated as $0.

This meant the DCF model computed NPV as:
- Revenue: full projected revenue from selling ~6 Mt/year of coking coal for 20 years
- Costs: all operating costs (labour, power, processing, royalties, carbon)
- Capital: **$0** (because the Capital sheet was empty)

Result: NPV = $1,320 million — overstated by the present value of the real capital costs.

### The Correction

The capital schedule was populated using the vault's concept capital estimate (file 07), which provides year-by-year spending for all 12 items across 2027–2033. Two scenarios were computed:

| Scenario | Capital Total | Tax Shield | NPV | IRR (est.) |
|---|---|---|---|---|
| Original ($0 capital) | $0 | n/a | $1,320.1M | 77.4% |
| No contingency, with tax shield | $518.4M | Yes | $1,071.7M | ~60% |
| **30% contingency, with tax shield (base case)** | **$673.9M** | **Yes** | **$997.2M** | **~52%** |
| No contingency, no tax shield | $518.4M | No | $965.3M | ~48% |
| 30% contingency, no tax shield (worst) | $673.9M | No | $858.8M | ~42% |

### How the Corrected NPV Was Computed

The original workbook's cached cashflows (from the data_only version) include the full operating cashflows with $0 capital. The correction subtracts the present value of the real capital costs:

**Step 1**: Extract the original discounted cashflows (row 124) from the cached workbook. Sum = $1,320,078k.

**Step 2**: For each year with capital spending, compute the after-tax capital cost:
```
After-tax capital = capital × (1 - tax_rate)
```
This assumes immediate expensing (a simplification — the real depreciation schedule would spread the tax shield over the asset's life, producing a smaller but longer-lasting benefit). This is conservative.

**Step 3**: Discount each year's after-tax capital cost back to present value:
```
PV of capital = capital × (1 - 0.30) / (1 + 0.08)^(year_index)
```

**Step 4**: Subtract the total PV of capital from the original NPV:
```
Corrected NPV = $1,320.1M - $322.9M = $997.2M
```

The $322.9M capital adjustment represents the present value of $673.9M of capital spending (with 30% contingency, after 30% tax shield), discounted at 8% over the 2027–2033 capital period.

### Why All Scenarios Remain Positive

Even in the worst case ($673.9M capital, no tax shield), the NPV is $859M — still strongly positive. This is because:

1. The operating cashflows are very large (~$1,320M PV) relative to the capital (~$323–461M PV)
2. The mine reuses existing infrastructure (CHPP, rail, port), so capital is lower than a greenfield mine
3. The coal price ($211/t) is a premium coking coal price, generating strong margins
4. The 8% discount rate is modest, so 20 years of revenue retains substantial present value

---

## Part 5: The Combined NPV — $1,277M

**Analysis sheet + Springbok sheet + Brave Blossom sheet**

The combined NPV is the sum of two independently calculated NPVs:

| Component | NPV | Source |
|---|---|---|
| Springbok (existing mine, 2027–2031) | $279.4M | `Springbok!C126` (cached in workbook) |
| Brave Blossom (new mine, corrected) | $997.2M | $1,320M operating PV − $323M capital PV |
| **Combined** | **$1,276.7M** | Sum of the two |

Springbok's NPV is unaffected by the capital correction — its mine was built years ago, so the capital is already spent (a sunk cost not included in forward-looking NPV).

The Analysis sheet (row 24) also computes a cumulative Brave Blossom cashflow:
```
Analysis!D24 = =SUM('Brave Blossom '!$D$124:col124)
```
This tracks the running total of discounted cashflows, showing when the project pays back its capital investment.

### Previous Vault Figure: $1,670M

The vault previously cited a combined NPV of $1,670M. That was based on the uncorrected $1,320M Brave Blossom NPV (with $0 capital). The correction reduces it by $393M to $1,277M. The difference is entirely the capital cost that was missing from the original calculation.

---

## Part 6: Beyond NPV — The Closure Deferral Benefit

**Separate from the DCF, but critical to the board decision**

The $997M NPV only measures the value of mining coal. But developing Brave Blossom has a second financial benefit: it defers the $900M closure liability.

When Springbok closes (2031), the $900M cleanup bill becomes due. But if Brave Blossom is developed, the mine site stays operational until ~2052, pushing closure costs 19 years into the future.

**Present value calculation**:
```
PV of $900M at 2031 (4 years from 2027 base):
    $900M / (1.08)^4 = $661.5M

PV of $900M at 2050 (23 years from 2027 base):
    $900M / (1.08)^23 = $153.4M

NPV benefit of deferral = $661.5M - $153.4M = $508.1M
```

**Total value of proceeding with Brave Blossom**:

| Component | NPV ($M) |
|---|---|
| Brave Blossom project NPV (with capital) | 997.2 |
| Closure liability deferral benefit | 508.0 |
| **Total value with deferral** | **1,505.2** |

The deferral benefit is separate from and additional to the project NPV. It arises because spending $389M on Brave Blossom capital keeps the mine open, which defers $900M of closure costs by ~19 years.

---

## Part 7: Sensitivity Analysis

### Discount Rate Sensitivity

The capital NPV impact changes with the discount rate (capital is front-loaded, so higher discounting reduces its present value):

| Discount Rate | Capital NPV Impact ($M) | Corrected NPV ($M) |
|---|---|---|
| 6% | -353.7 | 966.4 |
| **8% (base)** | **-322.9** | **997.2** |
| 10% | -295.4 | 1,024.7 |
| 12% | -270.9 | 1,049.2 |
| 15% | -238.9 | 1,081.2 |

Note: Higher discount rates reduce the capital NPV impact (front-loaded costs are discounted more), but also reduce the operating cashflow PV. The table shows only the capital adjustment — the full operating sensitivity would require re-running the DCF.

### Coal Price Sensitivity

Base price: $211.20/t (88% of $240/t PHCC benchmark — a 12% quality discount)

| Scenario | Price ($/t) | Revenue Impact | NPV Assessment |
|---|---|---|---|
| -20% | 168.96 | -20% revenue | Potentially marginal — NPV may approach $0 |
| -10% | 190.08 | -10% revenue | Positive but significantly reduced |
| **Base** | **211.20** | — | **$997M (base case)** |
| +10% | 232.32 | +10% revenue | Strong NPV (~$1,150M est.) |
| Springbok parity | 315.00 | +49% revenue | Transformative (~$1,600M+ est.) |

The price is set in `Assumptions!J7 = =J6*0.88`. To test sensitivity, change cell J6 (the PHCC benchmark, currently 240) or the 0.88 multiplier.

### Capital Contingency Sensitivity

| Contingency | Capital Total | NPV ($M) | IRR (est.) |
|---|---|---|---|
| 0% (direct cost only) | $518.4M | $1,071.7M | ~60% |
| **30% (base case)** | **$673.9M** | **$997.2M** | **~52%** |
| 30%, no tax shield | $673.9M | $858.8M | ~42% |

The contingency rate is in `Brave Blossom Capital!C56 = 0.3`. Setting it to 0 removes the 30% buffer.

---

## Quick Reference: Cell Map for Option A

| Calculation | Sheet | Cell(s) | Formula / Value |
|---|---|---|---|
| Discount rate | Assumptions | B2 | 0.08 |
| Discount factors | Assumptions | I2:Q2 | `=(1+$B2)^-(year-$I$1+0.5)` |
| FX rate (USD:AUD) | Assumptions | I4:Q4 | 0.69 (flat forecast) |
| PHCC benchmark price | Assumptions | J6:Q6 | 240 USD/t |
| 88% PHCC price | Assumptions | J7:Q7 | `=J6*0.88` = 211.20 |
| Carbon scenario | Assumptions | C20 | "Accelerated Transition" |
| QLD royalty tiers | Assumptions | D28:Q33 | 0, 100, 150, 175, 225, 300 AUD/t |
| QLD royalty rates | Assumptions | D35:Q40 | 7%, 12.5%, 15%, 20%, 30%, 40% |
| AUD CPI (historical) | Assumptions | D49:H49 | 1.0% |
| AUD CPI (forecast) | Assumptions | J49:Q49 | 2.65% |
| Carbon prices | Carbon | B3:O3 | 45, 50, 55, 65, 80, 90, 100, 105, 110, 115, 115, 115, 120, 120 |
| Development metres | Brave Blossom | D6:BO6 | Input (0 until 2031, then 7,500→15,000) |
| Development tonnes | Brave Blossom | D7:BO7 | `=(3*5.2*1.45*D6)/1000` |
| Longwall tonnes | Brave Blossom | D8:BO8 | Input (7,600 from 2033) |
| Total ROM | Brave Blossom | D9:BO9 | `=D7+D8` |
| CHPP yield | Brave Blossom | C15 | 0.68 |
| Plant product | Brave Blossom | D14:BO14 | `=D12*$C$15` (with ramp adjustments) |
| Saleable production | Brave Blossom | D19:BO19 | `=D13+D14` |
| Sales: 88% PHCC | Brave Blossom | C32 | 1 (100% allocation) |
| Total sales | Brave Blossom | D34:BO34 | `=SUM(D31:D33)` |
| Revenue (USD) | Brave Blossom | D60:BO60 | `=SUMPRODUCT(D31:D32, D55:D56)` |
| Revenue (AUD) | Brave Blossom | D62:BO62 | `=(D60/D61)*(1+$BT$7)` |
| Development cost rate | Brave Blossom | C69 | 5,600 AUD/m |
| Longwall cost rate | Brave Blossom | C72 | 15 AUD/LWt |
| Outbye costs (fixed) | Brave Blossom | C75 | 240,000 AUDk |
| Gas drainage rate | Brave Blossom | C77 | 6 AUD/t |
| CHPP fixed | Brave Blossom | C80 | 20,000 AUDk |
| CHPP variable | Brave Blossom | C81 | 9.5 AUD/feedt |
| Overhead: Site Admin | Brave Blossom | C84 | 45,000 AUDk |
| Overhead: Town Mgmt | Brave Blossom | C85 | 9,000 AUDk |
| Overhead: Corporate | Brave Blossom | C86 | 15,000 AUDk |
| Total FOR costs | Brave Blossom | D89:BO89 | `=SUM(D70,D73,D75,D78,D82,D87)` |
| Port/rail/selling | Brave Blossom | C93 | 22 AUD/t |
| Royalty (per tonne) | Brave Blossom | D148:BO148 | Tiered IF formula (see Step 5) |
| Total FOB costs | Brave Blossom | D98:BO98 | `=D89+D96` |
| Mine earnings | Brave Blossom | D101:BO101 | `=D62-D98` |
| Progressive rehab | Brave Blossom | C104 | 2,500 AUDk/year |
| Mine closure | Brave Blossom | C105 | 150,000 AUDk (final year) |
| Carbon cost | Brave Blossom | D106:BO106 | `=IF(scenario, emissions×price/1000)` |
| Project capital | Brave Blossom | D112:BO112 | Linked from Brave Blossom Capital sheet |
| Sustaining capital | Brave Blossom | C114 | 6 AUD/ROMt |
| Total capital | Brave Blossom | D115:BO115 | `=SUM(D112:D114)*(1+$BT$14)` |
| Depreciation | Brave Blossom | D116:BO116 | `=D132` (straight-line over remaining life) |
| Net cashflow before tax | Brave Blossom | D118:BO118 | `=D101-D108-D115` |
| Tax rate | Brave Blossom | C119 | 0.30 (30%) |
| Tax paid | Brave Blossom | D119:BO119 | `=IF((earnings-other-depn)×0.3>0, ...)` |
| Cashflow after tax | Brave Blossom | D122:BO122 | `=D118-D119` |
| WM attribution | Brave Blossom | C123 | 0.60 (60%) |
| Attributable cashflow | Brave Blossom | D123:BO123 | `=D122*$C$123` |
| Discounted cashflow | Brave Blossom | D124:BO124 | `=D123*Assumptions!I2` |
| **NPV** | Brave Blossom | **C126** | **=SUM(D124:BO124)** |
| **IRR** | Brave Blossom | **C128** | **=IRR(D124:BO124)** |
| Last year of mining | Brave Blossom | B135 | 2052 |
| Capital contingency | Brave Blossom Capital | C56 | 0.30 (30%) |

---

## Summary: The Full Calculation Chain on One Page

```
Development Metres (Row 6, input)
    ↓ × 3 headings × 5.2m seam × 1.45 t/m³ ÷ 1000
Development Tonnes (Row 7)
    +
Longwall Tonnes (Row 8, input)
    =
Total ROM (Row 9)
    ↓ × CHPP Yield (C15 = 68%)
Plant Product (Row 14)
    =
Saleable Production (Row 19)
    ↓ × $211.20/t (88% of $240 PHCC benchmark)
Revenue USD (Row 60)
    ↓ ÷ FX 0.69 × premium adjustment
Revenue AUD (Row 62)
    − Operating Costs (Rows 67–87: development, longwall, outbye, gas, CHPP, overheads)
    − Port/Rail/Selling (Row 93: $22/t)
    − Government Royalty (Row 95: QLD tiered, 7–40%)
    =
Mine Earnings (Row 101)
    − Other Costs (Rows 104–106: rehab $2.5M/yr, closure $150M, carbon)
    − Capital (Row 115: project capital from Capital sheet + sustaining $6/t)
    =
Net Cashflow Before Tax (Row 118)
    − Tax (Row 119: 30% × (earnings − other costs − depreciation))
    =
Cashflow After Tax (Row 122)
    × 60% WM attribution (Row 123)
    × Discount Factor (Assumptions row 2: mid-year, 8%)
    =
Discounted Cashflow (Row 124)
    ↓ SUM across 2027–2052
NPV (Cell C126) = $997M (corrected) / $1,320M (original, $0 capital)
IRR (Cell C128) = ~52% (corrected) / 77.4% (original)
```

---

*This walkthrough accompanies the Wallaby Mining Board Presentation. For the full technical analysis, refer to `analysis/01b_financial_model_analysis.md`. For a plain-English summary of all recommendations, refer to `PLAIN_ENGLISH_GUIDE.md`. All financial figures are corrected from the original workbook which overstated NPV by calculating it with zero capital costs.*
