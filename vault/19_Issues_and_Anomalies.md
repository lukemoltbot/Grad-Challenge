# 19 — Issues & Anomalies Found in 2024 Closure Cost Workbook

> **Source**: `2024_Springbok_Planned_Closure_Costs.xlsx`
> **Audit date**: 2026-08-29
> **Companion**: `18_Audit_Results.md` (calculation verification), `20_Vault_Crossrefs_and_SMART.md` (deliverable mappings)
> **Note**: All arithmetic is correct (see file 18) — these are structural/data issues, not calculation errors.

---

## Issue Register

| # | Issue | Severity | Potential Impact ($) | Type |
|---|-------|----------|---------------------|------|
| B1 | Duplicate TSF item 2.002 | 🔴 HIGH | $39.5M–$43.8M | Potential double-count |
| B2 | Duplicate Domain 4I / 4J | 🟡 MEDIUM | $4.76M | Potential double-count |
| B3 | Missing execution-phase holding cost detail | 🟡 LOW | N/A (data gap) | Missing data |
| B4 | Domain 3 sub-header structure | 🟢 INFO | None | Presentation |
| B5 | "Zeroed" comments on non-zero items | 🟢 INFO | None (misleading) | Documentation |
| B6 | 2021 double-counting error (self-noted) | 🟢 INFO | None (already corrected) | Historical |
| B7 | 10-year-old rate escalations | 🟡 MEDIUM | Unknown (rate basis) | Rate assumption |
| B8 | 100t fleet costed at 400t fleet rate | 🟡 MEDIUM | Unknown (overstated) | Rate substitution |

---

### B1. 🔴 POTENTIAL DUPLICATE — Domain 2, Item 2.002 (Tailings Storage Facility)

**Severity: HIGH — potential $39.5M or $43.8M double-count**

**Location**: Planned Closure Detail sheet, rows 129–130

| Row | Item # | Description | Qty | Rate (A$) | Cost (A$) |
|-----|--------|-------------|-----|-----------|-----------|
| 129 | 2.002 | Tailings Storage Facility (Out of Pit) | 1 | 39,538,755.34 | 39,538,755.34 |
| 130 | 2.002 | Tailings Storage Facility (Out of Pit) | 1 | 43,824,000.00 | 43,824,000.00 |

**Evidence of duplication**:
- Identical item number: 2.002
- Identical description: "Tailings Storage Facility (Out of Pit)"
- Identical comment: "Rate provided by site. Volumes and Rates. Hydraulic mining: 5473k m3 at $4.47. Reject Removal: 623k m3 at $12.46. Capping: 133k m3 at $22.73. TOTAL MAT: 6229k m3 at $5.66"
- Both have qty = 1 (lump sum)

**Analysis**:
Row 129 ($39.5M) appears to be a calculated rate derived from the volume × unit rate breakdown in the comment (6,229k m³ × $5.66/m³ ≈ $35.3M, with additional costs bringing it to $39.5M). Row 130 ($43.8M) appears to be a separate site-provided lump sum for the same work. Both are included in the Domain 2 total of $126.4M.

**If one is a duplicate**:
- Remove row 130 ($43.8M) → Domain 2 drops from $126.4M to **$82.6M** (–34.7%)
- Remove row 129 ($39.5M) → Domain 2 drops from $126.4M to **$86.9M** (–31.3%)

**Flow-on impact**: With 35% contingency applied, the grand total impact would be ~$54M–$59M. However, the grand total is exactly $900M — this may mean the duplicate was absorbed by a balancing adjustment, or the $900M was a round-number target the workbook was built to.

**Recommendation**: Flag in presentation as a specific finding. Engage independent quantity surveyor to verify which rate applies.

---

### B2. 🟡 POTENTIAL DUPLICATE — Domain 4I vs 4J (Identical Values)

**Severity: MEDIUM — potential $4.76M double-count**

**Location**: Planned Closure Detail sheet, rows 228–233 (Domain 4I) vs rows 234–239 (Domain 4J)

**Side-by-side comparison**:

| Attribute | Domain 4I (row 228) | Domain 4J (row 234) |
|-----------|---------------------|---------------------|
| Stated total | 4,756,679.42 | 4,756,679.42 |
| D11 bulk push qty | 1,583,588.33 m³ | 1,583,588.33 m³ |
| D11 bulk push rate | $2.3955/m³ | $2.3955/m³ |
| D11 bulk push cost | 3,793,437.75 | 3,793,437.75 |
| Load & haul qty | 168,154.98 m³ | 168,154.98 m³ |
| Load & haul rate | $5.7283/m³ | $5.7283/m³ |
| Load & haul cost | 963,241.67 | 963,241.67 |
| Haul distance label | "1-2km" | "<1km" |

**Evidence**:
- Identical quantities, rates, and costs to the cent
- Identical comments: "Volume provided in Material Volume RFI. Push Length 100-150m."
- Only difference: haul distance label (1-2km vs <1km)
- Row 236 (4J rehab) comment: "Zeroed as these are now consolidated at the bottom of the active mining domain" — yet earthworks line items below still carry full cost

**Analysis**:
4J appears to be a copy of 4I where the rehabilitation sub-category header was zeroed (per the comment), but the earthworks and mining line items beneath were NOT zeroed. The identical quantities suggest these represent the same physical area, not two different areas.

**If 4J is a duplicate**:
- Domain 4 drops from $216.5M to **$211.8M** (–2.2%)
- Direct works drops from $494.1M to **$489.4M**
- Grand total impact with 35% contingency: ~$6.4M

**Recommendation**: Verify with site whether 4I and 4J represent distinct physical areas or are the same area duplicated.

---

### B3. 🟡 Missing Execution-Phase Holding Cost Detail

**Severity: LOW (data gap, not an error)**

**Location**: Holding Costs sheet, column G ("Execution 5 yr")

The 5-year execution period ($97.98M) is presented as a single lump-sum column with no year-by-year breakdown, unlike the post-execution period which has 15 individual year columns.

**Impact**:
- Cannot verify the year 1–5 decline rate for the execution phase
- Cannot model the transition from execution-phase costs to post-execution-phase costs
- Cannot confirm whether execution-phase costs are flat or declining

**What IS known**:
- The "Holding" column (F) shows an annual figure that, multiplied by 5, approximately equals the execution total
- The post-execution period declines at exactly 10% per annum
- The ratio of execution to post-execution is 97.98M : 134.97M = 42% : 58%

**Recommendation**: Request year-by-year execution-phase breakdown from WM for NPV modeling.

---

### B4. 🟢 Domain 3 Sub-Header Structure

**Severity: INFORMATIONAL — no calculation impact**

**Location**: Planned Closure Detail sheet, rows 149–165 (Domain 3)

The "All Overburden emplacement areas" sub-header (row 150) has a stated value of $26.4M, but the line items below it (rows 151–165) show $0 in their G column. The domain total is correct (verified in grand total reconciliation), but the data appears structured so that the sub-header is a formula reference to a total calculated from another sheet (likely the BulkPush sheet volumes × rates from the rate library).

This is a workbook structure/presentation issue, not a calculation error. The costs exist but are referenced rather than listed inline.

---

### B5. 🟢 "Zeroed" Comments on Non-Zero Items

**Severity: INFORMATIONAL — misleading but not erroneous**

**Location**: Multiple rows in Planned Closure Detail sheet

Several line items have comments saying "Zeroed" but still carry cost:

| Row | Item | Comment | Actual Cost (A$) |
|-----|------|---------|-------------------|
| 131 | 2.003 | "Zeroed as this will be capped at LOM" | 11,922,564 |
| 132 | 2.004 | "Zeroed as this will be capped at LOM" | 8,635,992 |
| 133 | 2.005 | "Zeroed as this will be capped at LOM" | 11,118,403 |
| 236 | 4.035 | "Zeroed as these are now consolidated" | 0 (header only) |

**Analysis**:
The "Zeroed" comments appear to refer to specific sub-categories being consolidated elsewhere (e.g., the rehabilitation sub-category is zeroed because costs are consolidated under the earthworks sub-category), not the entire row being zeroed. The costs are legitimate — the comments are referring to the reclassification, not removal. However, the comment text is misleading because it says "zeroed" when the row cost is non-zero.

**Impact on audit**: Required manual verification of each "zeroed" row to confirm the cost was moved, not lost. All verified as present in the domain totals.

---

### B6. 🟢 2021 Double-Counting Error (Self-Noted in Workbook)

**Severity: INFORMATIONAL — already corrected, validates audit approach**

**Location**: Cell H72, Planned Closure Detail sheet

The workbook contains a note: "Cells R777-779 were double-counted in 2021 review (noted in H72)."

This indicates that a double-counting error occurred in the 2021 triennial review and was identified and corrected in the 2024 review. This is significant because:
1. It establishes precedent for double-counting errors in this workbook
2. It validates the need for independent audit of ERC workbooks
3. It supports the credibility of Issues B1 and B2 as genuine potential duplicates, not just formatting quirks

---

### B7. 🟡 10-Year-Old Rate Escalations

**Severity: MEDIUM — rate basis may be outdated**

**Location**: Rate library (Sheet 6), various rate build-up comments

Several rates in the WM 2024 rate library include "16% escalation for site specific rates in 2014" applied to 2014 base rates. This means:
- Some rates are based on 2014 costs with a 16% escalation factor
- These escalated rates have not been updated for 2025–2026 cost movements
- Over 10 years, construction cost inflation has been significantly higher than 16% in most categories
- However, some rates have been updated based on the "2024 WM schedule of rates"

**Impact**:
- If rates are understated (2014 base + 16% < 2024 actual costs), the $900M estimate may be conservative (too low)
- If rates are overstated (2014 base + 16% > 2024 actual costs due to productivity gains or technology improvements), the estimate may be inflated
- Without knowing which rates use 2014 escalations vs 2024 actual rates, cannot determine net direction

**Specific rates affected**: Not individually tagged in the workbook — the escalation note appears in the rate build-up comments. A line-by-line rate basis review would be needed to categorise each rate.

**Recommendation**: Challenge rates that are based on 2014 escalations — request 2024 actual cost data for verification.

---

### B8. 🟡 100t Fleet Costed at 400t Fleet Rate

**Severity: MEDIUM — potential overstatement of haulage costs**

**Location**: Planned Closure Detail sheet, multiple rows in Domain 4

Multiple load-and-haul line items in Domain 4 have the comment: "Bulk push as provided in site RFI. Note that data provided was for a 100t fleet, however as there is no WM rate for this, the 400t fleet rate has been used."

**Affected rows**:
- Row 233 (Domain 4I): Load and haul 400t fleet, flat, 1-2km — qty 168,155 m³ at $5.73/m³
- Row 239 (Domain 4J): Load and haul 100t fleet, flat, <1km — qty 168,155 m³ at $5.73/m³
- Row 247 (Domain 4K): Load and haul 400t fleet, flat, 1-2km — qty 436,559 m³ at $5.73/m³
- Additional rows in Domains 4L, 4O, 4P

**Analysis**:
A 400t fleet (400t digger / 190t truck) is significantly larger and more expensive per hour than a 100t fleet. Using the 400t fleet rate for 100t fleet volumes means:
- The unit rate ($5.73/m³) is based on larger, more expensive equipment
- The actual cost of using 100t equipment would likely be lower (smaller equipment, lower hourly rate, but more cycles)
- However, 100t equipment may require more cycles for the same volume, potentially making it MORE expensive per m³ if productivity is lower

**Net direction**: Uncertain — depends on whether the 100t fleet's lower hourly rate offsets its lower productivity. But the workbook acknowledges the substitution is incorrect by noting it in the comments.

**Recommendation**: Request WM develop a 100t fleet rate for accurate costing, or justify why the 400t rate is a conservative proxy.

---

## Summary

The workbook's arithmetic is flawless — every multiplication, addition, and percentage checks out to the cent. However, the audit found:

1. **Two potential duplicate line items** (B1: $39.5–$43.8M, B2: $4.76M) that inflate the total
2. **Two rate assumption issues** (B7: outdated escalations, B8: fleet substitution) that may misstate costs in either direction
3. **One data gap** (B3: missing execution-phase detail) that limits NPV modeling
4. **Three informational items** (B4, B5, B6) that are presentation/documentation issues

The combination of B1 (the 2021 double-count precedent) and B6 (the 2024 self-noted correction) creates a strong argument for independent review: the same type of error has occurred before in this exact workbook.
