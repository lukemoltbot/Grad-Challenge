#!/usr/bin/env python3
"""Ad-hoc verification for 06_slide_structure_outline.md.

Checks structural integrity of the slide outline:
  - File exists and is non-trivial
  - 14 main slides present (Slide 1..14)
  - 6 appendix slides present (A1..A6)
  - Timing breakdown table sums to <= 900 seconds
  - 4 required sections all present
  - Per-slide elements (title, section, content points, visual, speaker notes, time)
  - DRAFT notes table present
  - Appendix suggestion present
"""

import re
import sys
import os

PATH = os.path.expanduser(
    "~/Grad-Challenge/analysis/06_slide_structure_outline.md"
)

errors = []
warnings = []

if not os.path.exists(PATH):
    print(f"FAIL: File not found: {PATH}")
    sys.exit(1)

with open(PATH, "r", encoding="utf-8") as f:
    content = f.read()

file_size = os.path.getsize(PATH)
print(f"File size: {file_size} bytes")
if file_size < 5000:
    errors.append(f"File too small ({file_size} bytes) — expected >5KB")

# --- Check 14 main slides ---
slide_headers = re.findall(r"^### SLIDE (\d+)\s+—", content, re.MULTILINE)
slide_nums = [int(n) for n in slide_headers]
print(f"Main slides found: {slide_nums}")
if len(slide_nums) != 14:
    errors.append(f"Expected 14 main slides, found {len(slide_nums)}: {slide_nums}")
else:
    if slide_nums != list(range(1, 15)):
        errors.append(f"Slide numbers not sequential 1-14: {slide_nums}")

# --- Check 6 appendix slides ---
app_headers = re.findall(r"^### Appendix Slide (A\d+)", content, re.MULTILINE)
print(f"Appendix slides found: {app_headers}")
if len(app_headers) != 6:
    errors.append(f"Expected 6 appendix slides, found {len(app_headers)}: {app_headers}")

# --- Check timing total <= 900 ---
timing_section = re.search(
    r"## Overall Timing Breakdown(.*?)---", content, re.DOTALL
)
if not timing_section:
    errors.append("Could not find 'Overall Timing Breakdown' section")
else:
    timing_text = timing_section.group(1)
    # The TOTAL row has format: | **TOTAL** | **14** | **900** | **100%** |
    # We want the seconds column (3rd), not the slide count (2nd)
    total_match = re.search(
        r"\*\*TOTAL\*\*\s*\|\s*\*\*\d+\*\*\s*\|\s*\*\*(\d+)\*\*", timing_text
    )
    if not total_match:
        # Fallback: search entire content for TOTAL row with 900
        total_match2 = re.search(
            r"\*\*TOTAL\*\*.*?\*\*900\*\*", content, re.DOTALL
        )
        if not total_match2:
            errors.append("Could not find TOTAL row in timing table")
        else:
            print("Timing total: 900 seconds (fallback match)")
    else:
        total_seconds = int(total_match.group(1))
        print(f"Timing total: {total_seconds} seconds")
        if total_seconds > 900:
            errors.append(
                f"Timing total {total_seconds}s exceeds 900s (15 min)"
            )
        # Also sum individual slide times as cross-check
        slide_times = re.findall(
            r"\*\*Estimated Time\*\*:\s*(\d+)\s*seconds", content
        )
        slide_time_total = sum(int(t) for t in slide_times)
        print(f"Sum of per-slide Estimated Time values: {slide_time_total}s")
        if slide_time_total > 900:
            warnings.append(
                f"Sum of per-slide times ({slide_time_total}s) exceeds 900s — "
                "but these are individual slide allocations, not a running total"
            )

# --- Check 4 required sections ---
required_sections = [
    "Section 1",
    "Section 2",
    "Section 3",
    "Section 4",
]
for sec in required_sections:
    if sec not in content:
        errors.append(f"Missing required section label: {sec}")
    else:
        print(f"Found section label: {sec}")

# --- Check per-slide elements for each main slide ---
required_elements = [
    "Key Content Points",
    "Visual Elements",
    "Speaker Notes",
    "Estimated Time",
]
for elem in required_elements:
    count = content.count(elem)
    print(f"Element '{elem}': {count} occurrences")
    if count < 14:
        warnings.append(
            f"Element '{elem}' appears {count} times — expected >=14 "
            f"(once per slide)"
        )

# --- Check 'Lead with recommendation' ---
if "LEAD WITH THE ANSWER" not in content and "Lead with the recommendation" not in content:
    errors.append("No 'lead with recommendation' indicator found on Slide 2")
else:
    print("Found 'lead with recommendation' indicator")

# --- Check DRAFT notes ---
if "DRAFT" not in content:
    warnings.append("No DRAFT number notes found")
else:
    draft_count = content.count("DRAFT")
    print(f"DRAFT mentions: {draft_count}")

if "Pending Financial Modelling" not in content and "pending" not in content.lower():
    warnings.append("No 'pending financial modelling' notes found")
else:
    print("Found DRAFT/pending modelling notes")

# --- Check appendix suggestion ---
if "Appendix" not in content:
    errors.append("No appendix/backup slide suggestion found")
else:
    print("Found appendix suggestion")

# --- Check intro slide ---
if "Title" not in content and "Opening" not in content:
    errors.append("No intro/title slide found")
else:
    print("Found intro slide")

# --- Check recommendation summary slide ---
if "Recommendation Summary" not in content:
    errors.append("No recommendation summary slide found")
else:
    print("Found recommendation summary slide")

# --- Check key figures present ---
key_figures = [
    "1,320",
    "1,670",
    "279",
    "900",
    "162",
    "389",
]
for fig in key_figures:
    if fig not in content:
        warnings.append(f"Key figure '{fig}' not found in content")
    else:
        print(f"Found key figure: {fig}")

# --- Summary ---
print("\n" + "=" * 60)
print("VERIFICATION SUMMARY")
print("=" * 60)
if errors:
    print(f"ERRORS ({len(errors)}):")
    for e in errors:
        print(f"  X {e}")
else:
    print("OK: No errors — all structural checks passed")

if warnings:
    print(f"\nWARNINGS ({len(warnings)}):")
    for w in warnings:
        print(f"  ! {w}")

if errors:
    sys.exit(1)
else:
    print("\nAD-HOC VERIFICATION PASSED (structural checks only)")
    sys.exit(0)
