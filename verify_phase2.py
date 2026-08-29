#!/usr/bin/env python3
"""Ad-hoc verification for Phase 2 deliverables (final run)."""
import re, os, sys, subprocess

BASE = "/Users/lukemoltbot/Grad-Challenge"
errors = []
checks = 0

def check(name, condition, detail=""):
    global checks
    checks += 1
    status = "PASS" if condition else "FAIL"
    print(f"  [{status}] {name}" + (f" -- {detail}" if detail else ""))
    if not condition:
        errors.append(name)

# 1. HTML deck
print("\n=== HTML Deck (slides/deck.html) ===")
with open(os.path.join(BASE, "slides/deck.html"), 'r') as f:
    html = f.read()

slides = re.findall(r'data-slide="(\d+)"', html)
check("HTML has 20 slides", len(slides) == 20, f"found {len(slides)}")
notes = re.findall(r'data-note="(\d+)"', html)
check("HTML has 20 speaker note blocks", len(notes) == 20, f"found {len(notes)}")
check("HTML has keyboard navigation", "addEventListener('keydown'" in html)
check("HTML has presenter notes panel", 'id="presenter-notes"' in html)
check("HTML has progress bar", 'id="progress-bar"' in html)
check("HTML has $997M", "$997M" in html)
check("HTML has $1,277M", "$1,277M" in html)
check("HTML has ~52% IRR", "~52%" in html)
check("HTML has $180.3M", "$180.3M" in html)
check("HTML has $185M net positive", "$185M" in html)
check("HTML no $1,320M", "$1,320M" not in html)
check("HTML no $1,670M", "$1,670M" not in html)
check("HTML no 77.4%", "77.4%" not in html)

# 2. Speaker notes
print("\n=== Speaker Notes (slides/speaker_notes.md) ===")
with open(os.path.join(BASE, "slides/speaker_notes.md"), 'r') as f:
    notes_md = f.read()

header_times = re.findall(r'\((\d+)s\)', notes_md)
check("Speaker notes has 14 timed slides (header format)", len(header_times) == 14, f"found {len(header_times)}: {header_times}")
if header_times:
    total = sum(int(d) for d in header_times)
    check("Total timing = 900s", total == 900, f"got {total}s")
    expected = [30,75,90,75,90,75,75,75,60,60,60,45,45,45]
    actual = [int(d) for d in header_times]
    check("Timing matches expected breakdown", actual == expected, f"got {actual}")

check("Notes has $997M", "$997M" in notes_md or "$997 million" in notes_md)
check("Notes has $1,277M", "$1,277M" in notes_md or "$1.28 billion" in notes_md)
check("Notes has $180.3M", "$180.3M" in notes_md or "$180.3 million" in notes_md)
check("Notes has $185M", "$185M" in notes_md or "$185 million" in notes_md)
check("Notes no $1,320M", "$1,320M" not in notes_md and "$1.32 billion" not in notes_md)
check("Notes no $1,670M", "$1,670M" not in notes_md and "$1.67 billion" not in notes_md)

# 3. Q&A file
print("\n=== Q&A Prep (analysis/07_qa_preparation.md) ===")
with open(os.path.join(BASE, "analysis/07_qa_preparation.md"), 'r') as f:
    qa = f.read()

check("Q&A no $1,320M", "$1,320M" not in qa)
check("Q&A no $1,670M", "$1,670M" not in qa)
check("Q&A no 77.4%", "77.4%" not in qa)
check("Q&A has $997M", "$997M" in qa)
check("Q&A has $1,277M", "$1,277M" in qa)
check("Q&A has ~52%", "~52%" in qa)
check("Q&A has $180.3M", "$180.3M" in qa)
check("Q&A has $185M", "$185M" in qa)

draft_markers = re.findall(r'\[DRAFT[^\]]*\]', qa)
check("Q&A no [DRAFT] placeholders", len(draft_markers) == 0, f"found {len(draft_markers)}: {draft_markers[:3]}")

check("Q&A has solo presenter framework", "Solo Presenter" in qa and "solo presenter" in qa.lower())
check("Q&A has topic-to-appendix mapping", "Topic-to-Appendix" in qa or "topic-to-appendix" in qa.lower())
check("Q&A status CONFIRMED", "CONFIRMED" in qa)

# 4. Analysis files
print("\n=== Analysis files - old NPV figures absent ===")
analysis_files = [
    "02_brave_blossom_swot.md",
    "03_closure_liability_review.md",
    "04_other_projects_and_workplan.md",
    "05_recommendation_gonogo.md",
    "06_slide_structure_outline.md",
]
for fname in analysis_files:
    with open(os.path.join(BASE, "analysis", fname), 'r') as f:
        content = f.read()
    check(f"{fname}: no $1,320M", "$1,320M" not in content)
    check(f"{fname}: no $1,670M", "$1,670M" not in content)
    check(f"{fname}: no 77.4%", "77.4%" not in content)

# 5. TASK_LIST.md
print("\n=== TASK_LIST.md ===")
with open(os.path.join(BASE, "TASK_LIST.md"), 'r') as f:
    tl = f.read()
check("TASK_LIST references deck.html", "deck.html" in tl)
check("TASK_LIST references speaker_notes.md", "speaker_notes.md" in tl)
check("TASK_LIST has Phase 2 complete", "Phase 2" in tl and "COMPLETE" in tl)
check("TASK_LIST references solo presenter", "solo presenter" in tl.lower() or "Solo presenter" in tl)

# 6. Git status
print("\n=== Git status ===")
result = subprocess.run(["git", "-C", BASE, "log", "--oneline", "-1"], capture_output=True, text=True)
check("Latest commit exists", result.returncode == 0, result.stdout.strip())
print(f"  Latest: {result.stdout.strip()}")

result2 = subprocess.run(["git", "-C", BASE, "status", "--porcelain"], capture_output=True, text=True)
clean = result2.stdout.strip() == ""
check("Working tree clean", clean, f"uncommitted: {result2.stdout.strip()[:100]}")

# Summary
print(f"\n{'='*50}")
print(f"RESULTS: {checks - len(errors)}/{checks} passed, {len(errors)} failed")
if errors:
    print(f"FAILURES: {', '.join(errors)}")
    sys.exit(1)
else:
    print("ALL CHECKS PASSED")
    sys.exit(0)
