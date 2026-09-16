"""
LESSON 1 - openpyxl
Miltiades, this is a real check on your own file, not a toy example.

WHAT IT DOES
It opens an Annex 3, reads the budget each partner is given on the Cover page,
then adds up that partner's own sheet, and tells you if the two agree.

WHY THIS ONE
This is the check I run before I trust any budget file you send me. When the
EPIRUSMEDEYE 1st modification arrived, this is what showed me that only Dropull
had changed and only by 9,000 - which is how I knew it was a real modification
and not a stale draft.

HOW TO RUN IT
1. Put this file and your Annex 3 in the SAME folder, for example C:\\interreg
2. Edit the FILE line below so it matches your file name exactly
3. Open that folder in VS Code, then Terminal > New Terminal, and type:
       python lesson01_check_annex3.py

Never run it from C:\\Windows\\System32. That is a system folder, Windows may
block writes there, and it is not where your work belongs.

You will probably see two yellow warnings about "Conditional Formatting
extension is not supported" and "Data Validation extension". Ignore them.
openpyxl is telling you it does not understand the colour rules in the file.
It reads every number correctly. Warnings are not errors.
"""

import openpyxl

# ---- the only two lines you edit ----
FILE = "Annex 3_JoBC_revised_22.6_1st call_EPIRUSMEDEYE 1ST MODIFICATION.xlsx"
SHEETS = ["LB (PP01)", "PP02", "PP03", "PP04"]   # MINOR-MED also has "PP05"
# -------------------------------------

# data_only=True means "give me the VALUE Excel calculated, not the formula".
# Without it you get strings like "=SUM(J10:J208)" and nothing adds up.
wb = openpyxl.load_workbook(FILE, data_only=True)

# --- step 1: read the cover page ---
# We do not know which row each partner is on, so instead of guessing a row
# number we walk the rows and keep the ones that LOOK like a partner line:
# first cell is "LB (..." or "PB" + a number, and the fourth value is a number.
cover = {}
for row in wb["Cover page"].iter_rows(min_row=1, max_row=40, max_col=11,
                                      values_only=True):
    cells = [c for c in row if c not in (None, "")]
    if len(cells) < 4 or not isinstance(cells[0], str):
        continue
    ref = cells[0]
    if not (ref.startswith("LB (") or (ref.startswith("PB") and ref[2:].isdigit())):
        continue
    try:
        budget = float(cells[3])
    except (TypeError, ValueError):
        continue          # that cell was not a number, so it was not a budget
    if budget > 0:
        cover[ref] = budget

print(f"Cover page lists {len(cover)} partners\n")

# --- step 2: add up each partner sheet and compare ---
grand_cover = 0.0
grand_lines = 0.0

for ref, sheet in zip(cover, SHEETS):
    ws = wb[sheet]
    total = 0.0
    count = 0
    for row in ws.iter_rows(min_row=10, max_row=208, max_col=12, values_only=True):
        wp = row[0]
        if not (wp and str(wp).startswith("WP")):
            continue      # not a cost line - a heading, a blank, a note
        try:
            total += float(row[9])   # column J is the amount
            count += 1
        except (TypeError, ValueError):
            continue

    diff = total - cover[ref]
    verdict = "OK" if abs(diff) < 1.0 else f"MISMATCH of {diff:,.2f}"
    print(f"{ref:10s} {sheet:10s} {count:4d} cost lines "
          f"{total:14,.2f}  vs cover {cover[ref]:14,.2f}   {verdict}")

    grand_cover += cover[ref]
    grand_lines += total

print(f"\n{'PROJECT':21s} {'':4s}              {grand_lines:14,.2f} "
      f" vs cover {grand_cover:14,.2f}")

# --- step 3: the question that actually matters ---
# Compare the project total against the APPROVED application form.
# If these differ, the file you are holding is not the file the Managing
# Authority approved, and you need to find out why before you use it.
AF_TOTAL = 1362373.17     # EPIRUSMEDEYE approved AF. MINOR-MED is 887786.27
gap = grand_cover - AF_TOTAL
print(f"{'Approved application form':21s}              {AF_TOTAL:14,.2f}")
print(f"{'Difference':21s}              {gap:14,.2f}")
if abs(gap) < 1.5:
    print("\n-> This file matches the approved application form.")
else:
    print(f"\n-> This file is {gap:+,.2f} against the approved form. "
          f"That is either an approved modification or a draft. Find out which.")

# THREE IDEAS FOR WHEN THIS RUNS
# 1. Print which WP each cost line belongs to and total per WP.
# 2. Change FILE and SHEETS to the MINOR-MED Annex 3 and AF_TOTAL to 887786.27.
# 3. Make it print only the partners that do NOT match, so silence means fine.
