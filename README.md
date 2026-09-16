# Interreg Excel toolkit

Python for reading and checking Interreg budget and procurement spreadsheets —
written as teaching material for a project manager learning Python against his
own real files, rather than against a tutorial.

Everything here is commented to be read, not just run.

## The files

| File | What it does |
|---|---|
| `lesson01_check_annex3.py` | Opens an Annex 3 (Justification of Budget), adds up each partner sheet, checks it against the cover page, then checks the project total against the approved application form. |
| `lesson02_your_script_fixed.py` | A beginner's first openpyxl script, corrected line by line with the reason for each change. Covers the column/row/cell confusion, and the one mistake that can destroy a live budget file. |
| `compare_techsheets.py` | Compares successive MIS technical-sheet exports and reports what actually changed — keyed on meaning rather than on cell position. |

## Three things worth taking from this

**1. `data_only=True` plus `save()` destroys every formula.**

```python
wb = load_workbook("budget.xlsx", data_only=True)   # correct for READING
wb.save("budget.xlsx")                              # now it is a dead file
```

`data_only=True` loads the values Excel last cached instead of the formula
strings, and openpyxl only ever holds one or the other. So on save it writes
back what it holds. Verified: a cell containing `=SUM(A1:A2)` showing `42`
comes back as the literal number `42`. No warning. The file looks identical
until somebody edits an input and nothing recalculates.

Rule: if you loaded with `data_only=True`, never save that workbook. And never
save over the file you opened, whatever the flags.

**2. Do not compare spreadsheets by cell position.**

Two exports of the same report had 127 and 120 rows. From the first inserted
row onwards everything is shifted, so a positional comparison reports hundreds
of differences that are nothing but the offset, and the two real findings drown
in the noise.

Key each value on what it *means* — here `(beneficiary code, deliverable, cost
category)` — and then compare those keys. Rows can move, blocks can be
reordered, entries can be inserted or deleted, and the answer does not change.

**3. Make the code prove it understood the file.**

The dangerous failure with spreadsheets is almost never a crash. It is a number
that is wrong and looks fine. Two habits that catch it:

- *Prove the column mapping instead of assuming it.* One of these reports has a
  row of codes that looks like a column map and is not — the codes sit in
  different columns for different sections. So the script asserts that the
  category columns add up to the printed row total on every single row. If the
  mapping is wrong it stops, rather than producing a confident wrong answer.

- *Make the diff close.* The sum of the individual changes must equal the
  movement in the grand total. A comparison that quietly *misses* a change is
  the only genuinely dangerous outcome.

Two real bugs these habits caught, both of which produced no error at all:

- A join matched on partner name instead of on the MIS beneficiary code. One
  file names partners in English, the other in Greek. It matched 17 rows of 76
  and reported nothing wrong.
- A rule that applied to Greek partners only filtered on a country field — and
  that field was wrong on one row in the source file. The answer came out
  €348,145 instead of €605,473. No error, just a smaller number.

**4. A related trap: numbers stored as text.**

In one of these procurement plans every single amount was stored as text
(`'2016.13'` rather than `2016.13`). Excel's `SUM` over that column returns
`0` and `COUNT` returns `0` — the sheet cannot be totalled, and any formula
pointing at it is silently wrong rather than broken. The cells look completely
normal on screen. Both scripts coerce text numbers and count how often they had
to.

## Running them

```
pip install openpyxl
python lesson01_check_annex3.py
```

Put the spreadsheet in the same folder as the script, or give a full path.
Python looks in the current working directory, which is usually the folder your
editor opened rather than the folder the script sits in — so when a file is
"missing", print `os.getcwd()` first.
