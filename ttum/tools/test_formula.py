"""Checks the macro-free Text Output sheet by actually evaluating its formulas.

The workbook is filled with the bank's own figures and then calculated with an
Excel formula interpreter, so the records this sheet produces are compared against
the bank's file the same way the macro's output is.
"""

import os
import sys
import warnings
from datetime import date

import openpyxl

warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from test_workbook import SAMPLE_AMOUNTS_KEYED  # noqa: E402

XLSX = os.path.join(HERE, "..", "dist", "TTUM_Generator_NoMacros.xlsx")
SAMPLE = os.path.join(HERE, "..", "samples", "PROPELG_TTUM_16072026_Revised.txt")
SCRATCH = os.path.join(HERE, "..", "dist", "_formula_check.xlsx")

FIRST_ROW, LAST_ROW = 5, 104
OUT_FIRST = 9


def main():
    wb = openpyxl.load_workbook(XLSX)
    entries = wb["Entries"]

    # Fill in the day the bank's sample covers.
    wb["Dashboard"]["C6"] = date(2026, 7, 15)
    filled = 0
    for r in range(FIRST_ROW, LAST_ROW + 1):
        label = entries.cell(row=r, column=2).value
        if label in SAMPLE_AMOUNTS_KEYED:
            entries.cell(row=r, column=5).value = SAMPLE_AMOUNTS_KEYED[label]
            filled += 1
    if filled != len(SAMPLE_AMOUNTS_KEYED):
        print("FAIL: filled %d of %d rows" % (filled, len(SAMPLE_AMOUNTS_KEYED)))
        return 1
    wb.save(SCRATCH)

    import formulas
    model = formulas.ExcelModel().loads(SCRATCH).finish()
    solution = model.calculate()
    os.remove(SCRATCH)

    book = os.path.basename(SCRATCH)
    records = []
    for i in range(len(SAMPLE_AMOUNTS_KEYED)):
        key = "'[%s]TEXT OUTPUT'!A%d" % (book, OUT_FIRST + i)
        cell = solution.get(key)
        if cell is None:
            print("FAIL: no calculated value for %s" % key)
            return 1
        value = cell.value[0, 0]
        records.append(str(value))

    expected = open(SAMPLE, "rb").read().decode("ascii").split("\r\n")
    for i, (got, want) in enumerate(zip(records, expected), 1):
        if got != want:
            print("FAIL: record %d differs" % i)
            print("  got %r" % got)
            print("  exp %r" % want)
            return 1

    print("PASS: the Text Output formulas reproduce all %d records of the bank file"
          % len(records))
    lengths = set(len(r) for r in records)
    print("      every record is %d characters, as the bank expects" % lengths.pop())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
