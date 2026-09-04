"""Regression test: rebuild the known-good bank file from the standard row template."""
import sys, os
from datetime import date
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ttum_reference import build_file, RECORD_LEN
from standard_rows import STANDARD_ROWS

SAMPLE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                      "samples", "PROPELG_TTUM_16072026_Revised.txt")

def main():
    expected = open(SAMPLE, "rb").read()
    rows = [(r["account"], r["amount"], r["type"], r["narration"]) for r in STANDARD_ROWS]
    actual = build_file(rows, date(2026, 7, 15), trailing_newline=False)

    assert RECORD_LEN == 186, RECORD_LEN
    if actual == expected:
        print("PASS: regenerated %d bytes, byte-for-byte identical to the bank sample" % len(actual))
        return 0

    print("FAIL: %d bytes produced vs %d expected" % (len(actual), len(expected)))
    a, e = actual.split(b"\r\n"), expected.split(b"\r\n")
    for i, (x, y) in enumerate(zip(a, e), 1):
        if x != y:
            print("line %d differs:\n  got %r\n  exp %r" % (i, x, y))
    return 1

if __name__ == "__main__":
    raise SystemExit(main())
