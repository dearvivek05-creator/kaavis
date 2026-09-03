"""Reference implementation of the TTUM fixed-width record layout.

This mirrors, line for line, what modTTUM.bas does inside the workbook. It exists so the
layout can be regression-tested against the known-good bank file without opening Excel.

Layout (1-based columns, total 186 chars per record):
    1-14    Account number      left justified, space padded
    15-22   Value date          DDMMYYYY
    23-45   Amount              23 digits, zero padded, last 2 digits are paise
    46      Transaction type    D or C
    47-55   Value date          9 chars: "0" + DDMMYYYY
    56-81   Filler              26 spaces
    82-116  Narration           left justified, space padded
    117-186 Filler              70 spaces
"""

from datetime import date

ACCOUNT_LEN = 14
DATE_LEN = 8
AMOUNT_LEN = 23
TRAN_TYPE_LEN = 1
DATE2_LEN = 9
FILLER1_LEN = 26
NARRATION_LEN = 35
FILLER2_LEN = 70
RECORD_LEN = (ACCOUNT_LEN + DATE_LEN + AMOUNT_LEN + TRAN_TYPE_LEN +
              DATE2_LEN + FILLER1_LEN + NARRATION_LEN + FILLER2_LEN)


def substitute_tokens(template: str, value_date: date) -> str:
    """Replace date tokens in a narration template. Case-insensitive, longest token first."""
    out = template
    for token, value in (
        ("{DDMMYYYY}", value_date.strftime("%d%m%Y")),
        ("{DDMMYY}", value_date.strftime("%d%m%y")),
        ("{YYYY}", value_date.strftime("%Y")),
        ("{DD}", value_date.strftime("%d")),
        ("{MM}", value_date.strftime("%m")),
        ("{YY}", value_date.strftime("%y")),
    ):
        lowered = out.lower()
        needle = token.lower()
        while needle in lowered:
            i = lowered.index(needle)
            out = out[:i] + value + out[i + len(token):]
            lowered = out.lower()
    return out


def amount_field(amount_rupees) -> str:
    """Rupees (may carry paise) -> 23 char zero padded paise string."""
    paise = int(round(float(amount_rupees) * 100))
    if paise < 0:
        raise ValueError("amount must not be negative")
    text = str(paise)
    if len(text) > AMOUNT_LEN:
        raise ValueError("amount too large for the %d char field" % AMOUNT_LEN)
    return text.rjust(AMOUNT_LEN, "0")


def build_record(account_no, value_date: date, amount_rupees, tran_type, narration) -> str:
    account = str(account_no).strip()
    if len(account) > ACCOUNT_LEN:
        raise ValueError("account number longer than %d chars: %s" % (ACCOUNT_LEN, account))

    tran = str(tran_type).strip().upper()
    if tran not in ("D", "C"):
        raise ValueError("transaction type must be D or C, got %r" % tran_type)

    text = substitute_tokens(str(narration), value_date)
    if len(text) > NARRATION_LEN:
        text = text[:NARRATION_LEN]

    ddmmyyyy = value_date.strftime("%d%m%Y")
    return (account.ljust(ACCOUNT_LEN)
            + ddmmyyyy
            + amount_field(amount_rupees)
            + tran
            + "0" + ddmmyyyy
            + " " * FILLER1_LEN
            + text.ljust(NARRATION_LEN)
            + " " * FILLER2_LEN)


def build_file(rows, value_date: date, trailing_newline=False) -> bytes:
    """rows: iterable of (account_no, amount_rupees, tran_type, narration_template)."""
    records = [build_record(a, value_date, amt, t, n) for a, amt, t, n in rows]
    for r in records:
        assert len(r) == RECORD_LEN, len(r)
    blob = "\r\n".join(records)
    if trailing_newline:
        blob += "\r\n"
    return blob.encode("ascii")
