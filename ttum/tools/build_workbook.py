#!/usr/bin/env python3
"""Builds TTUM_Generator.xlsm - the workbook plus its embedded VBA project.

Run:  python3 build_workbook.py
Out:  ../dist/TTUM_Generator.xlsm
      ../dist/TTUM_Generator_NoMacros.xlsx  (same sheets, for the manual-import route)
"""

import os
import re
import shutil
import sys
import zipfile

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import vbaproject  # noqa: E402
from standard_rows import STANDARD_ROWS  # noqa: E402

DIST = os.path.join(HERE, "..", "dist")
VBA = os.path.join(HERE, "..", "vba")

FIRST_ROW, LAST_ROW = 5, 104

# ---------------------------------------------------------------- appearance
INK = "1F3355"
MUTED = "6B7280"
NAVY = PatternFill("solid", fgColor="1F3355")
BAND = PatternFill("solid", fgColor="EEF2F7")
INPUT_FILL = PatternFill("solid", fgColor="FFF7DB")
HEAD_FILL = PatternFill("solid", fgColor="D9E2EC")
WHITE = PatternFill("solid", fgColor="FFFFFF")

TITLE = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
SUBTITLE = Font(name="Calibri", size=10, italic=True, color="D6E0EA")
SECTION = Font(name="Calibri", size=11, bold=True, color=INK)
LABEL = Font(name="Calibri", size=11, color="1F2937")
NOTE = Font(name="Calibri", size=9, italic=True, color=MUTED)
VALUE = Font(name="Calibri", size=11, bold=True, color="1F2937")
HEADER = Font(name="Calibri", size=10, bold=True, color=INK)
MONO = Font(name="Consolas", size=10)

THIN = Side(style="thin", color="BFCBD9")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def banner(ws, title, subtitle, last_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    for row in (1, 2):
        for col in range(1, last_col + 1):
            ws.cell(row=row, column=col).fill = NAVY
    c = ws.cell(row=1, column=1)
    c.value = "   " + title
    c.font = TITLE
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30
    c = ws.cell(row=2, column=1)
    c.value = "   " + subtitle
    c.font = SUBTITLE
    ws.row_dimensions[2].height = 16


def input_cell(ws, addr, value=None, fmt=None):
    c = ws[addr]
    if value is not None:
        c.value = value
    c.fill = INPUT_FILL
    c.font = VALUE
    c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


# ------------------------------------------------------------------ sheets

def build_dashboard(ws):
    banner(ws, "TTUM GENERATOR",
           "Enter the day's amounts on the Entries sheet, check the date, then click Generate.", 8)
    widths = {"A": 2.5, "B": 34, "C": 30, "D": 2.5, "E": 2.5, "F": 2.5, "G": 2.5, "H": 26}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    def section(row, text):
        ws.cell(row=row, column=2, value=text).font = SECTION
        ws.cell(row=row, column=2).fill = BAND
        ws.cell(row=row, column=3).fill = BAND

    def label(row, text, note=None):
        ws.cell(row=row, column=2, value=text).font = LABEL
        if note:
            ws.cell(row=row + 1, column=2, value=note).font = NOTE

    section(4, "STEP 1   Set the date")
    label(5, "Value date  (goes into every record)")
    input_cell(ws, "C5", "=TODAY()", "dd-mmm-yyyy")
    label(6, "File-name date  (Config: offset in days)")
    input_cell(ws, "C6", "=ttValueDate+ttFileDateOffset", "dd-mmm-yyyy")
    ws["B7"] = "Today's date is filled in automatically. Type over it to build the file for another day."
    ws["B7"].font = NOTE
    ws.merge_cells("B7:C7")

    section(9, "STEP 2   Enter the amounts on the Entries sheet")
    label(10, "Rows included")
    ws["C10"] = '=COUNTIF(Entries!$A$%d:$A$%d,"Yes")' % (FIRST_ROW, LAST_ROW)
    label(11, "Total debit")
    ws["C11"] = "=Entries!$I$1"
    label(12, "Total credit")
    ws["C12"] = "=Entries!$I$2"
    label(13, "Difference  (must be 0.00)")
    ws["C13"] = "=Entries!$I$3"
    for addr in ("C11", "C12", "C13"):
        ws[addr].number_format = "#,##0.00"
    for addr in ("C10", "C11", "C12", "C13"):
        ws[addr].font = VALUE
        ws[addr].border = BOX
    ws.conditional_formatting.add("C13", CellIsRule(
        operator="notEqual", formula=["0"],
        fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add("C13", CellIsRule(
        operator="equal", formula=["0"],
        fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))

    section(15, "STEP 3   Generate")
    label(16, "Output folder")
    input_cell(ws, "C16")
    ws["B17"] = "Leave blank to write into a TTUM_Output folder next to this workbook."
    ws["B17"].font = NOTE
    ws.merge_cells("B17:C17")
    label(18, "File name")
    ws["C18"] = ('=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(ttFileNamePattern,'
                 '"{DDMMYYYY}",TEXT(ttFileDate,"ddmmyyyy")),'
                 '"{DDMMYY}",TEXT(ttFileDate,"ddmmyy")),'
                 '"{YYYY}",TEXT(ttFileDate,"yyyy"))')
    ws["C18"].font = VALUE
    ws["C18"].border = BOX

    ws.cell(row=20, column=2, value="Status").font = SECTION
    ws.merge_cells("B21:H21")
    ws["B21"] = "Ready."
    ws["B21"].font = Font(name="Calibri", size=11, bold=True, color="006E3C")
    ws["B21"].alignment = Alignment(vertical="center")
    ws.row_dimensions[21].height = 22

    ws["B23"] = ("If the buttons on the right are ever missing, press Alt+F8 and run TTUM_Setup. "
                 "Every button is also available from that same Alt+F8 list.")
    ws["B23"].font = NOTE
    ws.merge_cells("B23:H23")

    ws.sheet_view.showGridLines = False


def build_entries(ws):
    banner(ws, "DAILY TTUM ENTRIES",
           "Type today's amount and Dr/Cr against each line. Blank rows are ignored.", 6)
    for col, w in {"A": 10, "B": 30, "C": 18, "D": 9, "E": 16, "F": 44,
                   "G": 2.5, "H": 16, "I": 14}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A3:F3")
    ws["A3"] = ('="Total debit  "&TEXT($I$1,"#,##0.00")&"      Total credit  "'
                '&TEXT($I$2,"#,##0.00")&"      Difference  "&TEXT($I$3,"#,##0.00")'
                '&IF($I$3=0,"    (balanced)","    <<< THESE MUST MATCH")')
    ws["A3"].font = Font(name="Calibri", size=11, bold=True)
    ws["A3"].alignment = Alignment(vertical="center")
    ws.row_dimensions[3].height = 20
    ws.conditional_formatting.add("A3:F3", FormulaRule(
        formula=["$I$3<>0"], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add("A3:F3", FormulaRule(
        formula=["$I$3=0"], fill=PatternFill("solid", fgColor="C6EFCE"),
        font=Font(color="006100", bold=True)))

    ws["H1"] = "Total debit"
    ws["I1"] = '=SUMIFS($E$%d:$E$%d,$D$%d:$D$%d,"D",$A$%d:$A$%d,"Yes")' % (
        FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW)
    ws["H2"] = "Total credit"
    ws["I2"] = '=SUMIFS($E$%d:$E$%d,$D$%d:$D$%d,"C",$A$%d:$A$%d,"Yes")' % (
        FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW)
    ws["H3"] = "Difference"
    ws["I3"] = "=$I$1-$I$2"
    for row in (1, 2, 3):
        ws.cell(row=row, column=8).font = NOTE
        ws.cell(row=row, column=9).number_format = "#,##0.00"
        ws.cell(row=row, column=9).font = Font(size=10, bold=True)

    headers = ["Include", "Description", "Account Number", "Dr/Cr",
               "Amount (INR)", "Narration template"]
    for i, text in enumerate(headers, start=1):
        c = ws.cell(row=FIRST_ROW - 1, column=i, value=text)
        c.font = HEADER
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[FIRST_ROW - 1].height = 28

    for offset, row in enumerate(STANDARD_ROWS):
        r = FIRST_ROW + offset
        ws.cell(row=r, column=1, value="Yes")
        ws.cell(row=r, column=2, value=row["label"])
        ws.cell(row=r, column=3, value=row["account"])
        ws.cell(row=r, column=4, value=row["type"])
        ws.cell(row=r, column=6, value=row["narration"])

    for r in range(FIRST_ROW, LAST_ROW + 1):
        for col in range(1, 7):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            c.fill = WHITE
            c.font = Font(name="Calibri", size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3).number_format = "@"
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="left")
        ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5).number_format = "#,##0.00"
        ws.cell(row=r, column=5).fill = INPUT_FILL

    span = "%d:%d" % (FIRST_ROW, LAST_ROW)
    inc = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True,
                         showErrorMessage=True, errorTitle="Include?",
                         error="Choose Yes to put this row in today's file, or No to leave it out.")
    drcr = DataValidation(type="list", formula1='"D,C"', allow_blank=True,
                          showErrorMessage=True, errorTitle="Dr/Cr",
                          error="Enter D for a debit or C for a credit.")
    amt = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                         allow_blank=True, showErrorMessage=True, errorTitle="Amount",
                         error="Enter the amount in rupees, for example 4740.50.")
    for dv, col in ((inc, "A"), (drcr, "D"), (amt, "E")):
        ws.add_data_validation(dv)
        dv.add("%s%d:%s%d" % (col, FIRST_ROW, col, LAST_ROW))

    ws.freeze_panes = "A%d" % FIRST_ROW
    ws.sheet_view.showGridLines = False

    note_row = LAST_ROW + 2
    ws.cell(row=note_row, column=1,
            value="Narration tokens: {DDMMYY} {DDMMYYYY} {DD} {MM} {YY} {YYYY} are replaced with "
                  "the value date. Narration is capped at 35 characters after substitution.").font = NOTE
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=6)


def build_config(ws):
    banner(ws, "CONFIG", "Settings that rarely change. The daily inputs live on the Dashboard.", 4)
    for col, w in {"A": 2.5, "B": 36, "C": 32, "D": 62}.items():
        ws.column_dimensions[col].width = w

    rows = [
        ("File name pattern", "PROPELG_TTUM_{DDMMYYYY}.txt",
         "Tokens are replaced with the file-name date shown on the Dashboard."),
        ("File-name date offset (days)", 1,
         "The bank sample is dated one day after the value date it carries, so this is 1. "
         "Set 0 to use the value date itself."),
        ("Block generation when out of balance", "Yes",
         "Yes stops the file when total debit does not equal total credit. "
         "No asks first and lets you continue."),
        ("Write a line break after the last record", "No",
         "No matches the bank's sample file, which ends immediately after the last record."),
    ]
    ws.cell(row=4, column=2, value="Setting").font = HEADER
    ws.cell(row=4, column=3, value="Value").font = HEADER
    ws.cell(row=4, column=4, value="What it does").font = HEADER
    for col in (2, 3, 4):
        ws.cell(row=4, column=col).fill = HEAD_FILL
        ws.cell(row=4, column=col).border = BOX

    for i, (name, value, note) in enumerate(rows):
        r = 5 + i
        ws.cell(row=r, column=2, value=name).font = LABEL
        c = ws.cell(row=r, column=3, value=value)
        c.fill = INPUT_FILL
        c.font = VALUE
        c.border = BOX
        n = ws.cell(row=r, column=4, value=note)
        n.font = NOTE
        n.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=2).border = BOX
        ws.row_dimensions[r].height = 30

    yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(yn)
    yn.add("C7:C8")
    ws.sheet_view.showGridLines = False


def build_log(ws):
    banner(ws, "GENERATION LOG", "One row is added every time a file is written.", 10)
    headers = ["Generated at", "Value date", "File date", "File name", "Records",
               "Total debit", "Total credit", "Balance", "User", "Full path"]
    widths = [19, 14, 14, 30, 9, 14, 14, 16, 18, 60]
    for i, (text, w) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=3, column=i, value=text)
        c.font = HEADER
        c.fill = HEAD_FILL
        c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False


def build_layout(ws):
    banner(ws, "RECORD LAYOUT", "Reference only - the macro builds every record to this layout.", 6)
    for col, w in {"A": 2.5, "B": 8, "C": 14, "D": 24, "E": 10, "F": 60}.items():
        ws.column_dimensions[col].width = w

    headers = ["#", "Columns", "Field", "Length", "Content"]
    for i, text in enumerate(headers, start=2):
        c = ws.cell(row=4, column=i, value=text)
        c.font = HEADER
        c.fill = HEAD_FILL
        c.border = BOX

    fields = [
        ("1-14", "Account number", 14, "Left justified, padded on the right with spaces."),
        ("15-22", "Value date", 8, "DDMMYYYY."),
        ("23-45", "Amount", 23, "Whole paise, padded on the left with zeros. "
                                "20000 means 200.00."),
        ("46", "Transaction type", 1, "D for debit, C for credit."),
        ("47-55", "Value date", 9, "The same date again as 0 + DDMMYYYY."),
        ("56-81", "Filler", 26, "Spaces."),
        ("82-116", "Narration", 35, "Left justified, padded on the right with spaces."),
        ("117-186", "Filler", 70, "Spaces."),
    ]
    for i, (cols, name, length, content) in enumerate(fields, start=1):
        r = 4 + i
        ws.cell(row=r, column=2, value=i)
        ws.cell(row=r, column=3, value=cols)
        ws.cell(row=r, column=4, value=name)
        ws.cell(row=r, column=5, value=length)
        ws.cell(row=r, column=6, value=content)
        for col in range(2, 7):
            ws.cell(row=r, column=col).border = BOX
            ws.cell(row=r, column=col).font = Font(size=10)
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=13, column=4, value="Total").font = HEADER
    ws.cell(row=13, column=5, value=186).font = HEADER
    ws.cell(row=13, column=6, value="Records are separated by CR LF.").font = NOTE

    ws.cell(row=15, column=2,
            value="Worked example - the bank's own file for value date 15-07-2026").font = SECTION
    ws.cell(row=16, column=2,
            value="Amounts below are that day's figures; each line is exactly 186 characters.").font = NOTE

    sample = os.path.join(HERE, "..", "samples", "PROPELG_TTUM_16072026_Revised.txt")
    with open(sample, "rb") as handle:
        lines = handle.read().decode("ascii").split("\r\n")
    for i, line in enumerate(lines):
        c = ws.cell(row=17 + i, column=2, value=line)
        c.font = MONO
    ws.sheet_view.showGridLines = False


# ------------------------------------------------------------------- build

SHEETS = [
    ("Dashboard", "shDashboard", build_dashboard),
    ("Entries", "shEntries", build_entries),
    ("Config", "shConfig", build_config),
    ("Log", "shLog", build_log),
    ("Layout", "shLayout", build_layout),
]

NAMES = {
    "ttValueDate": "Dashboard!$C$5",
    "ttFileDate": "Dashboard!$C$6",
    "ttOutputFolder": "Dashboard!$C$16",
    "ttStatus": "Dashboard!$B$21",
    "ttFileNamePattern": "Config!$C$5",
    "ttFileDateOffset": "Config!$C$6",
    "ttEnforceBalance": "Config!$C$7",
    "ttTrailingNewline": "Config!$C$8",
}


def build_workbook_xlsx(path):
    wb = Workbook()
    wb.remove(wb.active)
    wb.code_name = "ThisWorkbook"
    for title, code_name, builder in SHEETS:
        ws = wb.create_sheet(title)
        ws.sheet_properties.codeName = code_name
        builder(ws)
    for name, ref in NAMES.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))
    wb.active = 0
    wb.save(path)


def read_source(name):
    with open(os.path.join(VBA, name), "r", encoding="utf-8") as handle:
        return handle.read()


def build_vba():
    """The VBA project: one document module per sheet, plus ThisWorkbook and modTTUM."""
    body = read_source("ThisWorkbook.cls").split("Attribute VB_Exposed = True", 1)[1]
    modules = [vbaproject.Module(
        "ThisWorkbook",
        vbaproject.document_header("ThisWorkbook", True) + body,
        vbaproject.MODULE_DOCUMENT)]
    for title, code_name, _ in SHEETS:
        modules.append(vbaproject.Module(
            code_name, vbaproject.document_header(code_name, False),
            vbaproject.MODULE_DOCUMENT))
    modules.append(vbaproject.Module("modTTUM", read_source("modTTUM.bas")))
    return vbaproject.build("VBAProject", modules)


def to_macro_enabled(xlsx_path, xlsm_path, vba_blob):
    """Repackages the .xlsx as a macro-enabled workbook carrying vbaProject.bin."""
    src = zipfile.ZipFile(xlsx_path)
    if os.path.exists(xlsm_path):
        os.remove(xlsm_path)
    out = zipfile.ZipFile(xlsm_path, "w", zipfile.ZIP_DEFLATED)

    for item in src.infolist():
        data = src.read(item.filename)

        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = text.replace(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            insert = ('<Default Extension="bin" '
                      'ContentType="application/vnd.ms-office.vbaProject"/>')
            text = text.replace("</Types>", insert + "</Types>")
            data = text.encode("utf-8")

        elif item.filename == "xl/workbook.xml":
            text = data.decode("utf-8")
            if "codeName=" not in text:
                if re.search(r"<workbookPr[^>]*/>", text):
                    text = re.sub(r"<workbookPr([^>]*)/>",
                                  r'<workbookPr\1 codeName="ThisWorkbook"/>', text, count=1)
                else:
                    text = text.replace("<sheets>",
                                        '<workbookPr codeName="ThisWorkbook"/><sheets>', 1)
            data = text.encode("utf-8")

        elif item.filename == "xl/_rels/workbook.xml.rels":
            text = data.decode("utf-8")
            rel = ('<Relationship Id="rIdVBA" '
                   'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                   'Target="vbaProject.bin"/>')
            text = text.replace("</Relationships>", rel + "</Relationships>")
            data = text.encode("utf-8")

        out.writestr(item, data)

    out.writestr("xl/vbaProject.bin", vba_blob)
    out.close()
    src.close()


def main():
    if not os.path.isdir(DIST):
        os.makedirs(DIST)
    xlsx = os.path.join(DIST, "TTUM_Generator_NoMacros.xlsx")
    xlsm = os.path.join(DIST, "TTUM_Generator.xlsm")

    build_workbook_xlsx(xlsx)
    to_macro_enabled(xlsx, xlsm, build_vba())
    shutil.copyfile(os.path.join(VBA, "modTTUM.bas"), os.path.join(DIST, "modTTUM.bas"))

    for path in (xlsm, xlsx):
        print("%-42s %8d bytes" % (os.path.basename(path), os.path.getsize(path)))


if __name__ == "__main__":
    main()
