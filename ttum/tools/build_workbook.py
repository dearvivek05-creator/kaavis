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
    widths = {"A": 2.5, "B": 34, "C": 46, "D": 2.5, "E": 2.5, "F": 2.5, "G": 2.5, "H": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("B3:H3")
    ws["B3"] = ("If the buttons on the right do nothing, Excel is blocking macros:  "
                "click Enable Editing, then Enable Content, at the top of the window.")
    ws["B3"].fill = PatternFill("solid", fgColor="FFF3CD")
    ws["B3"].font = Font(name="Calibri", size=10, bold=True, color="8A6100")
    ws["B3"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[3].height = 20

    def section(row, text):
        ws.cell(row=row, column=2, value=text).font = SECTION
        for col in (2, 3):
            ws.cell(row=row, column=col).fill = BAND

    def label(row, text):
        ws.cell(row=row, column=2, value=text).font = LABEL

    def note(row, text):
        c = ws.cell(row=row, column=2, value=text)
        c.font = NOTE
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)

    section(5, "STEP 1   Set the date")
    label(6, "Value date  (goes into every record)")
    input_cell(ws, "C6", "=TODAY()", "dd-mmm-yyyy")
    label(7, "File-name date  (offset is set on Config)")
    input_cell(ws, "C7", "=ttValueDate+ttFileDateOffset", "dd-mmm-yyyy")
    note(8, "Today's date is filled in automatically. Type over it to build the file for another day.")

    section(10, "STEP 2   Enter the amounts on the Entries sheet")
    label(11, "Rows included")
    ws["C11"] = '=COUNTIF(Entries!$A$%d:$A$%d,"Yes")' % (FIRST_ROW, LAST_ROW)
    label(12, "Total debit")
    ws["C12"] = "=Entries!$J$1"
    label(13, "Total credit")
    ws["C13"] = "=Entries!$J$2"
    label(14, "Difference  (must be 0.00)")
    ws["C14"] = "=Entries!$J$3"
    for addr in ("C12", "C13", "C14"):
        ws[addr].number_format = "#,##0.00"
    for addr in ("C11", "C12", "C13", "C14"):
        ws[addr].font = VALUE
        ws[addr].border = BOX
    ws.conditional_formatting.add("C14", CellIsRule(
        operator="notEqual", formula=["0"],
        fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add("C14", CellIsRule(
        operator="equal", formula=["0"],
        fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))

    section(16, "STEP 3   Choose where the file goes, then generate")
    label(17, "Output folder")
    input_cell(ws, "C17")
    ws["C17"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[17].height = 20
    note(18, "Type or paste a full folder path here, for example  D:\\TTUM\\Daily  -  or click "
             "Choose Output Folder. Left blank, files go to a TTUM_Output folder beside this workbook.")
    ws.row_dimensions[18].height = 26
    ws["B18"].alignment = Alignment(wrap_text=True, vertical="top")
    label(19, "File name")
    ws["C19"] = ('=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(ttFileNamePattern,'
                 '"{DDMMYYYY}",TEXT(ttFileDate,"ddmmyyyy")),'
                 '"{DDMMYY}",TEXT(ttFileDate,"ddmmyy")),'
                 '"{YYYY}",TEXT(ttFileDate,"yyyy"))')
    ws["C19"].font = VALUE
    ws["C19"].border = BOX

    ws.cell(row=21, column=2, value="Status").font = SECTION
    ws.merge_cells("B22:H22")
    ws["B22"] = "Ready."
    ws["B22"].font = Font(name="Calibri", size=11, bold=True, color="006E3C")
    ws["B22"].alignment = Alignment(vertical="center")
    ws.row_dimensions[22].height = 22

    ws["B24"] = ("Every button is also on the Alt+F8 macro list. If the buttons are missing "
                 "altogether, run TTUM_Setup from there.")
    ws["B24"].font = NOTE
    ws.merge_cells("B24:H24")

    ws.sheet_view.showGridLines = False


def build_entries(ws):
    banner(ws, "DAILY TTUM ENTRIES",
           "Type today's amount and Dr/Cr against each line. Blank rows are ignored.", 7)
    for col, w in {"A": 10, "B": 28, "C": 18, "D": 8, "E": 20, "F": 42,
                   "G": 16, "H": 2.5, "I": 14, "J": 14}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("A3:G3")
    ws["A3"] = ('="Total debit  "&TEXT($J$1,"#,##0.00")&"      Total credit  "'
                '&TEXT($J$2,"#,##0.00")&"      Difference  "&TEXT($J$3,"#,##0.00")'
                '&IF($J$3=0,"    (balanced)","    <<< THESE MUST MATCH")')
    ws["A3"].font = Font(name="Calibri", size=11, bold=True)
    ws["A3"].alignment = Alignment(vertical="center")
    ws.row_dimensions[3].height = 20
    ws.conditional_formatting.add("A3:G3", FormulaRule(
        formula=["$J$3<>0"], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add("A3:G3", FormulaRule(
        formula=["$J$3=0"], fill=PatternFill("solid", fgColor="C6EFCE"),
        font=Font(color="006100", bold=True)))

    divisor = 'IF(ttAmountUnit="Paise",100,1)'
    ws["I1"] = "Total debit"
    ws["J1"] = '=SUMIFS($E$%d:$E$%d,$D$%d:$D$%d,"D",$A$%d:$A$%d,"Yes")/%s' % (
        FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, divisor)
    ws["I2"] = "Total credit"
    ws["J2"] = '=SUMIFS($E$%d:$E$%d,$D$%d:$D$%d,"C",$A$%d:$A$%d,"Yes")/%s' % (
        FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, FIRST_ROW, LAST_ROW, divisor)
    ws["I3"] = "Difference"
    ws["J3"] = "=$J$1-$J$2"
    for row in (1, 2, 3):
        ws.cell(row=row, column=9).font = NOTE
        ws.cell(row=row, column=10).number_format = "#,##0.00"
        ws.cell(row=row, column=10).font = Font(size=10, bold=True)

    headers = ["Include", "Description", "Account Number", "Dr/Cr",
               "Amount as keyed", "Narration template", "Reads as (INR)"]
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
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            c.fill = WHITE
            c.font = Font(name="Calibri", size=10)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3).number_format = "@"
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="left")
        ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5).number_format = "#,##0"
        ws.cell(row=r, column=5).fill = INPUT_FILL
        # A live read-back of the keyed figure, so the decimal point is never in doubt.
        ws.cell(row=r, column=7,
                value='=IF($E%d="","",$E%d/%s)' % (r, r, divisor))
        ws.cell(row=r, column=7).number_format = "#,##0.00"
        ws.cell(row=r, column=7).font = Font(name="Calibri", size=10, bold=True, color="1F3355")

    inc = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True,
                         showErrorMessage=True, errorTitle="Include?",
                         error="Choose Yes to put this row in today's file, or No to leave it out.")
    drcr = DataValidation(type="list", formula1='"D,C"', allow_blank=True,
                          showErrorMessage=True, errorTitle="Dr/Cr",
                          error="Enter D for a debit or C for a credit.")
    amt = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                         allow_blank=True, showErrorMessage=True, errorTitle="Amount",
                         error="Enter the amount as it appears on the settlement report.")
    for dv, col in ((inc, "A"), (drcr, "D"), (amt, "E")):
        ws.add_data_validation(dv)
        dv.add("%s%d:%s%d" % (col, FIRST_ROW, col, LAST_ROW))

    ws.freeze_panes = "A%d" % FIRST_ROW
    ws.sheet_view.showGridLines = False

    note_row = LAST_ROW + 2
    ws.cell(row=note_row, column=1,
            value="Amount as keyed: with Config set to Paise (the default), type the figure exactly as "
                  "the settlement report shows it - the last two digits are the paise, so 47400 is "
                  "474.00. The Reads as (INR) column shows how each figure will be understood. "
                  "Narration tokens {DDMMYY} {DDMMYYYY} {DD} {MM} {YY} {YYYY} are replaced with the "
                  "value date; the narration must fit 35 characters after substitution.").font = NOTE
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row + 1, end_column=7)
    ws.cell(row=note_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")


def build_config(ws):
    banner(ws, "CONFIG", "Settings that rarely change. The daily inputs live on the Dashboard.", 4)
    for col, w in {"A": 2.5, "B": 36, "C": 32, "D": 62}.items():
        ws.column_dimensions[col].width = w

    rows = [
        ("Amount column is entered as", "Paise",
         "Paise: the Amount column is the figure exactly as the settlement report shows it, "
         "with the last two digits being the paise, so 47400 means 474.00. "
         "Rupees: type 474.00 instead. The Reads as (INR) column on Entries shows the effect."),
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
        ws.row_dimensions[r].height = 44 if i == 0 else 30

    unit = DataValidation(type="list", formula1='"Paise,Rupees"', allow_blank=False)
    ws.add_data_validation(unit)
    unit.add("C5")
    yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    ws.add_data_validation(yn)
    yn.add("C8:C9")
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


# --------------------------------------------------------------- buttons

EMU_PER_POINT = 12700

# label, macro, and whether it is the primary action
BUTTONS = [
    ("Generate TTUM File", "GenerateTTUM", True),
    ("Preview Records", "PreviewTTUM", False),
    ("Validate Entries", "ValidateEntries", False),
    ("Reset Date to Today", "ResetDateToToday", False),
    ("Clear Amounts", "ClearAmounts", False),
    ("Choose Output Folder...", "BrowseOutputFolder", False),
    ("Open Output Folder", "OpenOutputFolder", False),
]

BUTTON_COL = 7          # column H, zero based
BUTTON_ROW = 4          # row 5, zero based
BUTTON_WIDTH = 160      # points
BUTTON_HEIGHT = 30
BUTTON_PITCH = 38


def _button_shape(index, label, macro, primary):
    fill = "1F3355" if primary else "F1F5FA"
    line = "16283F" if primary else "9FB2C6"
    text = "FFFFFF" if primary else "1F3355"
    return (
        '<xdr:oneCellAnchor>'
        '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
        '<xdr:ext cx="%d" cy="%d"/>'
        '<xdr:sp macro="modTTUM.%s" textlink="">'
        '<xdr:nvSpPr><xdr:cNvPr id="%d" name="btnTT%d"/><xdr:cNvSpPr/></xdr:nvSpPr>'
        '<xdr:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
        '</xdr:spPr>'
        '<xdr:txBody>'
        '<a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/>'
        '<a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/>'
        '<a:r><a:rPr lang="en-US" sz="1000" b="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="Calibri"/></a:rPr><a:t>%s</a:t></a:r></a:p>'
        '</xdr:txBody>'
        '</xdr:sp>'
        '<xdr:clientData/>'
        '</xdr:oneCellAnchor>'
    ) % (BUTTON_COL, BUTTON_ROW, index * BUTTON_PITCH * EMU_PER_POINT,
         BUTTON_WIDTH * EMU_PER_POINT, BUTTON_HEIGHT * EMU_PER_POINT,
         macro, index + 2, index, fill, line, 1 if primary else 0, text, label)


def build_drawing_xml():
    """The Dashboard buttons, stored in the file rather than drawn at run time.

    Shapes carrying a `macro` attribute are ordinary drawings, so they are present
    and visible even before macros are enabled - which is the difference between
    a workbook that looks broken and one that just needs Enable Content.
    """
    shapes = "".join(_button_shape(i, label, macro, primary)
                     for i, (label, macro, primary) in enumerate(BUTTONS))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
            'spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            + shapes + '</xdr:wsDr>').encode("utf-8")


def attach_drawing(name, data, dashboard_part):
    """Wires the drawing into the Dashboard worksheet part."""
    if name == dashboard_part:
        text = data.decode("utf-8")
        if "xmlns:r=" not in text:
            text = text.replace(
                "<worksheet ",
                '<worksheet xmlns:r="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships" ', 1)
        text = text.replace("</worksheet>", '<drawing r:id="rIdDrawing"/></worksheet>')
        return text.encode("utf-8")
    return data


# ------------------------------------------------------------------- build

SHEETS = [
    ("Dashboard", "shDashboard", build_dashboard),
    ("Entries", "shEntries", build_entries),
    ("Config", "shConfig", build_config),
    ("Log", "shLog", build_log),
    ("Layout", "shLayout", build_layout),
]

NAMES = {
    "ttValueDate": "Dashboard!$C$6",
    "ttFileDate": "Dashboard!$C$7",
    "ttOutputFolder": "Dashboard!$C$17",
    "ttStatus": "Dashboard!$B$22",
    "ttAmountUnit": "Config!$C$5",
    "ttFileNamePattern": "Config!$C$6",
    "ttFileDateOffset": "Config!$C$7",
    "ttEnforceBalance": "Config!$C$8",
    "ttTrailingNewline": "Config!$C$9",
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

    dashboard_part = None
    for name in src.namelist():
        if name.startswith("xl/worksheets/sheet") and \
                b'codeName="shDashboard"' in src.read(name):
            dashboard_part = name
            break
    if dashboard_part is None:
        raise RuntimeError("could not find the Dashboard worksheet part")

    for item in src.infolist():
        data = attach_drawing(item.filename, src.read(item.filename), dashboard_part)

        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = text.replace(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
            insert = ('<Default Extension="bin" '
                      'ContentType="application/vnd.ms-office.vbaProject"/>'
                      '<Override PartName="/xl/drawings/drawing1.xml" ContentType='
                      '"application/vnd.openxmlformats-officedocument.drawing+xml"/>')
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
    out.writestr("xl/drawings/drawing1.xml", build_drawing_xml())
    out.writestr(
        "xl/worksheets/_rels/%s.rels" % os.path.basename(dashboard_part),
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
        '</Relationships>')
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
