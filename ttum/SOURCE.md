# Source guide

Everything needed to build, change and verify the utility. No hidden steps: the
shipped workbook in `dist/` is produced entirely by the scripts in `tools/` from
the VBA in `vba/`, and four tests check the result against real bank files.

## What is source and what is generated

| Path | |
|---|---|
| `vba/modTTUM.bas` | **Source.** The whole utility, ~1,290 lines of VBA. |
| `vba/ThisWorkbook.cls` | **Source.** Runs on open: restores buttons, fills the output folder. |
| `tools/standard_rows.py` | **Source.** The nine settlement lines: accounts, narrations, import keys. |
| `tools/build_workbook.py` | **Source.** Builds every sheet, the buttons, and the packaging. |
| `tools/vbaproject.py`, `cfb.py`, `msovba.py` | **Source.** Write the embedded VBA project from scratch. |
| `tools/ttum_reference.py` | **Source.** The record layout mirrored in Python, for testing. |
| `tools/test_*.py` | **Source.** The four tests. |
| `tools/build_prd_docx.js` | **Source.** Builds the Word PRD. |
| `samples/` | **Fixtures.** The bank's spec, its accepted file, a real settlement file. |
| `dist/` | **Generated.** Do not edit by hand; rebuild instead. |
| `PRD.html` | The requirements document. |

## Build

```bash
pip install openpyxl
cd tools
python3 build_workbook.py
```

Produces `dist/TTUM_Generator.xlsm`, `dist/TTUM_Generator_NoMacros.xlsx` and a
copy of `modTTUM.bas` for manual import.

## Test

```bash
pip install openpyxl oletools formulas
cd tools
python3 test_reference.py   # the layout rebuilds the bank's file byte for byte
python3 test_workbook.py    # the shipped rows and settings reproduce it too
python3 test_import.py      # a real settlement file imports and generates cleanly
python3 test_formula.py     # the macro-free sheets, run through a formula engine
```

All four must pass before shipping a build. They run without Excel.

## How the VBA gets into the workbook

No tool can create an `.xlsm`'s VBA project from an `.xlsx`, so `tools/` writes
one directly:

- `msovba.py` — the MS-OVBA compression every stream uses.
- `cfb.py` — a Compound File Binary writer (the container format).
- `vbaproject.py` — the `dir`, `PROJECT`, `PROJECTwm` and module streams.

`build_workbook.py` then repackages the `.xlsx` as macro-enabled: swaps the
content type, adds the `vbaProject.bin` relationship, sets `codeName` on the
workbook and every sheet, and injects the button shapes as a drawing part.

The compressor is cross-checked against `oletools`' independent decompressor,
and the finished project is read back with `oletools`' parser.

## Changing things without touching code

Most changes are sheet data, not code:

- **Settlement lines** — add rows on the Entries sheet (100 available): account,
  Dr/Cr, narration template, import match text. To change the shipped defaults,
  edit `tools/standard_rows.py` and rebuild.
- **File naming, folders, the paise convention, the balance rule** — Config sheet.
- **Narration wording** — the template column, with `{DDMMYY}` style tokens.

Code changes are needed only if the bank changes the record layout. That lives in
one place per language: the `LEN_*` constants at the top of `modTTUM.bas`, and
the same constants in `tools/ttum_reference.py`. Change both, then run the tests.

## Known constraints

- Windows Excel 2010+. `FileDialog`, `Dir$`, `FileDateTime` and `Shell` are used.
- The embedded VBA project has not been confirmed to load in Microsoft Excel;
  `dist/` therefore also ships a macro-free workbook and the `.bas` for manual
  import. See the Setup sheet.
- Amounts are converted through VBA `Currency` (exact to four decimals). Values
  above about 92 trillion paise would overflow; settlement volumes are far below.
