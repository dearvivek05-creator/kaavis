# TTUM Daily Generator

An Excel macro workbook that turns the day's settlement amounts into the bank's
fixed-width TTUM upload file.

You type the amounts and the Dr/Cr flags, check the date, and click one button.
The macro lays every record out to the bank's specification, checks that debits
equal credits, and writes the `.txt` file.

**Start here:** `dist/TTUM_Generator.xlsm`

---

## Daily use

1. **Open the workbook** and click **Enable Content** when Excel asks about macros.
2. **Dashboard - check the date.** The value date is already today's. To build a
   file for a different day, just type that date over it. The file-name date and
   the file name update themselves.
3. **Get the day's amounts in** - either way works, and you can mix them:

   * **From the settlement file:** click **Import Latest Input File**. It reads the
     newest matching file from the input folder set on Config, shows you a summary -
     value date, record count, every matched amount, and whether debits and credits
     balance - and only puts the figures onto Entries once you say yes. **Choose
     Input File** does the same for a file you pick.
   * **By hand:** The nine standard settlement lines are already
   set up with their account numbers and narrations. Fill in the
   **Amount as keyed** column and adjust **Dr/Cr** if the day's flow differs.
     type the figure exactly as the settlement report shows it - **the last two
     digits are the paise**, so `47400` is 474.00 and `20000` is 200.00. The
     **Reads as (INR)** column shows how each figure will be understood.

   Either way, the banner at the top of Entries turns green when debits and
   credits match.
4. **Dashboard - set the output folder and click Generate TTUM File.** The
   **Output folder** box holds a full path such as `D:\TTUM\Daily`; type or paste
   one in, or click **Choose Output Folder**. It is filled in with a
   `TTUM_Output` folder beside the workbook the first time you open the file.
   The file is written there and a row is added to the Log sheet.

Use **Preview Records** first if you want to see the exact 186-character lines
without writing anything to disk. **Clear Amounts** blanks the amount column
ready for the next day.

---

## If the buttons do nothing

The buttons are drawings stored in the workbook, so they are always visible;
what they run is a macro, and Excel refuses to run macros in a file it does not
trust. The **Setup** sheet inside the workbook has this same list.

1. **Unblock the file - try this first.** Close the workbook. Right-click it in
   File Explorer, choose **Properties**, tick **Unblock** on the General tab,
   click OK, reopen. Files that arrive by email or download carry a mark that
   makes current Excel block their macros outright, with a red bar and **no**
   Enable Content button. Nothing else will help until this is cleared.
2. **Trust the folder.** File > Options > Trust Center > Trust Center Settings >
   Trusted Locations > Add new location, and add the folder you keep it in.
3. **Check the macro setting.** Same dialog, Macro Settings. It must be
   *Disable all macros with notification*, not *without notification*.
4. **See whether the macros arrived.** Press `Alt+F8`. If `GenerateTTUM` is
   listed, the macros are present and this is purely a trust problem - you can
   run it straight from that list meanwhile. If the list is empty, do step 5.
5. **Load the macros by hand, once.** `Alt+F11`, then **File > Import File**,
   pick `modTTUM.bas`, `Alt+Q`, save the workbook. The buttons are already wired
   to it, so they start working.

**If macros are blocked and cannot be unblocked at all**, use the **Text Output**
sheet. It builds the same 186-character records with worksheet formulas and no
macro: set the date and amounts as usual, select the filled cells in column A,
copy, paste into Notepad, and save with Encoding set to ANSI. Those formulas are
tested against the bank's own file, so the output is the same either way.

---

## The sheets

| Sheet | What it is for |
|---|---|
| **Dashboard** | The date, the output folder, the totals, and the buttons. |
| **Entries** | The daily grid: include, description, account number, Dr/Cr, amount, narration, and a read-back of the amount in rupees. 100 rows. |
| **Config** | Settings that rarely change: file-name pattern, date offset, balance rule. |
| **Log** | A row for every file generated: when, by whom, how many records, the totals. |
| **Import** | The last settlement file read, line by line, and what each line was matched to. |
| **Text Output** | The same records built by formula - the route that works with macros blocked. |
| **Layout** | The record specification, and the bank's own file as a worked example. |
| **Setup** | What to try when a button does nothing. |
| **Preview** | Created by the Preview button; safe to delete. |

### Narration tokens

A narration is written as a template. These placeholders are replaced with the
**value date** when the file is built:

`{DDMMYY}` `{DDMMYYYY}` `{DD}` `{MM}` `{YY}` `{YYYY}`

So `PROPELG VI settlement {DDMMYY}` becomes `PROPELG VI settlement 150726`.
A narration longer than 35 characters after substitution is rejected rather than
silently trimmed.

### Importing the settlement file

The settlement system's file is in the same 186-character layout as the TTUM file,
so the importer reads it with the same field positions.

**A line is matched to an Entries row by the row's Import match text**: the row
claims a line when that text appears anywhere in the line's narration. The keys
shipped on the sheet are the full wordings (`MC comm recd`, `MC GST comm recd`,
`MC Non GST comm recd` and so on) precisely so that no line can be claimed by two
rows. Change a key if the settlement system changes its wording.

**Only the amount and the Dr/Cr flag are taken from the file.** The account number
and the narration that go into the generated TTUM always come from the Entries
sheet. That matters for the net settlement line: the settlement file books it to
the nodal account, while the TTUM has to carry the Prp India account, and the
importer keeps them apart.

The summary appears before anything is written. If you say no, nothing on the
Entries sheet changes. If you say yes:

- matched rows get their amount and Dr/Cr, and are set to `Yes`;
- rows with an import key that the file did not mention are **cleared and set to
  `No`**, so a figure from a previous day can never be left behind;
- rows with no import key are left completely alone, so anything you maintain by
  hand survives an import;
- the value date is taken from the file's own records, unless you turn that off
  on Config.

Lines the importer could not place are listed in the summary and shown on the
Import sheet as `-- no match --`. They are left out, which puts the sheet out of
balance, which in turn stops generation - so an unmatched line cannot slip
through into a file.

**With macros blocked**, the Import sheet still works: open the settlement file in
Notepad, copy every line, and paste into cell `B10`. The columns beside it pull out
the account, date, amount, Dr/Cr and narration, show which Entries row each line
belongs to, and give you an **Amount to key** column to copy across.

### Amounts

The **Amount column is entered as** setting on Config decides how the figure you
type is read:

| Setting | You type | The file carries | Reads as |
|---|---|---|---|
| `Paise` (default) | `47400` | `...00000047400` | 474.00 |
| `Rupees` | `474.00` | `...00000047400` | 474.00 |

`Paise` matches the settlement report, where the last two digits are already the
paise, so nothing has to be converted by hand. In that mode the amount must be a
whole number - a typed decimal is rejected rather than guessed at.

### Config

| Setting | Default | Notes |
|---|---|---|
| Amount column is entered as | `Paise` | See above. |
| File name pattern | `PROPELG_TTUM_{DDMMYYYY}.txt` | Tokens use the file-name date. |
| File-name date offset (days) | `1` | The bank's sample is named one day after the value date it carries. Set `0` to use the value date. |
| Block generation when out of balance | `Yes` | `No` warns and lets you continue. |
| Write a line break after the last record | `No` | `No` matches the bank's sample, which ends immediately after the last record. |
| Input folder | *(blank)* | Where the settlement file arrives. Blank means you are asked each time. |
| Input file name pattern | `NET_MERPAY_PROPELG*.txt` | Which files in that folder count as settlement files. |
| Imported file sets the value date | `Yes` | `No` keeps whatever date is on the Dashboard. |

---

## Record layout

186 characters per record, records separated by CR LF, no line break after the
last record.

| Columns | Field | Length | Content |
|---|---|---|---|
| 1-14 | Account number | 14 | Left justified, space padded |
| 15-22 | Value date | 8 | `DDMMYYYY` |
| 23-45 | Amount | 23 | Whole paise, zero padded on the left. `20000` = 200.00 |
| 46 | Transaction type | 1 | `D` or `C` |
| 47-55 | Value date | 9 | `0` + `DDMMYYYY` |
| 56-81 | Filler | 26 | Spaces |
| 82-116 | Narration | 35 | Left justified, space padded |
| 117-186 | Filler | 70 | Spaces |

### What the macro refuses to do

- Generate when debits do not equal credits (unless you turn that rule off).
- Generate a record whose account number, narration or amount will not fit its field.
- Generate with a blank or non-numeric amount, or a Dr/Cr that is not `D` or `C`.
- Generate from a fractional amount while the Config setting is `Paise`.
- Overwrite an existing file without asking.

Rows that fail a check are highlighted on the Entries sheet and listed in the
message, and nothing is written.

---

## If Excel will not open the .xlsm at all

The workbook and its VBA project are generated by the scripts in `tools/`. The
output is validated against the format specifications and read back with an
independent parser, but it has not been opened in Microsoft Excel here. If your
Excel refuses the file, the same utility can be assembled by hand in about a
minute:

1. Open `dist/TTUM_Generator_NoMacros.xlsx` - identical sheets, no macros.
2. **File > Save As** and choose **Excel Macro-Enabled Workbook (*.xlsm)**.
3. Press `Alt+F11`, then **File > Import File**, and pick `dist/modTTUM.bas`.
4. Press `Alt+F8`, run `TTUM_Setup`, and save. The buttons appear on the Dashboard.

Also worth checking first: Windows blocks macros in files that arrive by email or
download. Right-click the file, choose **Properties**, tick **Unblock**, and
click OK.

---

## Repository layout

```
ttum/
  PRD.html  the product requirements document
  dist/     TTUM_Generator.xlsm             the utility
            TTUM_Generator_PRD.docx         the PRD as a Word file
            TTUM_Generator_NoMacros.xlsx    same sheets, macro-free fallback
            modTTUM.bas                     the macro source, for manual import
  vba/      modTTUM.bas, ThisWorkbook.cls   VBA source of record
  tools/    build_workbook.py               builds dist/ from vba/ and the row template
            standard_rows.py                the nine standard settlement lines
            ttum_reference.py               the record layout, mirrored in Python
            vbaproject.py, cfb.py, msovba.py  writes the embedded VBA project
            test_reference.py               layout regression test
            test_workbook.py                end-to-end test of the built workbook
            test_import.py                  end-to-end test of the import route
            test_formula.py                 evaluates the Text Output and Import
                                            sheets with a real formula engine
  samples/  the bank's file, a settlement input file, and the specification
```

### Rebuilding

```
pip install openpyxl
cd ttum/tools
python3 build_workbook.py
python3 test_reference.py     # layout reproduces the bank's file byte for byte
python3 test_workbook.py      # the built workbook's own rows do too
python3 test_import.py        # the settlement file imports and generates cleanly
python3 test_formula.py       # so do the macro-free Text Output and Import sheets
```

`test_formula.py` needs `pip install formulas`.

Edit the macro in `vba/modTTUM.bas` and rebuild; do not edit `dist/` by hand, it
is regenerated. Changing the nine standard lines is a change to
`tools/standard_rows.py`.
