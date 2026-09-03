const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, ExternalHyperlink,
} = require("docx");

// ---------------------------------------------------------------- palette
const INK = "16212E", INK2 = "42505F", MUTED = "6E7887";
const NAVY = "1F3355", RULE = "DBE1E9";
const F_ACCT = "E7EDF6", F_DATE = "E2F0EC", F_AMT = "FBF0D2",
      F_TYPE = "F7E9E2", F_FILL = "EFF2F6";

const BODY = "Calibri", DISPLAY = "Georgia", MONO = "Consolas";
const CONTENT_W = 9360;                       // Letter minus 1" margins, in DXA

// ---------------------------------------------------------------- helpers
const run = (text, o = {}) => new TextRun({
  text, font: o.font || BODY, size: o.size || 21,
  color: o.color || INK2, bold: !!o.bold, italics: !!o.italics,
  shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
});
const mono = (t, o = {}) => run(t, { ...o, font: MONO, size: o.size || 18, color: o.color || INK });
const b = (t) => run(t, { bold: true, color: INK });

// Cell and body content may arrive as a string, one run, or a list of runs.
const kids = (c, o = {}) =>
  typeof c === "string" ? [run(c, o)] : (Array.isArray(c) ? c : [c]);

// Body text. Accepts a string or a list of runs.
const P = (content, o = {}) => new Paragraph({
  children: kids(content, o),
  spacing: { after: o.after === undefined ? 140 : o.after, line: 276 },
  alignment: o.align,
});

const H1 = (num, text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 380, after: 60 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 8 } },
  children: [
    new TextRun({ text: num + "   ", font: MONO, size: 20, color: NAVY, bold: true }),
    new TextRun({ text, font: DISPLAY, size: 30, color: NAVY }),
  ],
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 90 },
  children: [new TextRun({ text, font: BODY, size: 21, bold: true, color: INK })],
});

const bullet = (content) => new Paragraph({
  children: kids(content),
  bullet: { level: 0 }, spacing: { after: 70, line: 276 },
});

const caption = (text) => new Paragraph({
  children: [run(text, { size: 17, color: MUTED, italics: true })],
  spacing: { after: 90 },
});

const cellBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

// A table with a shaded header row. `rows` cells may be a string or a run list.
function table(headers, rows, widths, opts = {}) {
  const head = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: F_FILL },
      margins: { top: 90, bottom: 90, left: 110, right: 110 },
      borders: cellBorders,
      children: [new Paragraph({
        children: [new TextRun({ text: h.toUpperCase(), font: MONO, size: 15, color: MUTED, bold: true })],
        spacing: { after: 0 },
      })],
    })),
  });
  const body = rows.map((cells) => new TableRow({
    children: cells.map((c, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 110, right: 110 },
      borders: cellBorders,
      children: [new Paragraph({
        children: kids(c, { size: 19 }),
        spacing: { after: 0, line: 250 },
      })],
    })),
  }));
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, x) => a + x, 0), type: WidthType.DXA },
    rows: [head, ...body],
    ...opts,
  });
}

// Requirement rows: an id in the margin, the requirement beside it.
function reqTable(items) {
  const widths = [1100, CONTENT_W - 1100];
  return new Table({
    columnWidths: widths,
    width: { size: CONTENT_W, type: WidthType.DXA },
    rows: items.map(([id, title, body]) => new TableRow({
      children: [
        new TableCell({
          width: { size: widths[0], type: WidthType.DXA },
          margins: { top: 110, bottom: 110, left: 0, right: 110 },
          borders: cellBorders,
          children: [new Paragraph({
            children: [new TextRun({ text: id, font: MONO, size: 17, color: NAVY, bold: true })],
            spacing: { after: 0 },
          })],
        }),
        new TableCell({
          width: { size: widths[1], type: WidthType.DXA },
          margins: { top: 110, bottom: 110, left: 0, right: 110 },
          borders: cellBorders,
          children: [
            new Paragraph({ children: [b(title)], spacing: { after: 50 } }),
            new Paragraph({
              children: kids(body, { size: 20 }),
              spacing: { after: 0, line: 264 },
            }),
          ],
        }),
      ],
    })),
  });
}

// A bordered callout.
function callout(tag, paras, accent) {
  return new Table({
    columnWidths: [CONTENT_W],
    width: { size: CONTENT_W, type: WidthType.DXA },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CONTENT_W, type: WidthType.DXA },
        margins: { top: 150, bottom: 150, left: 170, right: 170 },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "FAFBFC" },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
          left: { style: BorderStyle.SINGLE, size: 18, color: accent || NAVY },
          right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
        },
        children: [
          new Paragraph({
            children: [new TextRun({ text: tag.toUpperCase(), font: MONO, size: 15, color: MUTED, bold: true })],
            spacing: { after: 80 },
          }),
          ...paras.map((p, i) => new Paragraph({
            children: kids(p, { size: 20 }),
            spacing: { after: i === paras.length - 1 ? 0 : 100, line: 264 },
          })),
        ],
      })],
    })],
  });
}

// ------------------------------------------------- the record, byte by byte
const SEGMENTS = [
  ["00993562511731", F_ACCT], ["15072026", F_DATE],
  ["00000000000000000020000", F_AMT], ["D", F_TYPE], ["015072026", F_DATE],
  [" ".repeat(26), F_FILL], ["PROPELG VI settlement 150726       ", F_ACCT],
  [" ".repeat(70), F_FILL],
];

// Word cannot fit 186 monospaced characters across a Letter page, so the record
// is shown in two labelled halves with the field shading carried across.
function recordParagraphs(perLine) {
  const out = [];
  let line = [], used = 0, start = 1;
  const flush = () => {
    out.push(caption("Columns " + start + "–" + (start + used - 1)));
    out.push(new Paragraph({ children: line, spacing: { after: 160 } }));
    start += used; line = []; used = 0;
  };
  for (const [text, fill] of SEGMENTS) {
    let rest = text;
    while (rest.length) {
      const room = perLine - used;
      const take = rest.slice(0, room);
      line.push(mono(take, { size: 15, fill }));
      used += take.length;
      rest = rest.slice(take.length);
      if (used >= perLine) flush();
    }
  }
  if (used) flush();
  return out;
}

const legend = () => new Paragraph({
  spacing: { after: 200 },
  children: [
    mono("    ", { fill: F_ACCT }), run("  Account & narration    ", { size: 17, color: MUTED }),
    mono("    ", { fill: F_DATE }), run("  Value date    ", { size: 17, color: MUTED }),
    mono("    ", { fill: F_AMT }), run("  Amount in paise    ", { size: 17, color: MUTED }),
    mono("    ", { fill: F_TYPE }), run("  Dr / Cr    ", { size: 17, color: MUTED }),
    mono("    ", { fill: F_FILL }), run("  Filler", { size: 17, color: MUTED }),
  ],
});

// ---------------------------------------------------------------- content
const children = [];

// Masthead
children.push(new Paragraph({
  children: [new TextRun({ text: "PRODUCT REQUIREMENTS  ·  SETTLEMENT OPERATIONS",
    font: MONO, size: 15, color: NAVY, bold: true })],
  spacing: { after: 160 },
}));
children.push(new Paragraph({
  children: [new TextRun({ text: "TTUM Daily Generator", font: DISPLAY, size: 56, color: NAVY })],
  spacing: { after: 140 },
}));
children.push(new Paragraph({
  children: [run("An Excel utility that turns each day’s PROPELG settlement figures — typed by hand or read from the settlement system’s file — into the bank’s fixed-width TTUM upload file, without anyone counting characters.",
    { size: 23, color: INK2 })],
  spacing: { after: 260, line: 300 },
}));
children.push(table(
  ["Status", "Version", "Owner", "Updated"],
  [["Built & in acceptance", "1.0", "Vivek Yadav", "3 September 2026"]],
  [2600, 1400, 2600, 2760]));
children.push(P("", { after: 320 }));

// Contents
children.push(new Paragraph({
  children: [new TextRun({ text: "CONTENTS", font: MONO, size: 15, color: MUTED, bold: true })],
  spacing: { after: 120 },
}));
const TOC = ["Summary", "Background", "Goals and non-goals", "Users", "The record",
  "Domain rules", "Functional requirements", "Validation", "Workbook structure",
  "Operator flows", "Configuration", "Non-functional requirements", "Verification",
  "Open questions"];
TOC.forEach((t, i) => children.push(new Paragraph({
  children: [
    new TextRun({ text: String(i + 1).padStart(2, " ") + "   ", font: MONO, size: 18, color: MUTED }),
    run(t, { size: 20, color: INK }),
  ],
  spacing: { after: 40 },
})));

// 1 --------------------------------------------------------------------
children.push(H1("1", "Summary"));
children.push(P("Every business day, settlement figures for the PROPELG portfolio have to reach the bank as a TTUM upload file: nine transactions laid out in fixed-width records of exactly 186 characters, where a single misplaced digit is a failed upload or a misposted entry."));
children.push(P("This utility is a macro-enabled Excel workbook. The operator sets a date, gets the day’s amounts in — by typing them or by importing the settlement system’s own file — and clicks one button. The workbook lays out every record to the bank’s specification, refuses to write a file that does not balance or will not fit its fields, and keeps a log of what was produced."));
children.push(P("It runs entirely on the operator’s machine, needs no network, no add-ins and no server, and carries a formula-only fallback for environments where macros are blocked."));

// 2 --------------------------------------------------------------------
children.push(H1("2", "Background"));
children.push(P("The TTUM file is produced daily and is unforgiving: fields are positional, amounts are in paise with no decimal point, dates repeat in two different shapes, and more than a third of every record is padding. Preparing it by hand in a text editor is slow and error-prone, and an error is only discovered when the bank rejects the upload — or worse, accepts it."));
children.push(P("Two source documents defined the target: the bank’s written specification, and a real file the bank had already accepted."));
children.push(callout("Finding — specification vs. reality", [
  [run("The field lengths in the specification sheet sum to ", { size: 20 }), b("172"),
   run(" characters, but every record in the accepted file is ", { size: 20 }), b("186"),
   run(". The difference is the trailing filler: the spec says 56 characters, the real file carries 70. Every other field matches the spec exactly.", { size: 20 })],
  "The utility is built to the accepted file, on the grounds that the bank’s parser is the authority and that file demonstrably passed it. This should be confirmed with the bank so the specification can be corrected.",
], "B8860B"));
children.push(P("", { after: 200 }));

// 3 --------------------------------------------------------------------
children.push(H1("3", "Goals and non-goals"));
children.push(H2("Goals"));
children.push(reqTable([
  ["G1", "A correct file, every day, without character counting",
   "The operator supplies amounts and a date. Positions, padding, paise and date formats are the utility’s problem, never the operator’s."],
  ["G2", "Two ways in, one way out",
   "Amounts can be typed or imported from the settlement file, and the two can be mixed. Everything after that — checks, preview, generation, logging — is identical."],
  ["G3", "Wrong files are refused, not warned about",
   "An out-of-balance day, an overlong narration or a blank amount stops generation outright. The operator is shown exactly which row is at fault."],
  ["G4", "Nothing hidden",
   "The exact records can be inspected before any file is written, and every file produced is logged with its date, totals and path."],
  ["G5", "Survives a locked-down desktop",
   "Where macros are blocked by policy, the same records are still obtainable from the workbook using worksheet formulas alone."],
]));
children.push(H2("Non-goals"));
[
  "No connection to the bank. The utility writes a file; uploading it stays a manual, separately controlled step.",
  "No accounting logic. It does not calculate commission, GST or net settlement — those figures arrive already computed.",
  "No multi-entity or multi-portfolio support in v1. The row template is PROPELG’s.",
  "No reconciliation against bank statements or ledgers.",
  "No server, database, scheduler or unattended operation. A person triggers every run.",
].forEach((t) => children.push(bullet(t)));

// 4 --------------------------------------------------------------------
children.push(H1("4", "Users"));
children.push(P([b("The settlement operator"), run(" is the only daily user: a finance or operations colleague, fluent in Excel, not a developer, working on a managed Windows desktop. They run the utility once a day under time pressure, and they carry the consequence of an error. The design assumes they will not read documentation before clicking, so the workbook explains itself on the sheet.")]));
children.push(P([b("A reviewer or auditor"), run(" is an occasional user, needing to answer “what was sent on the 3rd, and who sent it?” from the workbook alone.")]));

// 5 --------------------------------------------------------------------
children.push(H1("5", "The record"));
children.push(P("Records are 186 characters, separated by CR LF, with no line break after the last one. Below is a real record from the accepted file — ₹200.00 debited for VI settlement on 15 July 2026 — coloured by field, split across two halves because a full record does not fit a page. The padding is shown at full width because it is two-thirds of the record."));
recordParagraphs(93).forEach((p) => children.push(p));
children.push(legend());
children.push(table(
  ["Columns", "Field", "Len", "Content"],
  [
    ["1–14", "Account number", "14", "Left justified, padded right with spaces"],
    ["15–22", "Value date", "8", "DDMMYYYY"],
    ["23–45", "Amount", "23", "Whole paise, padded left with zeros. 20000 is ₹200.00"],
    ["46", "Transaction type", "1", "D debit or C credit"],
    ["47–55", "Value date, again", "9", "0 followed by DDMMYYYY"],
    ["56–81", "Filler", "26", "Spaces"],
    ["82–116", "Narration", "35", "Left justified, padded right with spaces"],
    ["117–186", "Filler", "70", "Spaces"],
  ].map((r) => [mono(r[0], { size: 18 }), r[1], mono(r[2], { size: 18 }), r[3]]),
  [1500, 2100, 800, 4960]));

// 6 --------------------------------------------------------------------
children.push(H1("6", "Domain rules"));
children.push(H2("Amounts are keyed in paise"));
children.push(P([run("The settlement report gives figures with the paise as the last two digits and no decimal point, so that is what the operator types: "), mono("47400"), run(" means ₹474.00. The sheet shows a live "), b("Reads as (INR)"), run(" column beside each entry so the decimal point is never in doubt, and a fractional entry is rejected rather than guessed at. A configuration setting switches the column to rupees for anyone who prefers "), mono("474.00"), run(".")]));
children.push(H2("Two dates, and they differ"));
children.push(P([b("The value date"), run(" is carried inside every record, twice. "), b("The file-name date"), run(" is what the file is called. In the bank’s accepted example these differ by one day — records dated 15 July in a file named for 16 July — so the offset is a setting, defaulting to one day, and the resolved file name is shown on screen before generation.")]));
children.push(H2("The day must balance"));
children.push(P("Total debits must equal total credits. Both sample days do so exactly, and the balancing entry is the net settlement line. This is treated as a hard rule: an unbalanced day blocks generation by default, and the running difference is displayed while amounts are entered."));
children.push(H2("Narrations are templates"));
children.push(P([run("A narration is stored with date placeholders — "), mono("{DDMMYY}"), run(", "), mono("{DDMMYYYY}"), run(", "), mono("{DD}"), run(", "), mono("{MM}"), run(", "), mono("{YY}"), run(", "), mono("{YYYY}"), run(" — resolved against the value date at generation. "), mono("PROPELG VI settlement {DDMMYY}"), run(" becomes "), mono("PROPELG VI settlement 150726"), run(", so no date is ever retyped.")]));
children.push(H2("The nine standard lines"));
children.push(table(
  ["Line", "Account", "Dr/Cr", "Import match text"],
  [
    ["VI settlement", "00993562511731", "D", "VI settlement"],
    ["MC settlement", "00993562511732", "D", "MC settlement"],
    ["MC commission received", "00993564610119", "C", "MC comm recd"],
    ["VI commission received", "00993564610122", "C", "VI comm recd"],
    ["MC GST on commission", "00993564610119", "C", "MC GST comm recd"],
    ["VI GST on commission", "00993564610122", "C", "VI GST comm recd"],
    ["MC Non-GST commission", "00993564610119", "C", "MC Non GST comm recd"],
    ["VI Non-GST commission", "00993564610122", "C", "VI Non GST comm recd"],
    ["Net settlement to Prp India", "200999103427", "C", "Nodal"],
  ].map((r) => [r[0], mono(r[1], { size: 18 }), r[2], mono(r[3], { size: 18 })]),
  [2900, 2100, 800, 3560]));
children.push(P("", { after: 120 }));
children.push(P([run("The match keys carry their full wording deliberately: "), mono("MC comm recd"), run(" must not also claim "), mono("MC GST comm recd"), run(". Rows are operator-editable, and the sheet holds a hundred of them, so lines can be added without code changes.")]));

// 7 --------------------------------------------------------------------
children.push(H1("7", "Functional requirements"));
children.push(reqTable([
  ["FR-1", "Date selection", "The value date defaults to the system date and accepts any date typed over it. The file-name date follows at a configurable offset and is itself editable. Both are visible before generation."],
  ["FR-2", "Manual entry", "A hundred-row grid holding include flag, description, account number, Dr/Cr, amount and narration template, pre-loaded with the nine standard lines. Dropdowns constrain the include and Dr/Cr columns; account numbers are held as text so leading zeros survive."],
  ["FR-3", "Import from the settlement file", "Reads the newest file matching a pattern from a designated folder, or a file the operator picks. The settlement file uses the same 186-character layout, so the same field positions parse it."],
  ["FR-4", "Import summary before any change", "A summary shows the file name, value date, record count, how many lines matched, every matched amount, and the debit/credit totals with their difference. Declining leaves the sheet exactly as it was."],
  ["FR-5", "Import takes figures only", "Only the amount and the Dr/Cr flag come from the file. Account numbers and narrations always come from the entry sheet — which is what keeps the net settlement line correct, since the settlement file books it to the nodal account while the TTUM must carry the Prp India account."],
  ["FR-6", "Import leaves nothing stale", "Rows with a match key that the file did not mention are cleared and excluded, so a previous day’s figure cannot survive. Rows with no match key are untouched, so anything maintained by hand persists across imports."],
  ["FR-7", "Unmatched lines cannot slip through", "A line matching no row, or matching more than one, is reported and left out. That leaves the day out of balance, which blocks generation — so silently dropping a transaction is not a reachable outcome."],
  ["FR-8", "Preview", "The exact records can be rendered on screen, character for character, with their lengths, without writing anything to disk."],
  ["FR-9", "Generation", "Writes ASCII, CR LF separated records with no trailing line break by default. An existing file of the same name is never replaced without asking."],
  ["FR-10", "Output location", "A folder is typed, pasted or chosen through a picker, and defaults to a folder beside the workbook. Missing folders are created. The resolved file name is shown live as the date changes."],
  ["FR-11", "Audit log", "Every generated file adds a row recording timestamp, value date, file date, file name, record count, both totals, the balance state, the Windows user and the full path."],
  ["FR-12", "Macro-free fallback", "A sheet builds the same 186-character records using worksheet formulas alone, for copying into a text editor. A second sheet parses a pasted settlement file the same way. Both work with macros fully disabled."],
]));

// 8 --------------------------------------------------------------------
children.push(H1("8", "Validation"));
children.push(P("Checks run before anything is written. Failing cells are highlighted on the entry sheet and named in a single message; no partial file is ever produced."));
children.push(table(
  ["ID", "Rule", "On failure"],
  [
    ["VR-1", "Total debits equal total credits", "Blocks generation. Configurable to warn-and-continue instead"],
    ["VR-2", "Account number is present and at most 14 characters", "Blocks, names the row"],
    ["VR-3", "Dr/Cr is D or C", "Blocks, names the row"],
    ["VR-4", "Amount is present, numeric and greater than zero", "Blocks, distinguishing blank from non-numeric"],
    ["VR-5", "Amount is a whole number when keyed in paise", "Blocks — a decimal here is ambiguous, not roundable"],
    ["VR-6", "Amount fits the 23-digit field", "Blocks"],
    ["VR-7", "Narration is present and at most 35 characters after date substitution", "Blocks — never silently truncated"],
    ["VR-8", "Every assembled record is exactly 186 characters", "Aborts generation — a last-line defence against a layout change"],
    ["VR-9", "The value date is a valid date", "Blocks"],
    ["VR-10", "Output file does not already exist", "Asks before replacing"],
  ].map((r) => [mono(r[0], { size: 18 }), r[1], r[2]]),
  [900, 4560, 3900]));

// 9 --------------------------------------------------------------------
children.push(H1("9", "Workbook structure"));
children.push(table(
  ["Sheet", "Purpose"],
  [
    ["Dashboard", "Dates, output folder, live totals, status line, and the nine action buttons"],
    ["Entries", "The daily grid, 100 rows, with the running debit/credit banner"],
    ["Import", "The last settlement file line by line, with each field pulled out and the row it matched"],
    ["Text Output", "The same records built by formula — the macro-free route"],
    ["Config", "Eight settings that rarely change"],
    ["Log", "One row per generated file"],
    ["Layout", "The record specification and the bank’s file as a worked example"],
    ["Setup", "What to try when a button does nothing"],
  ].map((r) => [[b(r[0])], r[1]]),
  [2100, 7260]));
children.push(P("", { after: 120 }));
children.push(P("Cells are addressed through named ranges rather than fixed positions, so the sheets can be re-laid-out without touching the code. The macro source is a single module of roughly 1,240 lines."));

// 10 -------------------------------------------------------------------
children.push(H1("10", "Operator flows"));
children.push(H2("Import route"));
[
  ["Open the workbook.", " Today’s date is already in place."],
  ["Import Latest Input File.", " The newest matching file is read from the designated folder."],
  ["Read the summary.", " Value date, record count, every matched amount, and whether the day balances. Decline and nothing changes."],
  ["Accept.", " Amounts land on the entry sheet, the value date is taken from the file, and the operator is returned to the dashboard."],
  ["Check the totals, then Generate.", " The file is written and logged."],
].forEach(([head, rest], i) => children.push(new Paragraph({
  children: [
    new TextRun({ text: (i + 1) + "   ", font: MONO, size: 18, color: MUTED }),
    b(head), run(rest),
  ],
  spacing: { after: 90, line: 276 },
})));
children.push(H2("Manual route"));
[
  ["Set the date", " if it is not today."],
  ["Type the amounts", " against the nine standard lines, adjusting Dr/Cr where the day’s flow differs. The banner turns green when the day balances."],
  ["Generate.", " Identical from here on."],
].forEach(([head, rest], i) => children.push(new Paragraph({
  children: [
    new TextRun({ text: (i + 1) + "   ", font: MONO, size: 18, color: MUTED }),
    b(head), run(rest),
  ],
  spacing: { after: 90, line: 276 },
})));
children.push(P("The two mix freely: an imported day can be adjusted by hand before generating, and rows without a match key are never disturbed by an import.", { after: 60 }));

// 11 -------------------------------------------------------------------
children.push(H1("11", "Configuration"));
children.push(table(
  ["Setting", "Default", "Effect"],
  [
    ["Amount column is entered as", "Paise", "Whether a keyed figure carries its paise as the last two digits, or is a rupee amount"],
    ["File name pattern", "PROPELG_TTUM_{DDMMYYYY}.txt", "Date tokens resolve against the file-name date"],
    ["File-name date offset", "1", "Days between the value date and the date in the file name"],
    ["Block when out of balance", "Yes", "Whether an unbalanced day stops generation or merely warns"],
    ["Line break after last record", "No", "Matches the bank’s accepted file, which ends at the last character"],
    ["Input folder", "—", "Where the settlement file arrives. Blank means the operator is asked each time"],
    ["Input file name pattern", "NET_MERPAY_PROPELG*.txt", "Which files in that folder count as settlement files"],
    ["Imported file sets the value date", "Yes", "Whether the file’s own date overrides the dashboard"],
  ].map((r) => [r[0], mono(r[1], { size: 17 }), r[2]]),
  [2700, 2500, 4160]));

// 12 -------------------------------------------------------------------
children.push(H1("12", "Non-functional requirements"));
children.push(reqTable([
  ["NFR-1", "Environment", "Microsoft Excel on Windows, 2010 or later. No add-ins, no references beyond those Excel ships with, no installer, no server. The workbook is a single file that can be copied between machines."],
  ["NFR-2", "Offline by construction", "No network calls, no credentials, no external services. Inputs and outputs are local files."],
  ["NFR-3", "Exact money arithmetic", "Amount conversion uses fixed-point currency arithmetic, not floating point, so no rounding drift can reach a settlement figure."],
  ["NFR-4", "Byte-level output control", "The file is written as raw bytes rather than through text output, so line endings and the absence of a trailing break are exactly as intended."],
  ["NFR-5", "Speed", "A day’s import, validation and generation complete in well under a second at the working volume of nine to a hundred records."],
  ["NFR-6", "Degrades rather than fails", "Buttons are stored in the file so they are visible before macros run; if macros are blocked the formula route still produces the records; if the macro project fails to load it can be imported by hand in about a minute."],
  ["NFR-7", "Maintainable without a developer", "Rows, accounts, narrations, match keys, file naming and folders are all sheet data. Code changes are needed only if the bank changes the record layout."],
]));

// 13 -------------------------------------------------------------------
children.push(H1("13", "Verification"));
children.push(P("The workbook is generated from source by a build script, and four automated tests run against the built file. All four pass."));
children.push(table(
  ["Test", "What it proves", "Result"],
  [
    ["test_reference", "The layout rebuilds the bank’s accepted file byte for byte — 1,690 bytes, 9 records", "Pass"],
    ["test_workbook", "The shipped rows and settings reproduce that same file, so the workbook’s own content is correct, not just the constants", "Pass"],
    ["test_import", "The real settlement file parses, all 9 lines match 1:1, the day balances at ₹78,386,928.00, and a valid file is produced", "Pass"],
    ["test_formula", "The macro-free sheets, evaluated with an Excel formula engine, reproduce every record and parse every settlement line", "Pass"],
  ].map((r) => [mono(r[0], { size: 17 }), r[1],
                [run(r[2], { bold: true, color: "12695A", size: 19 })]]),
  [1900, 6160, 1300]));
children.push(P("", { after: 120 }));
children.push(callout("Acceptance still outstanding", [
  "These tests run outside Excel. The remaining acceptance step is a live run on the operator’s own desktop: buttons respond, an import completes, and the generated file is accepted by the bank’s upload.",
]));
children.push(P("", { after: 200 }));

// 14 -------------------------------------------------------------------
children.push(H1("14", "Open questions"));
children.push(reqTable([
  ["Q1", "Which account carries the net settlement line?",
   [run("The settlement file books it to the nodal account ", { size: 20 }), mono("00993564610251", { size: 18 }),
    run("; the bank’s accepted TTUM carries ", { size: 20 }), mono("200999103427", { size: 18 }),
    run(" for Prp India. The utility keeps the TTUM account and takes only the figure. Confirm this is right, or change the account on that row.", { size: 20 })]],
  ["Q2", "Is the trailing filler 70 characters or 56?",
   "The accepted file says 70, the written specification says 56. Built to the file; worth correcting the specification."],
  ["Q3", "Is the one-day file-name offset a rule?",
   "Inferred from a single example. If the offset varies — over weekends, say — the setting covers it, but the default should be confirmed."],
  ["Q4", "Do the standard narrations always fit 35 characters?",
   "The longest currently reaches exactly 35. Any added wording will be rejected rather than truncated, which is safe but will stop a day’s run."],
  ["Q5", "Does the desktop policy permit macros at all?",
   "Determines whether the utility runs as designed or through the formula route. Unresolved on the operator’s machine at the time of writing."],
]));
children.push(H2("Beyond v1"));
[
  "Support for further portfolios or entities using the same row-template mechanism.",
  "Archiving each generated file alongside the settlement file that produced it.",
  "A reconciliation view comparing what was generated against what the bank acknowledged.",
].forEach((t) => children.push(bullet(t)));

// Colophon
children.push(new Paragraph({
  spacing: { before: 360, after: 0 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 12 } },
  children: [run("Built from source in ttum/: macro source under vba/, the build script and tests under tools/, the shipped workbook under dist/. The bank’s specification, its accepted file and a real settlement file are kept in samples/ as the fixtures the tests run against.",
    { size: 18, color: MUTED })],
}));

// ---------------------------------------------------------------- assemble
const doc = new Document({
  creator: "Vivek Yadav",
  title: "TTUM Daily Generator — Product Requirements",
  description: "Product requirements for the Excel utility that builds the bank's daily fixed-width TTUM settlement upload file.",
  styles: {
    default: {
      document: { run: { font: BODY, size: 21, color: INK2 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "TTUM_Generator_PRD.docx", buf);
  console.log("written", buf.length, "bytes");
});
