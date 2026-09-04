"""The nine standard PROPELG settlement rows, as seen in the bank's sample file.

`amount` values here are the sample-day figures in rupees; they are only used to
regression-test the layout. In the workbook the amounts are typed daily by the user,
or imported from the settlement file.

`key` is the text an incoming settlement line's narration must contain for it to be
matched to this row. The keys have to stay distinct from one another - "MC comm recd"
must not also match "MC GST comm recd" - which is why each carries its full wording.
Only the amount and the Dr/Cr flag are taken from a matched line; the account number
and the narration that go into the generated file always come from this sheet.
"""

STANDARD_ROWS = [
    dict(label="VI settlement",          account="00993562511731", amount=200.00, type="D",
         narration="PROPELG VI settlement {DDMMYY}",
         key="VI settlement"),
    dict(label="MC settlement",          account="00993562511732", amount=300.00, type="D",
         narration="PROPELG MC settlement {DDMMYY}",
         key="MC settlement"),
    dict(label="MC commission received", account="00993564610119", amount=10.00,  type="C",
         narration="PROPELG MC comm recd {DDMMYY}",
         key="MC comm recd"),
    dict(label="VI commission received", account="00993564610122", amount=10.00,  type="C",
         narration="PROPELG VI comm recd {DDMMYY}",
         key="VI comm recd"),
    dict(label="MC GST on commission",   account="00993564610119", amount=2.00,   type="C",
         narration="PROPELG MC GST comm recd {DDMMYY}",
         key="MC GST comm recd"),
    dict(label="VI GST on commission",   account="00993564610122", amount=2.00,   type="C",
         narration="PROPELG VI GST comm recd {DDMMYY}",
         key="VI GST comm recd"),
    dict(label="MC Non-GST commission",  account="00993564610119", amount=1.00,   type="C",
         narration="PROPELG MC Non GST comm recd {DDMMYY}",
         key="MC Non GST comm recd"),
    dict(label="VI Non-GST commission",  account="00993564610122", amount=1.00,   type="C",
         narration="PROPELG VI Non GST comm recd {DDMMYY}",
         key="VI Non GST comm recd"),
    dict(label="Net settlement to Prp India", account="200999103427", amount=474.00, type="C",
         narration="PROPELG-TXN Prp India {DDMMYY} {DDMMYY}",
         key="Nodal"),
]
