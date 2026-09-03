"""The nine standard PROPELG settlement rows, as seen in the bank's sample file.

`amount` values here are the sample-day figures in rupees; they are only used to
regression-test the layout. In the workbook the amounts are typed daily by the user.
"""

STANDARD_ROWS = [
    dict(label="VI settlement",          account="00993562511731", amount=200.00, type="D",
         narration="PROPELG VI settlement {DDMMYY}"),
    dict(label="MC settlement",          account="00993562511732", amount=300.00, type="D",
         narration="PROPELG MC settlement {DDMMYY}"),
    dict(label="MC commission received", account="00993564610119", amount=10.00,  type="C",
         narration="PROPELG MC comm recd {DDMMYY}"),
    dict(label="VI commission received", account="00993564610122", amount=10.00,  type="C",
         narration="PROPELG VI comm recd {DDMMYY}"),
    dict(label="MC GST on commission",   account="00993564610119", amount=2.00,   type="C",
         narration="PROPELG MC GST comm recd {DDMMYY}"),
    dict(label="VI GST on commission",   account="00993564610122", amount=2.00,   type="C",
         narration="PROPELG VI GST comm recd {DDMMYY}"),
    dict(label="MC Non-GST commission",  account="00993564610119", amount=1.00,   type="C",
         narration="PROPELG MC Non GST comm recd {DDMMYY}"),
    dict(label="VI Non-GST commission",  account="00993564610122", amount=1.00,   type="C",
         narration="PROPELG VI Non GST comm recd {DDMMYY}"),
    dict(label="Net settlement to Prp India", account="200999103427", amount=474.00, type="C",
         narration="PROPELG-TXN Prp India {DDMMYY} {DDMMYY}"),
]
