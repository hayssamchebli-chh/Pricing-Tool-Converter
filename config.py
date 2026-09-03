"""Layout + routing configuration for the pricing tool converter.

The template workbook (`template/pricing_tool_template.xlsx`) contains six
"offer" sheets that are built from a supplier catalogue extract laid out like
`Sheet7` / `Sheet9` in the reference file.  Two physical layouts are in use:

  * layout ``A`` (compact, 20 columns)  -> O (UE), O (UE3), O (UE4)
  * layout ``B`` (extended, 29 columns) -> O (UU), O (UE1), O (UE2)

Layout ``B`` carries three extra "Last U.P. / Cur. / Date" history blocks
(columns F..N) which push the costing columns to the right.
"""

from __future__ import annotations

from dataclasses import dataclass, field

CATALOG_HEADERS = [
    "Item No.1",
    "Description",
    "Unit Price",
    "Advanced Reserved",
    "Stock Available Quantity",
    "PO Qty",
    "PO not Shipped",
    "Landed USD",
]

# Column index (1-based) of each catalogue field inside the Catalogue sheet.
CAT_COL = {
    "code": 1,
    "description": 2,
    "unit_price": 3,
    "advanced_reserved": 4,
    "stock": 5,
    "po_qty": 6,
    "po_not_shipped": 7,
    "landed_usd": 8,
}

FIRST_DATA_ROW = 3          # offer sheets: rows 1-2 are constants + headers
FACTOR_ROW = 1              # freight factor sits above the U. Landed header
CATALOG_SHEET = "Catalogue"  # the single lookup table every offer sheet reads
CATALOG_FIRST_ROW = 2       # Catalogue sheet: row 1 is the header
VAT_RATE = 0.11


@dataclass(frozen=True)
class Layout:
    """Column map of an offer sheet (all values are 1-based column indexes)."""

    name: str
    num: int
    code: int
    description: int
    qty: int
    unit_price: int
    landed: int          # "U. Landed (USD)"
    ex_works: int        # "U.P. Ex."      - manually keyed by the estimator
    disc_unit: int       # "D.U.P. Ex."
    disc: int            # "Disc."
    disc_total: int      # "D.T.P. Ex."
    total_landed: int    # "T. Landed (USD)"
    total: int           # "Total"
    margin: int
    advanced_reserved: int
    stock: int
    po_qty: int
    po_not_shipped: int
    landed_usd: int
    last_col: int


LAYOUT_A = Layout(
    name="A", num=1, code=2, description=3, qty=4, unit_price=5,
    landed=6, ex_works=7, disc_unit=8, disc=9, disc_total=10,
    total_landed=11, total=12, margin=14,
    advanced_reserved=16, stock=17, po_qty=18, po_not_shipped=19,
    landed_usd=20, last_col=20,
)

LAYOUT_B = Layout(
    name="B", num=1, code=2, description=3, qty=4, unit_price=5,
    landed=15, ex_works=16, disc_unit=17, disc=18, disc_total=19,
    total_landed=20, total=21, margin=23,
    advanced_reserved=25, stock=26, po_qty=27, po_not_shipped=28,
    landed_usd=29, last_col=29,
)

LAYOUTS = {"A": LAYOUT_A, "B": LAYOUT_B}


@dataclass
class SheetSpec:
    """One offer sheet: where its rows come from and which codes land in it."""

    sheet: str
    layout: str
    source: str            # the lookup table this sheet's VLOOKUPs read from
    currency: str          # "EUR" or "USD" - drives the landed-cost fallback
    prefixes: list = field(default_factory=list)

    @property
    def cols(self) -> Layout:
        return LAYOUTS[self.layout]


# Default routing, reverse-engineered from the reference workbook: every offer
# sheet holds one supplier / product family, identified by the item-code prefix.
# All six read one lookup table; ``source`` stays per-sheet so a second table
# can be introduced by editing this list if two suppliers ever collide on a code.
DEFAULT_SPECS = [
    SheetSpec("O (UE)",  "A", CATALOG_SHEET, "EUR", ["ABS"]),
    SheetSpec("O (UU)",  "B", CATALOG_SHEET, "USD", ["ABR"]),
    SheetSpec("O (UE1)", "B", CATALOG_SHEET, "EUR", ["GAV", "TKM"]),
    SheetSpec("O (UE2)", "B", CATALOG_SHEET, "EUR", ["ABD"]),
    SheetSpec("O (UE3)", "A", CATALOG_SHEET, "EUR", ["ABF"]),
    SheetSpec("O (UE4)", "A", CATALOG_SHEET, "EUR",
              ["AAB", "ABA", "ABAE", "ABB", "ABE", "ABJ", "ABZ"]),
]

# Sheet that receives catalogue rows whose prefix matches no rule.
FALLBACK_SHEET = "O (UE4)"

SUMMARY_SHEET = "Summary"
SUMMARY_FIRST_ROW = 2

# Landed-cost fallback used when the catalogue reports no landed cost: the
# estimator keys an ex-works price and it is grossed up by these factors.
DEFAULT_FREIGHT_FACTOR = 1.15
DEFAULT_EUR_FACTOR = 1.16
