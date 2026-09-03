# Pricing Tool Converter

Streamlit app that turns a supplier catalogue extract — the `Sheet7` / `Sheet9`
layout — into the six offer sheets and the `Summary` of `pricing tool.xlsx`.

## Running it

```bash
pip install -r requirements.txt
```

```bash
streamlit run app.py
```

On Windows, double-clicking `run.cmd` does the same thing.

## What goes in

**Catalogue extract** (required). Any worksheet whose header row carries an item
code, a description and a unit price in three consecutive columns is read; a
file holding both a `Sheet7`-style and a `Sheet9`-style tab is read in one pass.
The eight recognised columns are the ones in the reference workbook:

| Column | Used for |
| --- | --- |
| Item No.1 | the item code, and the key every VLOOKUP matches on |
| Description | offer sheet description |
| Unit Price | selling price, drives `Total` |
| Advanced Reserved | reference column on the offer sheet |
| Stock Available Quantity | reference column on the offer sheet |
| PO Qty | reference column on the offer sheet |
| PO not Shipped | reference column on the offer sheet |
| Landed USD | landed cost, drives `T. Landed` and the margin |

Header matching is tolerant of case, spacing and common variants (`Part No.`,
`List Price`, `Stock Qty`, …).

**Quantities** (optional). Either the raw BOQ or the "sum of qty" pivot built
from it — any sheet pairing an item-code column with a quantity column. Both
usually live in the same file and carry the same figures, so the app keeps them
apart and lets you pick one rather than adding them together. A code appearing
on several panels is summed. Without this file every quantity starts at 0 and
can be typed into the table.

## What comes out

A workbook built on `template/pricing_tool_template.xlsx`, so column widths,
number formats, fills and the constants on row 1 of every offer sheet are
carried over untouched. Only the data band, the footers and the `Summary`
cross-references are rewritten.

Every derived cell is a live formula, so the result stays as editable as the
original: descriptions and prices are `VLOOKUP`s into the `Catalogue` sheet,
totals are `SUM`s over the actual row range, and the `Summary` points at each
sheet's real footer rows.

### Sheet routing

Each offer sheet holds one supplier or product family, matched on the item-code
prefix — the longest matching prefix wins, so `ABAE-…` beats a plain `ABA`
rule. The defaults reproduce the reference workbook:

| Sheet | Prefixes | Currency |
| --- | --- | --- |
| O (UE) | ABS | EUR |
| O (UU) | ABR | USD |
| O (UE1) | GAV, TKM | EUR |
| O (UE2) | ABD | EUR |
| O (UE3) | ABF | EUR |
| O (UE4) | AAB, ABA, ABAE, ABB, ABE, ABJ, ABZ | EUR |

All of it is editable in the app, per run. A code matching no rule lands on
`O (UE4)` and is reported rather than dropped, and individual rows can be moved
to another sheet in the review table.

### The Catalogue sheet

The output carries one extra tab, `Catalogue`, holding every item read from the
upload. It is the lookup table the offer sheets read: each row's description and
unit price is a `VLOOKUP` into it, so deleting the tab breaks every offer row.
The reference workbook did the same thing across two tabs named `Sheet7` and
`Sheet9`; item codes are unique, so one table serves all six offer sheets and no
tab can come out empty.

The sheet's name has nothing to do with the upload — a file whose only tab is
called `Sheet1` still produces a `Catalogue` tab. To route two suppliers that
collide on a code to separate tables, give the relevant `SheetSpec` a different
`source` in `config.py`; the builder writes one table per distinct source.

### Landed cost

`U. Landed (USD)` is written as:

```
=IF(<Qty> < <Stock>, <Landed USD>, <U.P. Ex.> * $freight [* EUR factor])
```

Quantity decides the branch. While stock covers the quantity the row takes the
landed cost exactly as uploaded — what that stock actually cost to land. Once
Qty reaches or exceeds stock the order has to be imported, so it is priced off
the ex-works figure instead, grossed up by freight and the EUR factor on a EUR
sheet and by freight alone on the USD one — and its **Qty cell turns yellow**,
so the rows still waiting on an ex-works price are visible at a glance.

The yellow is a conditional-formatting rule rather than a painted fill, so it
keeps up as quantities are retyped in Excel.

The freight factor is written into the cell directly above the `U. Landed`
header — `F1` on a compact sheet, `O1` on an extended one — and every landed
formula multiplies by it, so retyping it there reprices the sheet. The EUR
conversion factor stays a literal in the formula, fixed from the sidebar at
build time and deliberately not surfaced as a cell.

`U.P. Ex.` is left empty for you to fill, and the three columns that read it —
`D.U.P. Ex.`, `Disc.` and `D.T.P. Ex.`, plus the `D.T.P. Ex.` footer — stay
blank rather than showing 0.00 on every row that is still priced off the
catalogue. They come to life on the rows where a price is actually keyed.

### Two layouts

The template's offer sheets come in two shapes and the app writes each in its
own:

* **compact** (20 columns) — `O (UE)`, `O (UE3)`, `O (UE4)`
* **extended** (29 columns) — `O (UU)`, `O (UE1)`, `O (UE2)`, which carry three
  extra *Last U.P. / Cur. / Date* history blocks that push the costing columns
  from F–T out to O–AC

## Files

| File | Contents |
| --- | --- |
| `app.py` | Streamlit UI |
| `ui.py` | masthead, section headers, metric cards, sheet chips |
| `.streamlit/config.toml` | palette, type and radii for light and dark |
| `catalog.py` | reading and normalising the uploaded workbooks |
| `builder.py` | writing the offer sheets, footers and Summary |
| `config.py` | column maps for both layouts, routing defaults |
| `template/pricing_tool_template.xlsx` | the styled template |

### Theme

The palette is navy on a cool near-white, with green kept for success only, and
a dark variant tuned separately rather than inverted. Type is Inter throughout
with JetBrains Mono for figures, and numerals are tabular so item codes and
prices line up in columns.

Streamlit publishes no CSS variables for its active theme, so the chrome in
`ui.py` derives its tones from `currentColor` — that keeps text and surfaces
correct in either theme instead of baking in light-mode literals. The one fixed
hue, the brand accent, is pinned per render from `st.context.theme`. All text
pairs measure at or above 4.5:1 in both themes.

Changing the routing defaults permanently, or the VAT rate and gross-up
factors, means editing `config.py`; everything else is adjustable per run in
the app.
