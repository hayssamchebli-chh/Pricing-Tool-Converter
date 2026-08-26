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
original: descriptions and prices are `VLOOKUP`s into `Sheet7` / `Sheet9`,
totals are `SUM`s over the actual row range, and the `Summary` points at each
sheet's real footer rows.

### Sheet routing

Each offer sheet holds one supplier or product family, matched on the item-code
prefix — the longest matching prefix wins, so `ABAE-…` beats a plain `ABA`
rule. The defaults reproduce the reference workbook:

| Sheet | Prefixes | Lookup table | Currency |
| --- | --- | --- | --- |
| O (UE) | ABS | Sheet9 | EUR |
| O (UU) | ABR | Sheet7 | USD |
| O (UE1) | GAV, TKM | Sheet7 | EUR |
| O (UE2) | ABD | Sheet7 | EUR |
| O (UE3) | ABF | Sheet7 | EUR |
| O (UE4) | AAB, ABA, ABAE, ABB, ABE, ABJ, ABZ | Sheet7 | EUR |

All of it is editable in the app, per run. A code matching no rule lands on
`O (UE4)` and is reported rather than dropped, and individual rows can be moved
to another sheet in the review table.

`Sheet7` and `Sheet9` in the output are rebuilt from the catalogue, split the
same way the offer sheets are — so `Sheet9` ends up holding exactly the items
whose offer sheet reads from it.

### Landed cost

`U. Landed (USD)` is written as:

```
=IF(<U.P. Ex.>=0, <Landed USD>, <U.P. Ex.> * freight [* EUR factor])
```

The catalogue's landed cost applies until an ex-works price is keyed into the
`U.P. Ex.` column, which is the manual override for an item never actually
imported or one whose recorded cost is stale. EUR sheets gross up by freight
and the EUR factor, the USD sheet by freight alone; both factors are set in the
sidebar (1.15 and 1.16 by default). The reference workbook did this by editing
cells from constant to formula by hand — here the override is live.

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
| `catalog.py` | reading and normalising the uploaded workbooks |
| `builder.py` | writing the offer sheets, footers and Summary |
| `config.py` | column maps for both layouts, routing defaults |
| `template/pricing_tool_template.xlsx` | the styled template |

Changing the routing defaults permanently, or the VAT rate and gross-up
factors, means editing `config.py`; everything else is adjustable per run in
the app.
