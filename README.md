# Pricing Tool Converter

Three tendering tools in one Streamlit app, a tab each:

| Tab | Turns | Into |
| --- | --- | --- |
| **Pricing Tool Converter** | a supplier catalogue extract, the `Sheet7` / `Sheet9` layout | the six offer sheets and the `Summary` of `pricing tool.xlsx` |
| **Last Purchase Price** | a purchase-history export | an Order workbook carrying the three most recent prices per item |
| **Cost + Customer Pricing Tool** | a customer sales report, a catalogue extract and either a Calcul workbook or a BOQ | the `Calcul` offer sheet with its price history, descriptions and reference columns filled |

## Running it

```bash
pip install -r requirements.txt
```

```bash
streamlit run home.py
```

On Windows, double-clicking `run.cmd` does the same thing.

`home.py` is only the tab bar. Each tab is its own script and runs on its own,
so `streamlit run app.py` still opens the Pricing Tool Converter by itself,
exactly as it did before the second tab existed.

---

# The Pricing Tool Converter tab

## What goes in

**Catalogue extract** (required). A worksheet is treated as catalogue when its
header row carries a column headed **`Item No.1`** — the catalogue's own name
for the code column — alongside a description and a unit price. Offer sheets
say "Item Code" instead, which is what keeps them from being read as catalogue
when the reference workbook itself is uploaded. Columns need not be adjacent:
anything extra can sit between them. A file holding several catalogue tabs is
read in one pass. The eight recognised columns are the ones in the reference
workbook:

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
`List Price`, `Stock Qty`, …). If nothing in the file carries `Item No.1`, the
reader falls back first to the three sitting in consecutive columns, then to
finding them anywhere in the row.

**Quantities** (optional). Either the raw BOQ or the "sum of qty" pivot built
from it — any sheet pairing an item-code column with a quantity column. Both
usually live in the same file and carry the same figures, so the app keeps them
apart and lets you pick one rather than adding them together. A code appearing
on several panels is summed. Without this file every quantity starts at 0 and
can be typed into the table.

The quantity column is found by header (`Qty`, `Quantity`, `Sum of QTY`, …).
The **code** column is found by content: each column is scored against the codes
already read from the catalogue and the best match wins. Header names for it
vary per export — `No.2` in one, `Part No.` in another — and exports often carry
a second identifier in the neighbouring column, so a name match alone picks the
wrong column as readily as the right one. Names are still used as a fallback
when nothing overlaps.

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

All of it is editable in the app, per run — **including which sheets exist**.
Adding a row to the routing table creates a sheet: name it, give it prefixes,
pick a layout and a currency. Renaming a row renames the sheet, and deleting one
drops it. A new sheet is cloned from the template sheet of its layout, so it
carries the same columns, widths and number formats as the built-in ones, and
its currency-bearing headers are rewritten to match the currency chosen.

Names are checked against Excel's own rules — 31 characters, no `[ ] : * ? / \`
— and against `Summary` and `Catalogue`, which the workbook needs for itself.

### New prefixes

Prefixes the built-in rules do not reach are worked out from the catalogue and
listed on the fallback sheet, so the routing table shows every prefix the file
actually contains rather than sweeping the strays up silently. The routing
panel opens by itself when a catalogue brings some, and moving them to a sheet
of their own is then a matter of editing rows.

A prefix is the segment before the first dash when that is two or three
characters — `TE-1SNA179534R2200` gives `TE`, `PHX-3003347` gives `PHX`.
Anything else takes the first three characters, which covers a code with no
dash (`RSHXM07285081B000000` gives `RSH`) and one whose first segment runs long
(`ABDACS180-04S-17A0-4` gives `ABD`). Dashed and undashed spellings of the same
supplier therefore land on one prefix.

A code still matching no rule goes to the fallback sheet — `O (UE4)` while it
is listed, otherwise the last row — and is reported rather than dropped.
Individual rows can also be moved to another sheet in the review table.

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
=IF(<Qty> < <Stock>, <Landed USD>,
    IF(<U.P. Ex.> = "", "", <U.P. Ex.> * $freight [* EUR factor]))
```

Quantity decides the branch. While stock covers the quantity the row takes the
landed cost exactly as uploaded — what that stock actually cost to land. Once
Qty reaches or exceeds stock the order has to be imported, so it is priced off
the ex-works figure instead, grossed up by freight and the EUR factor on a EUR
sheet and by freight alone on the USD one — and its **Qty cell turns yellow**,
so the rows still waiting on an ex-works price are visible at a glance.

Such a row stays **blank** rather than showing 0, which would read like a
costed line worth nothing. `T. Landed` and `Margin` blank out with it, since
blank times a quantity is `#VALUE!` and a 0.00% margin on a cost nobody has
entered is worse than no figure at all. All three fill in together the moment
`U.P. Ex.` is keyed. The column totals ignore the blanks, so the footers and
the Summary stay correct while the sheet is part-priced.

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

---

# The Last Purchase Price tab

Upload a purchase-history export — one row per invoice line — and get back the
Order workbook with each item's three most recent unit prices already in place.

## What goes in

A `.xlsx`, `.xlsm`, `.xls` or `.csv` carrying an item code, a posting date and a
unit cost; a description is used when there is one. Headers are matched loosely,
so `No_` / `Item Code` / `Part No`, `InvLine_Posting_Date` / `Posting Date` /
`Date` and `InvLine_Unit_Cost` / `Unit Price` / `Cost` all read correctly. Lines
with no usable date or price are skipped and counted back to you.

## What comes out

Per item the lines are sorted newest first and the top three fill:

| Order column | Holds |
| --- | --- |
| `Last U.P.(1) (EUR)` / `Date(1)` | most recent purchase |
| `Last U.P.(2) (EUR)` / `Date(2)` | second most recent |
| `Last U.P.(3) (EUR)` / `Date(3)` | third most recent |

`Item Code` and `Description` are filled too. `Qty`, `U.P.`, `Agr.U.P.` and the
FOB columns are left for whoever is pricing, and the template's own formulas —
`Diff.UP/AGR`, `Diff. UP/LAST UP`, the totals — come down with every row, with
the `Total (EUR)` row re-pointed at the new block. Items bought fewer than three
times leave the unused slots empty.

## Two modes

**New Order from the template** builds a fresh Order listing every item the
history mentions, one row each, sorted by item code. It uses
`template/order_template.xlsx` unless you upload a template of your own.

**Fill an Order I already have** leaves your rows exactly as they are and writes
only the six price and date columns, matching on `Item Code`. Rows whose code
the history does not mention are left untouched and listed back to you.

## Repeated lines

Exports often carry the same invoice line twice. **Collapse repeated lines**, on
by default, counts rows sharing a date *and* a price as one purchase, so the
three slots hold three distinct purchases; turn it off to take the last three
rows as they come. Where two purchases share a date at different prices, the one
further down the file is taken as the more recent.

## Note on the template

In `Order.xlsx` the `Last U.P.(3) (EUR)` and `FOB U.P.(EUR) 2025` columns carry
a date number format, so a price typed into either shows as a date. The app
corrects this for the six columns it fills; `FOB U.P.(EUR) 2025` is left alone,
being outside what it writes.

---

# The Cost + Customer Pricing Tool tab

Fills a `Calcul`-layout offer sheet: what this customer last paid, what the
catalogue says about the item, and the landed cost that follows from both.

## What goes in

| Upload | Supplies |
| --- | --- |
| **Customer sales report** (required) | the three `Last U.P. / Cur. / Date` blocks |
| **Special Inquiry Worksheet** (required) | `Description`, `Stock AV (BC)`, `Landed usd (BC)` |
| **Calcul workbook** *or* **BOQ** | the offer lines themselves |

The sales report needs a posting date, an item no., a price and — to be useful —
a quantity and a currency. Headers are matched loosely, so both the RICH
MOTORS-style export (`OC Net Price`, `Currency`) and the plainer
`Unit Price Excl. VAT` / `Line Discount %` / `Currency Code` shape read
correctly. The Special Inquiry Worksheet is the same catalogue extract the
first tab reads, so the two tabs accept the same file.

## Which price lands in Last U.P.

The sidebar picks what the figure should *mean*, not which column it comes
from, because exports differ on the columns they carry:

* **Net of line discount** (the default) — the price as actually invoiced.
* **Before line discount** — the unit price with the discount still to come off.

Reports carrying a net price outright (`OC Net Price`, `Net Price`) are taken at
their word. One carrying only a unit price and a discount percentage
(`Unit Price Excl. VAT` + `Line Discount %`) has the discount applied here,
which reaches the same figure by another route. A discount column is read as a
percentage unless every non-zero entry is below 1, in which case it is a
fraction.

Where a report offers both an order-currency net price and a USD one, the
order-currency figure wins: that is the one that pairs with `Cur.`, and a USD
number sitting beside `EUR` would be a lie. The review step names the column it
settled on, so it is worth a glance on an unfamiliar export.

A negative quantity is a credit memo reversing an invoice, not a price the
customer paid, so those lines are skipped — **Include credit memos** overrides
that. **Collapse repeated lines** counts lines sharing a date, price *and*
currency as one sale. The same item sold twice on one date in two currencies is
two distinct sales and takes two slots.

## The BOQ box

Upload a BOQ in the sidebar and the offer is raised from its item codes and
quantities on the app's own Calcul template, so no Calcul workbook is needed.
Descriptions come from the catalogue. The totals block keeps its formulas and
slides to sit under the data, however many rows that turns out to be — above or
below the template's hundred.

Leave the box empty and you upload a Calcul workbook instead; its own figures,
`U.P. Ex.` included, are left exactly as they are.

## U. Landed

Rewritten to choose by quantity against the stock on hand:

```
=IF(Qty <= Stock AV, Landed usd, IF(D.U.P. Ex.="", "", D.U.P. Ex. * factors))
```

A quantity the stock covers takes the landed cost as uploaded; going past it
means the order has to be imported, so the row prices off the ex-works figure
instead. The gross-up factors are lifted from the sheet's own formula rather
than restated, so editing them in row 1 still reprices every row.

The `<=` matches the Qty colouring below exactly: a green cell always means the
catalogue's landed cost, a yellow one always means an import. The Pricing Tool
Converter tab draws that line at `<` instead, so an order for exactly the stock
on hand is treated as an import there and as covered here.

`U.P. Ex.` stays yours to key — every costed column follows from it.

## Uncosted rows stay blank

A row nobody has priced yet shows nothing, rather than a column of `0.00` and
`#DIV/0!` that reads like priced work which came out worthless. `U. Landed`,
`U.P. (USD)`, `Disc.`, `D.T.P. Ex.`, `T. Landed`, `Total` and `Margin` all keep
their formulas and fill themselves in the moment `U.P. Ex.` is keyed:

| | Qty vs Stock | U.P. Ex. | U. Landed | Disc. | Margin |
| --- | --- | --- | --- | --- | --- |
| covered | 10 ≤ 288 | — | 1.49 from the catalogue | blank | 0.75 |
| short | 9999 > 79 | — | blank | blank | blank |
| short, priced | 9999 > 44 | 100 | 126.26, filled on the spot | 0.00 | 0.75 |

Each cell watches the one it divides or multiplies by, because blank times a
number is `#VALUE!` — which would be no better than the zeros. The guards are
the same ones the Pricing Tool Converter's offer sheets use, and filling an
already-filled workbook rebuilds them rather than nesting them.

## Qty colours

The `Qty` cells are coloured against `Stock AV (BC)`: **green** while the stock
covers the order, **yellow** once the quantity is greater. These are rules
rather than painted fills, so they keep up as quantities are retyped in Excel.
A row with no stock figure stays uncoloured — no figure is not a stock of zero.

The colours and `U. Landed` share the one boundary, so the colour always tells
you which cost the row is on.

---

## Files

| File | Contents |
| --- | --- |
| `home.py` | the tab bar — the entry point |
| `app.py` | Pricing Tool Converter UI |
| `ui.py` | masthead, section headers, metric cards, sheet chips |
| `.streamlit/config.toml` | palette, type and radii for light and dark |
| `catalog.py` | reading and normalising the uploaded workbooks |
| `builder.py` | writing the offer sheets, footers and Summary |
| `config.py` | column maps for both layouts, routing defaults |
| `template/pricing_tool_template.xlsx` | the styled template |
| `last_purchase.py` | Last Purchase Price UI |
| `purchases.py` | reading the history, ranking prices, writing the Order |
| `template/order_template.xlsx` | the blank Order |
| `cost_customer.py` | Cost + Customer Pricing Tool UI |
| `pricing.py` | reading the sales report, ranking sales, writing Calcul |
| `template/calcul_template.xlsx` | the blank Calcul offer sheet |

`purchases.py` shares nothing with `builder.py` or `catalog.py`, and `pricing.py`
borrows only `catalog.py`'s readers — no tab can break another's code. Both also
run on their own:

```bash
python purchases.py "last purchases.xlsx" -o "Order.xlsx"
python purchases.py "last purchases.xlsx" --into "Order.xlsx" -o "Order filled.xlsx"
```

```bash
python pricing.py "RICH MOTORS REPORT.xlsx" "Special Inquiry Worksheet.xlsx" --into "Calcul.xlsx" -o "Calcul filled.xlsx"
python pricing.py "RICH MOTORS REPORT.xlsx" "Special Inquiry Worksheet.xlsx" --boq "BOQ.xlsx" -o "Calcul.xlsx"
```

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
