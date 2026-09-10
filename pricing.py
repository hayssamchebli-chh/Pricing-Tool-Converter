"""Cost + customer pricing: a customer sales history becomes the Calcul sheet.

Three sources feed one offer sheet laid out like ``Calcul.xlsx``:

* a **sales report** (the RICH MOTORS export) fills the three
  ``Last U.P. / Cur. / Date`` blocks, most recent sale first;
* a **catalogue extract** — the Special Inquiry Worksheet, read by
  :mod:`catalog` — fills ``Description`` and the two Business Central
  reference columns, ``Stock AV (BC)`` and ``Landed usd (BC)``;
* an optional **BOQ** raises the sheet from nothing when there is no Calcul
  file to start from, taking item codes and quantities from it.

``U. Landed (USD)`` is rewritten to the same condition the offer sheets use:
stock covering the quantity takes the landed cost as uploaded, otherwise the
row prices off the ex-works figure grossed up by the template's own factors.
"""

from __future__ import annotations

import io
import re
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

TEMPLATE_PATH = Path(__file__).with_name("template") / "calcul_template.xlsx"

MAX_HISTORY = 3

# Which of the report's three price columns fills "Last U.P.". Only the order
# currency pairs with the "Cur." column: the other two are USD whatever the
# sale was struck in, so a EUR row would read as a USD figure labelled EUR.
PRICE_FIELDS = {
    "oc_net": "OC Net Price (as invoiced, with its currency)",
    "net": "Net Price (USD)",
    "pre_discount": "Pre-Discount Price (USD)",
}
DEFAULT_PRICE_FIELD = "oc_net"

SALES_ALIASES = {
    "item": ["no", "no1", "itemno", "itemno1", "itemcode", "item", "code",
             "partno", "partnumber"],
    "description": ["description", "desc", "itemdescription", "itemname"],
    "date": ["postingdate", "date", "invoicedate", "documentdate",
             "postingdt"],
    "oc_net": ["ocnetprice", "ocnet", "ordercurrencynetprice"],
    "net": ["netprice", "net"],
    "pre_discount": ["prediscountpriceusd", "prediscountprice", "prediscount",
                     "grossprice", "listprice"],
    "currency": ["currency", "cur", "curr", "ccy"],
    "quantity": ["quantity", "qty"],
    "document": ["documentno", "document", "invoiceno", "docno"],
}

# Header text of the Calcul sheet, normalised. "U.P.\n(USD)" -> "upusd".
CALCUL_TARGETS = {
    "num": ["#", "no", "line"],
    "code": ["itemcode", "itemno", "code"],
    "description": ["description"],
    "qty": ["qty", "quantity"],
    "unit_price": ["upusd", "up"],
    "price1": ["lastup1"], "cur1": ["cur1"], "date1": ["date1"],
    "price2": ["lastup2"], "cur2": ["cur2"], "date2": ["date2"],
    "price3": ["lastup3"], "cur3": ["cur3"], "date3": ["date3"],
    "landed": ["ulandedusd", "ulanded"],
    "ex_works": ["upexeur", "upex"],
    "disc_unit": ["dupexeur", "dupex"],
    "stock": ["stockavbc", "stockav", "stockavailablequantity"],
    "landed_ref": ["landedusdbc", "landedusd"],
}

PRICE_FORMAT = "#,##0.00"
DATE_FORMAT = "mm-dd-yy"
TEXT_FORMAT = "General"


class BuildError(Exception):
    """Raised when an uploaded file cannot be interpreted."""


def norm(value):
    """Normalise a header: 'Last U.P.\\n(1)' -> 'lastup1'."""
    return re.sub(r"[^a-z0-9]", "", str(value).lower()) if value is not None else ""


def _resolve(columns, aliases):
    """Map logical names to actual labels, preferring earlier aliases."""
    normalised = {}
    for col in columns:
        normalised.setdefault(norm(col), col)
    resolved = {}
    for logical, names in aliases.items():
        for name in names:
            if name in normalised:
                resolved[logical] = normalised[name]
                break
    return resolved


# --- the sales report ------------------------------------------------------


def read_sales(source, sheet_name=0, price_field=DEFAULT_PRICE_FIELD):
    """Load a customer sales report into a tidy frame.

    Returns one row per invoice line: item, description, date, price, currency,
    quantity and document number.
    """
    if price_field not in PRICE_FIELDS:
        raise BuildError("Unknown price column {!r}.".format(price_field))

    name = getattr(source, "name", str(source)).lower()
    if name.endswith(".csv") or name.endswith(".txt"):
        frame = pd.read_csv(source, sep=None, engine="python")
    else:
        frame = pd.read_excel(source, sheet_name=sheet_name)

    columns = _resolve(frame.columns, SALES_ALIASES)
    # The chosen price column, falling back through the others so a report
    # missing one still builds rather than refusing outright.
    for candidate in (price_field, "net", "oc_net", "pre_discount"):
        if candidate in columns:
            price_column, price_used = columns[candidate], candidate
            break
    else:
        price_column = price_used = None

    missing = [key for key in ("item", "date") if key not in columns]
    if missing or price_column is None:
        missing = missing + ([] if price_column else ["price"])
        raise BuildError(
            "Could not find the " + ", ".join(missing) + " column(s) in the "
            "sales report. Columns found: "
            + ", ".join(str(c) for c in frame.columns)
        )

    tidy = pd.DataFrame(
        {
            "item": frame[columns["item"]],
            "description": frame[columns["description"]] if "description" in columns else "",
            "date": pd.to_datetime(frame[columns["date"]], errors="coerce"),
            "price": pd.to_numeric(frame[price_column], errors="coerce"),
            "currency": frame[columns["currency"]] if "currency" in columns else "",
            "quantity": (
                pd.to_numeric(frame[columns["quantity"]], errors="coerce")
                if "quantity" in columns else 1.0
            ),
            "document": frame[columns["document"]] if "document" in columns else "",
        }
    )
    tidy["item"] = tidy["item"].astype(str).str.strip()
    tidy["description"] = tidy["description"].fillna("").astype(str).str.strip()
    tidy["currency"] = tidy["currency"].fillna("").astype(str).str.strip().str.upper()
    tidy["document"] = tidy["document"].fillna("").astype(str).str.strip()
    tidy["key"] = tidy["item"].str.replace(r"\s+", "", regex=True).str.upper()
    tidy["source_row"] = range(len(tidy))
    tidy.attrs["price_used"] = price_used
    keep = (tidy["item"] != "") & (tidy["item"].str.lower() != "nan")
    result = tidy.loc[keep].reset_index(drop=True)
    result.attrs["price_used"] = price_used
    return result


@dataclass
class SaleHistory:
    """The three most recent sales of one item, newest first."""

    item: str
    description: str = ""
    prices: list = field(default_factory=list)
    currencies: list = field(default_factory=list)
    dates: list = field(default_factory=list)
    total_sales: int = 0

    def cell(self, index):
        """(price, currency, date) at `index` (0 = most recent), or blanks."""
        if index < len(self.prices):
            return self.prices[index], self.currencies[index], self.dates[index]
        return None, None, None


def build_histories(sales, collapse_duplicates=True, include_returns=False):
    """Per item, the `MAX_HISTORY` most recent sales, newest first.

    `include_returns` keeps negative-quantity lines — credit memos, which pair
    with the invoice they reverse and are not a price the customer paid.
    `collapse_duplicates` folds lines sharing a date, price and currency into
    one, so the three slots hold three distinct sales.
    """
    histories = {}
    frame = sales if include_returns else sales.loc[sales["quantity"].fillna(1) > 0]
    ordered = frame.sort_values(
        ["date", "source_row"], ascending=[False, False], na_position="last"
    )
    for key, group in ordered.groupby("key", sort=False):
        rows = group.dropna(subset=["date", "price"])
        if collapse_duplicates:
            rows = rows.drop_duplicates(subset=["date", "price", "currency"])
        head = rows.head(MAX_HISTORY)
        described = group["description"].loc[group["description"] != ""]
        histories[key] = SaleHistory(
            item=group["item"].iloc[0],
            description=described.iloc[0] if len(described) else "",
            prices=[round(float(price), 4) for price in head["price"]],
            currencies=list(head["currency"]),
            dates=[when.to_pydatetime() for when in head["date"]],
            total_sales=len(rows),
        )
    return histories


def catalog_index(items):
    """Catalogue rows keyed by cleaned item code, for lookup by the sheet."""
    return {re.sub(r"\s+", "", item.code).upper(): item for item in items}


# --- the Calcul sheet ------------------------------------------------------


def _find_header_row(worksheet, wanted="itemcode", limit=12):
    for row in range(1, min(limit, worksheet.max_row) + 1):
        for col in range(1, worksheet.max_column + 1):
            if norm(worksheet.cell(row, col).value) == wanted:
                return row
    raise BuildError("No 'Item Code' header found in the Calcul sheet.")


def _map_columns(worksheet, header_row):
    labels = {}
    for col in range(1, worksheet.max_column + 1):
        value = worksheet.cell(header_row, col).value
        if value is not None:
            labels[value] = col
    resolved = _resolve(labels, CALCUL_TARGETS)
    columns = {logical: labels[label] for logical, label in resolved.items()}
    # The line-number header is "#", which normalises to nothing at all, so it
    # is matched on its raw text instead of through the alias table.
    if "num" not in columns:
        for label, col in labels.items():
            if str(label).strip() == "#":
                columns["num"] = col
                break
    return columns


def _find_footer_row(worksheet, header_row):
    """First row of the totals block — the row whose label starts 'Total'."""
    for row in range(header_row + 1, worksheet.max_row + 1):
        for col in range(1, worksheet.max_column + 1):
            if norm(worksheet.cell(row, col).value).startswith("total"):
                return row
    return None


_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _shift_rows(formula, at_or_after, delta):
    """Move every reference at or below `at_or_after` down by `delta` rows.

    Used when the totals block slides because the data block grew or shrank:
    the margin column's ``$D$104`` and the net-total's ``U103`` have to follow
    it, and openpyxl's translator leaves absolute references where they are.
    """
    if not delta:
        return formula

    def move(match):
        col_abs, col, row_abs, row = match.groups()
        number = int(row)
        if number >= at_or_after:
            number += delta
        return "{}{}{}{}".format(col_abs, col, row_abs, number)

    return _REF.sub(move, formula)


def _rewrite_sum(formula, first_row, last_row):
    """Point a totals SUM at the new data block: SUM(U3:U102) -> SUM(U3:U31)."""
    return re.sub(
        r"(SUM\()([A-Z]{1,3})\d+:([A-Z]{1,3})\d+(\))",
        lambda m: "{}{}{}:{}{}{}".format(
            m.group(1), m.group(2), first_row, m.group(3), last_row, m.group(4)),
        formula,
        flags=re.IGNORECASE,
    )


def _replay(formula, col, proto_row, row, footer_row, delta):
    """Move a prototype data-row formula to `row`, following the totals block.

    The footer shift has to happen **before** the translation, while a
    reference into the totals block is still recognisable by its row number.
    Afterwards it is not: on a sheet grown past its template, a data row's own
    ``E152`` sits where the old ``$D$104`` block used to be, and shifting last
    would drag the row's own references along with it.
    """
    shifted = _shift_rows(formula, footer_row, delta)
    letter = get_column_letter(col)
    return Translator(
        shifted, origin="{}{}".format(letter, proto_row)
    ).translate_formula("{}{}".format(letter, row))


def _capture_row(worksheet, row):
    """Snapshot a row's styles and formulas so it can be replayed elsewhere."""
    return {
        "height": worksheet.row_dimensions[row].height,
        "cells": {
            col: (copy(worksheet.cell(row, col)._style), worksheet.cell(row, col).value)
            for col in range(1, worksheet.max_column + 1)
        },
    }


def _landed_formula(columns, row, gross_up):
    """``U. Landed`` under the offer sheets' own stock condition.

    A quantity the stock covers takes the landed cost as uploaded; reaching it
    means the order has to be imported, so the row prices off the ex-works
    figure grossed up by whatever factors the template already applies —
    lifted from its own formula rather than restated here.
    """
    needed = ("qty", "stock", "landed_ref", "disc_unit")
    if not all(key in columns for key in needed):
        return None
    at = lambda key: "{}{}".format(get_column_letter(columns[key]), row)
    return '=IF({qty}<{stock},{ref},IF({ex}="","",{gross}))'.format(
        qty=at("qty"), stock=at("stock"), ref=at("landed_ref"),
        ex=at("disc_unit"), gross=gross_up)


def _write_history(worksheet, row, columns, history):
    """Fill the nine Last U.P. / Cur. / Date cells, newest sale first."""
    for slot in range(MAX_HISTORY):
        price, currency, when = history.cell(slot) if history else (None, None, None)
        for key, value, fmt in (
            ("price{}".format(slot + 1), price, PRICE_FORMAT),
            ("cur{}".format(slot + 1), currency, TEXT_FORMAT),
            ("date{}".format(slot + 1), when, DATE_FORMAT),
        ):
            col = columns.get(key)
            if not col:
                continue
            cell = worksheet.cell(row, col)
            cell.number_format = fmt
            cell.value = value if value not in ("", None) else None


def _write_catalog(worksheet, row, columns, item):
    """Description and the two Business Central reference columns."""
    if "description" in columns:
        worksheet.cell(row, columns["description"]).value = (
            item.description if item else None)
    if "stock" in columns:
        worksheet.cell(row, columns["stock"]).value = item.stock if item else None
    if "landed_ref" in columns:
        worksheet.cell(row, columns["landed_ref"]).value = (
            item.landed_usd if item else None)


def _fill_row(worksheet, row, columns, code, histories, catalog, gross_up):
    """Everything this app owns on one offer row."""
    history = histories.get(code.upper()) if code else None
    item = catalog.get(code.upper()) if code else None
    _write_history(worksheet, row, columns, history)
    _write_catalog(worksheet, row, columns, item)
    if gross_up and "landed" in columns:
        formula = _landed_formula(columns, row, gross_up)
        if formula:
            worksheet.cell(row, columns["landed"]).value = formula
    return history is not None, item is not None


def _gross_up_for(worksheet, columns, proto_row, row):
    """The template's own ex-works gross-up, moved to `row`.

    ``=Q3*$O$1*$B$1`` on the prototype row becomes ``Q7*$O$1*$B$1`` — the
    expression without its ``=``, ready to drop into the landed IF.
    """
    if "landed" not in columns:
        return None
    original = worksheet.cell(proto_row, columns["landed"]).value
    if not isinstance(original, str) or not original.startswith("="):
        return None
    moved = Translator(
        original, origin="{}{}".format(get_column_letter(columns["landed"]), proto_row)
    ).translate_formula("{}{}".format(get_column_letter(columns["landed"]), row))
    return moved.lstrip("=")


def read_calcul_lines(source):
    """The (code, qty, description) lines an existing Calcul workbook lists."""
    workbook = load_workbook(source, data_only=True, read_only=False)
    worksheet = workbook.active

    header_row = _find_header_row(worksheet)
    columns = _map_columns(worksheet, header_row)
    if "code" not in columns:
        raise BuildError("The Calcul sheet has no 'Item Code' column.")

    footer_row = _find_footer_row(worksheet, header_row)
    last_row = (footer_row - 1) if footer_row else worksheet.max_row

    lines = []
    for row in range(header_row + 1, last_row + 1):
        raw = worksheet.cell(row, columns["code"]).value
        code = re.sub(r"\s+", "", str(raw)).upper() if raw is not None else ""
        if not code or code.lower() == "none":
            continue
        qty = worksheet.cell(row, columns["qty"]).value if "qty" in columns else None
        description = (worksheet.cell(row, columns["description"]).value
                       if "description" in columns else "")
        lines.append((code, qty, "" if description is None else str(description)))
    return lines


def fill_calcul(source, histories, catalog, price_field=DEFAULT_PRICE_FIELD):
    """Fill an existing Calcul workbook in place, leaving its own figures alone.

    Returns (bytes, matched codes, codes with no sales history, codes not in
    the catalogue).
    """
    workbook = load_workbook(source)
    worksheet = workbook.active

    header_row = _find_header_row(worksheet)
    columns = _map_columns(worksheet, header_row)
    if "code" not in columns:
        raise BuildError("The Calcul sheet has no 'Item Code' column.")

    footer_row = _find_footer_row(worksheet, header_row)
    last_row = (footer_row - 1) if footer_row else worksheet.max_row
    first_data_row = header_row + 1

    matched, no_history, no_catalog = [], [], []
    for row in range(first_data_row, last_row + 1):
        raw = worksheet.cell(row, columns["code"]).value
        code = re.sub(r"\s+", "", str(raw)).upper() if raw is not None else ""
        if not code or code.lower() == "none":
            continue
        gross_up = _gross_up_for(worksheet, columns, row, row)
        had_history, had_item = _fill_row(
            worksheet, row, columns, code, histories, catalog, gross_up)
        matched.append(code)
        if not had_history:
            no_history.append(code)
        if not had_item:
            no_catalog.append(code)

    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue(), matched, no_history, no_catalog


def build_calcul(lines, histories, catalog, template=None):
    """Raise a Calcul workbook from BOQ lines: (code, qty) pairs.

    The totals block keeps its formulas and slides to sit under the data,
    however many rows that turns out to be.
    """
    workbook = load_workbook(template or TEMPLATE_PATH)
    worksheet = workbook.active

    header_row = _find_header_row(worksheet)
    columns = _map_columns(worksheet, header_row)
    if "code" not in columns:
        raise BuildError("The Calcul template has no 'Item Code' column.")

    first_data_row = header_row + 1
    footer_row = _find_footer_row(worksheet, header_row)
    if footer_row is None:
        raise BuildError("The Calcul template has no totals row.")

    proto = _capture_row(worksheet, first_data_row)
    footer = [_capture_row(worksheet, row)
              for row in range(footer_row, worksheet.max_row + 1)]
    template_rows = footer_row - first_data_row
    count = max(len(lines), 1)
    delta = count - template_rows
    new_footer_row = footer_row + delta
    last_data_row = first_data_row + count - 1

    # Grow first so the footer never lands on a row that is about to be written.
    if delta > 0:
        worksheet.insert_rows(footer_row, delta)
    elif delta < 0:
        worksheet.delete_rows(first_data_row + count, -delta)

    gross_up_proto = _gross_up_for(worksheet, columns, first_data_row, first_data_row)

    for offset, (code, qty, description) in enumerate(lines):
        row = first_data_row + offset
        for col, (style, value) in proto["cells"].items():
            cell = worksheet.cell(row, col)
            cell._style = copy(style)
            if isinstance(value, str) and value.startswith("="):
                cell.value = _replay(value, col, first_data_row, row,
                                     footer_row, delta)
            else:
                cell.value = None
        if proto["height"]:
            worksheet.row_dimensions[row].height = proto["height"]

        if "num" in columns:
            worksheet.cell(row, columns["num"]).value = offset + 1
        worksheet.cell(row, columns["code"]).value = code
        if "qty" in columns:
            worksheet.cell(row, columns["qty"]).value = qty
        gross_up = (
            _replay("=" + gross_up_proto, columns["landed"], first_data_row,
                    row, footer_row, delta).lstrip("=")
            if gross_up_proto else None)
        _fill_row(worksheet, row, columns, code, histories, catalog, gross_up)
        item = catalog.get(code.upper())
        if "description" in columns and not (item and item.description) and description:
            worksheet.cell(row, columns["description"]).value = description

    for index, captured in enumerate(footer):
        row = new_footer_row + index
        if captured["height"]:
            worksheet.row_dimensions[row].height = captured["height"]
        for col, (style, value) in captured["cells"].items():
            cell = worksheet.cell(row, col)
            cell._style = copy(style)
            if isinstance(value, str) and value.startswith("="):
                shifted = _shift_rows(value, footer_row, delta)
                cell.value = _rewrite_sum(shifted, first_data_row, last_data_row)
            else:
                cell.value = value

    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def histories_to_frame(codes, histories, catalog):
    """The preview table: one row per offer line, with what was found for it."""
    rows = []
    for index, (code, qty, description) in enumerate(codes, start=1):
        key = code.upper()
        history = histories.get(key)
        item = catalog.get(key)
        row = {
            "#": index,
            "Item Code": code,
            "Description": (item.description if item else "") or description,
            "Qty": qty,
        }
        for slot in range(MAX_HISTORY):
            price, currency, when = (
                history.cell(slot) if history else (None, None, None))
            row["Last U.P.({})".format(slot + 1)] = price
            row["Cur.({})".format(slot + 1)] = currency
            row["Date({})".format(slot + 1)] = when.date() if when else None
        row["Stock AV"] = item.stock if item else None
        row["Landed usd"] = item.landed_usd if item else None
        rows.append(row)
    return pd.DataFrame(rows)


def main(argv=None):
    import argparse

    from catalog import best_quantity_sheet, read_catalog, read_quantities

    parser = argparse.ArgumentParser(description="Build the Calcul offer sheet.")
    parser.add_argument("sales", help="customer sales report (RICH MOTORS export)")
    parser.add_argument("catalogue", help="Special Inquiry Worksheet extract")
    parser.add_argument("-o", "--output", default="Calcul (generated).xlsx")
    parser.add_argument("--into", default=None, help="fill this Calcul workbook")
    parser.add_argument("--boq", default=None, help="build from this BOQ instead")
    parser.add_argument("--keep-returns", action="store_true",
                        help="keep negative-quantity credit memo lines")
    args = parser.parse_args(argv)

    histories = build_histories(read_sales(args.sales),
                                include_returns=args.keep_returns)
    items, _ = read_catalog(args.catalogue)
    catalog = catalog_index(items)

    if args.into:
        data, matched, no_history, no_catalog = fill_calcul(
            args.into, histories, catalog)
        print("Filled {} row(s); {} without sales history, {} not in the "
              "catalogue.".format(len(matched), len(no_history), len(no_catalog)))
    elif args.boq:
        per_sheet = read_quantities(args.boq, tuple(catalog))
        if not per_sheet:
            raise SystemExit("No quantity column found in that BOQ.")
        chosen = per_sheet[best_quantity_sheet(per_sheet)]
        lines = [(code, qty, "") for code, qty in sorted(chosen.items())]
        data = build_calcul(lines, histories, catalog)
        print("Built {} row(s) from the BOQ.".format(len(lines)))
    else:
        raise SystemExit("Pass --into a Calcul workbook or --boq a BOQ.")

    Path(args.output).write_bytes(data)
    print("Saved {}".format(args.output))


if __name__ == "__main__":
    main()
