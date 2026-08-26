"""Reading and normalising the uploaded workbooks.

Two kinds of file are understood:

* a **catalogue** extract - one or more sheets laid out like ``Sheet7`` /
  ``Sheet9`` (item code, description, unit price, stock figures, landed cost);
* an optional **quantity** source - either the raw BOQ or the "sum of qty"
  pivot produced from it, used to fill the ``Qty`` column of the offer sheets.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import openpyxl

from config import CATALOG_HEADERS

_NUMERIC_FIELDS = (
    "unit_price", "advanced_reserved", "stock",
    "po_qty", "po_not_shipped", "landed_usd",
)

_FIELD_ALIASES = {
    "code": {"item no.1", "item no", "item no.", "item code", "part no.",
             "part no", "part number", "material", "row labels"},
    "description": {"description", "desc", "item description"},
    "unit_price": {"unit price", "list price", "price"},
    "advanced_reserved": {"advanced reserved", "advance reserved", "reserved"},
    "stock": {"stock available quantity", "stock available qty", "stock qty",
              "stock", "available quantity"},
    "po_qty": {"po qty", "po quantity", "purchase order qty"},
    "po_not_shipped": {"po not shipped", "po not-shipped", "not shipped"},
    "landed_usd": {"landed usd", "landed (usd)", "landed cost", "landed"},
}

_QTY_ALIASES = {"sum of qty", "sum of quantity", "qty", "quantity", "total qty"}


def _norm(value) -> str:
    """Lower-case, collapse whitespace - for tolerant header matching."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def _clean_code(value) -> str:
    """Item codes arrive with stray spaces (pivot tables add a leading one)."""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).upper()


def _number(value, default=0.0) -> float:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^\d.\-]", "", str(value))
    try:
        return float(text)
    except ValueError:
        return default


@dataclass
class CatalogItem:
    code: str
    description: str = ""
    unit_price: float = 0.0
    advanced_reserved: float = 0.0
    stock: float = 0.0
    po_qty: float = 0.0
    po_not_shipped: float = 0.0
    landed_usd: float = 0.0
    origin: str = ""          # which uploaded sheet the row came from

    def as_row(self) -> list:
        return [
            self.code, self.description, self.unit_price,
            self.advanced_reserved, self.stock, self.po_qty,
            self.po_not_shipped, self.landed_usd,
        ]


def _map_headers(header_row) -> dict:
    """Map ``field name -> 0-based column index`` for one candidate header row."""
    mapping = {}
    for idx, cell in enumerate(header_row):
        label = _norm(cell)
        if not label:
            continue
        for field, aliases in _FIELD_ALIASES.items():
            if label in aliases and field not in mapping:
                mapping[field] = idx
    return mapping


def _find_header_row(rows, required, scan=12, contiguous=False):
    """Locate the header row: catalogue exports sometimes carry a title above.

    With ``contiguous`` the code / description / unit-price headers must sit in
    three consecutive columns, which is what distinguishes a real ``Sheet7``
    style catalogue from an offer sheet (whose header row uses the same words
    but spreads them across columns B, C and E).
    """
    for row_idx, row in enumerate(rows[:scan]):
        mapping = _map_headers(row)
        if not all(field in mapping for field in required):
            continue
        if contiguous:
            code = mapping["code"]
            if mapping["description"] != code + 1 or mapping["unit_price"] != code + 2:
                continue
        return row_idx, mapping
    return None, None


def read_catalog(file_like) -> tuple[list[CatalogItem], list[str]]:
    """Return every catalogue row found in an uploaded workbook.

    Any worksheet carrying an item-code + description + unit-price header is
    treated as a catalogue sheet, so a file holding both a ``Sheet7``-style and
    a ``Sheet9``-style tab is read in one pass.  Later duplicates of a code are
    ignored, and the winning row keeps whichever landed cost is non-zero.
    """
    workbook = openpyxl.load_workbook(file_like, data_only=True, read_only=True)
    items: dict[str, CatalogItem] = {}
    skipped: list[str] = []

    sheets = [(s.title, [list(r) for r in s.iter_rows(values_only=True)])
              for s in workbook.worksheets]
    required = ("code", "description", "unit_price")
    # Prefer the strict Sheet7/Sheet9 shape; only fall back to a loose match if
    # nothing in the file looks like a proper catalogue.
    strict = any(_find_header_row(rows, required, contiguous=True)[1]
                 for _, rows in sheets if rows)

    for title, rows in sheets:
        if not rows:
            continue
        header_idx, mapping = _find_header_row(
            rows, required, contiguous=strict)
        if mapping is None:
            skipped.append(title)
            continue

        for raw in rows[header_idx + 1:]:
            code = _clean_code(raw[mapping["code"]] if mapping["code"] < len(raw) else None)
            if not code or _norm(code) in {_norm(h) for h in CATALOG_HEADERS}:
                continue
            values = {}
            for field, col in mapping.items():
                cell = raw[col] if col < len(raw) else None
                values[field] = _number(cell) if field in _NUMERIC_FIELDS else cell
            item = CatalogItem(
                code=code,
                description=str(values.get("description") or "").strip(),
                unit_price=values.get("unit_price", 0.0),
                advanced_reserved=values.get("advanced_reserved", 0.0),
                stock=values.get("stock", 0.0),
                po_qty=values.get("po_qty", 0.0),
                po_not_shipped=values.get("po_not_shipped", 0.0),
                landed_usd=values.get("landed_usd", 0.0),
                origin=title,
            )
            existing = items.get(code)
            if existing is None:
                items[code] = item
            elif not existing.landed_usd and item.landed_usd:
                items[code] = item

    workbook.close()
    return list(items.values()), skipped


def read_quantities(file_like) -> dict[str, dict[str, float]]:
    """Return ``{worksheet: {item code: total qty}}`` for a BOQ workbook.

    Both the raw BOQ and the "sum of qty" pivot built from it usually live in
    the same file and describe the same quantities, so they are kept apart per
    sheet rather than merged - the caller picks one.  Within a sheet a repeated
    code is summed, which is how the same part on several panels adds up.
    """
    workbook = openpyxl.load_workbook(file_like, data_only=True, read_only=True)
    per_sheet: dict[str, dict[str, float]] = {}

    for sheet in workbook.worksheets:
        rows = [list(r) for r in sheet.iter_rows(values_only=True)]
        if not rows:
            continue

        header_idx = qty_col = code_col = None
        for row_idx, row in enumerate(rows[:12]):
            labels = [_norm(c) for c in row]
            candidate_qty = next(
                (i for i, lab in enumerate(labels) if lab in _QTY_ALIASES), None)
            mapping = _map_headers(row)
            if candidate_qty is not None and "code" in mapping:
                header_idx, qty_col, code_col = row_idx, candidate_qty, mapping["code"]
                break
        if header_idx is None:
            continue

        quantities: dict[str, float] = {}
        for raw in rows[header_idx + 1:]:
            code = _clean_code(raw[code_col] if code_col < len(raw) else None)
            if not code or code.startswith("GRANDTOTAL"):
                continue
            qty = _number(raw[qty_col] if qty_col < len(raw) else None)
            if qty:
                quantities[code] = quantities.get(code, 0.0) + qty
        if quantities:
            per_sheet[sheet.title] = quantities

    workbook.close()
    return per_sheet


def best_quantity_sheet(per_sheet: dict[str, dict[str, float]]) -> str | None:
    """Pick the most likely quantity sheet: the one covering the most codes."""
    if not per_sheet:
        return None
    return max(per_sheet, key=lambda name: len(per_sheet[name]))
