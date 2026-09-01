"""Streamlit front end: catalogue extract -> six offer sheets.

Upload a supplier catalogue laid out like ``Sheet7`` / ``Sheet9`` (item code,
description, unit price, stock figures, landed cost), optionally a BOQ for the
quantities, and download a pricing workbook whose six offer sheets and Summary
are built and cross-referenced exactly like the reference file.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

import ui
from builder import BuildOptions, assign_sheets, build_workbook
from catalog import best_quantity_sheet, read_catalog, read_quantities
from config import (
    DEFAULT_EUR_FACTOR, DEFAULT_FREIGHT_FACTOR, DEFAULT_SPECS, SheetSpec,
    VAT_RATE,
)

st.set_page_config(page_title="Pricing Tool Converter", page_icon="📊",
                   layout="wide")
ui.inject_theme()

MIME_XLSX = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")

# Read off the specs rather than importing the constant: Streamlit reruns
# app.py without reloading already-imported modules, so a freshly added name
# in config would fail here until the app is rebooted.
LOOKUP_TABLES = sorted({spec.source for spec in DEFAULT_SPECS})


# --------------------------------------------------------------------------- #
# cached parsing
# --------------------------------------------------------------------------- #

@st.cache_data(show_spinner=False)
def _load_catalog(payload: bytes):
    return read_catalog(io.BytesIO(payload))


@st.cache_data(show_spinner=False)
def _load_quantities(payload: bytes):
    return read_quantities(io.BytesIO(payload))


def _specs_from_editor(frame: pd.DataFrame) -> list[SheetSpec]:
    specs = []
    for base in DEFAULT_SPECS:
        row = frame[frame["Sheet"] == base.sheet]
        if row.empty:
            specs.append(base)
            continue
        row = row.iloc[0]
        prefixes = [p.strip().upper()
                    for p in str(row["Item code prefixes"]).split(",") if p.strip()]
        specs.append(SheetSpec(base.sheet, base.layout, base.source,
                               str(row["Currency"]), prefixes))
    return specs


# --------------------------------------------------------------------------- #
# sidebar
# --------------------------------------------------------------------------- #

st.sidebar.header("Costing assumptions")
st.sidebar.caption(
    "Every *U. Landed* cell reads the catalogue cost until an ex-works price "
    "is keyed into *U.P. Ex.*, at which point it grosses that up by these "
    "factors instead — EUR sheets by both, USD sheets by freight alone."
)
freight_factor = st.sidebar.number_input(
    "Freight / clearing factor", value=DEFAULT_FREIGHT_FACTOR,
    min_value=1.0, max_value=3.0, step=0.01, format="%.3f")
eur_factor = st.sidebar.number_input(
    "EUR conversion factor", value=DEFAULT_EUR_FACTOR,
    min_value=0.5, max_value=3.0, step=0.01, format="%.3f")
vat_rate = st.sidebar.number_input(
    "VAT rate", value=VAT_RATE, min_value=0.0, max_value=0.5,
    step=0.01, format="%.3f")


# --------------------------------------------------------------------------- #
# uploads
# --------------------------------------------------------------------------- #

ui.masthead(
    "Pricing Tool Converter",

)

ui.section(1, "Source files", "the catalogue is required, quantities optional")

left, right = st.columns(2, gap="medium")
with left:
    with st.container(border=True):
        catalog_file = st.file_uploader(
            "Catalogue extract", type=["xlsx", "xlsm"],
            help="Item No.1 · Description · Unit Price · Advanced Reserved · Stock "
                 "Available Quantity · PO Qty · PO not Shipped · Landed USD")
        st.caption("Item code, description and unit price in three consecutive "
                   "columns is enough to be recognised.")
with right:
    with st.container(border=True):
        boq_file = st.file_uploader(
            "Quantities — BOQ or its pivot", type=["xlsx", "xlsm"],
            help="Any sheet pairing an item code column with a quantity column.")
        st.caption("Leave this empty to start every quantity at 0 and type them "
                   "in yourself.")

if catalog_file is None:
    st.info("Upload a catalogue extract to start.")
    st.stop()

items, skipped = _load_catalog(catalog_file.getvalue())
if not items:
    st.error(
        "No catalogue rows found. Expected a sheet whose header row carries "
        "an item code, a description and a unit price in three consecutive "
        "columns."
    )
    st.stop()

quantities: dict[str, float] = {}
if boq_file is not None:
    per_sheet = _load_quantities(boq_file.getvalue())
    if not per_sheet:
        st.warning("No quantity columns found in that file — quantities left at 0.")
    else:
        default = best_quantity_sheet(per_sheet)
        names = list(per_sheet)
        chosen = st.selectbox(
            "Quantity sheet", names, index=names.index(default),
            help="The raw BOQ and its pivot usually both appear here and carry "
                 "the same figures — pick either.")
        quantities = per_sheet[chosen]
        st.caption(
            "{} codes read from **{}**.".format(len(quantities), chosen))

# Filtering on quantity only means anything once a quantity source is loaded;
# without one every quantity is 0 and the filter would empty the offer.
st.sidebar.divider()
st.sidebar.header("Scope")
qty_only = st.sidebar.checkbox(
    "Only items with a quantity", value=bool(quantities),
    disabled=not quantities,
    help="Off puts every catalogue item on the offer, quantity 0.")
if not quantities:
    st.sidebar.caption(
        "No quantities loaded, so every catalogue item goes on the offer at "
        "quantity 0 — type them into the table below, or into the Qty column "
        "of the workbook afterwards."
    )


# --------------------------------------------------------------------------- #
# routing rules
# --------------------------------------------------------------------------- #

ui.section(2, "Review", "check the routing, then set quantities")

with st.expander("Sheet routing and discounts", expanded=False):
    st.caption(
        "Each offer sheet holds one supplier or product family, matched on the "
        "item-code prefix; the longest matching prefix wins. Every sheet reads "
        "its descriptions and prices from the **{}** table.".format(
            "** / **".join(LOOKUP_TABLES))
    )
    rules_frame = st.data_editor(
        pd.DataFrame([
            {
                "Sheet": spec.sheet,
                "Item code prefixes": ", ".join(spec.prefixes),
                "Currency": spec.currency,
                "Discount": 0.0,
            }
            for spec in DEFAULT_SPECS
        ]),
        column_config={
            "Sheet": st.column_config.TextColumn(disabled=True),
            "Item code prefixes": st.column_config.TextColumn(width="large"),
            "Currency": st.column_config.SelectboxColumn(options=["EUR", "USD"]),
            "Discount": st.column_config.NumberColumn(
                format="%.2f", min_value=0.0, max_value=0.95,
                help="Written into the sheet's Discount cell."),
        },
        hide_index=True, width="stretch", key="rules",
    )

specs = _specs_from_editor(rules_frame)
discounts = dict(zip(rules_frame["Sheet"], rules_frame["Discount"]))
assignment = assign_sheets([item.code for item in items], specs)


# --------------------------------------------------------------------------- #
# review and edit
# --------------------------------------------------------------------------- #

frame = pd.DataFrame([
    {
        "Item Code": item.code,
        "Description": item.description,
        "Sheet": assignment[item.code],
        "Qty": float(quantities.get(item.code, 0.0)),
        "Unit Price": item.unit_price,
        "Landed USD": item.landed_usd,
        "Stock": item.stock,
        "Source": item.origin,
    }
    for item in items
]).sort_values(["Sheet", "Item Code"], ignore_index=True)

if qty_only:
    frame = frame[frame["Qty"] > 0].reset_index(drop=True)

missing_qty = [code for code in quantities
               if code not in {item.code for item in items}]
no_landed = int((frame["Landed USD"] == 0).sum())

metrics = st.columns(4)
metrics[0].metric("Catalogue items", len(items))
metrics[1].metric("On the offer", len(frame))
metrics[2].metric("No landed cost", no_landed)
metrics[3].metric("Skipped sheets", len(skipped))

if frame.empty:
    st.error(
        "Every one of the {} catalogue items was filtered out, so the workbook "
        "would come out empty. **Only items with a quantity** is on but none of "
        "them has a quantity — either upload a BOQ, or turn that off in the "
        "sidebar to put the whole catalogue on the offer at quantity 0.".format(
            len(items))
    )
    st.stop()

if skipped:
    st.caption("Ignored, no catalogue header: " + ", ".join(skipped))
if missing_qty:
    with st.expander("{} quantity codes not in the catalogue".format(len(missing_qty))):
        st.write(", ".join(sorted(missing_qty)))
if no_landed:
    st.info(
        "{} item(s) carry no landed cost in the catalogue and will price at 0 "
        "until an ex-works price is keyed into their *U.P. Ex.* cell.".format(
            no_landed)
    )

st.caption("Quantities and sheet assignments are editable here before building.")
edited = st.data_editor(
    frame,
    column_config={
        "Item Code": st.column_config.TextColumn(disabled=True),
        "Description": st.column_config.TextColumn(disabled=True, width="large"),
        "Sheet": st.column_config.SelectboxColumn(
            options=[spec.sheet for spec in specs]),
        "Qty": st.column_config.NumberColumn(format="%.0f", min_value=0.0),
        "Unit Price": st.column_config.NumberColumn(format="%.2f", disabled=True),
        "Landed USD": st.column_config.NumberColumn(format="%.2f", disabled=True),
        "Stock": st.column_config.NumberColumn(format="%.0f", disabled=True),
        "Source": st.column_config.TextColumn(disabled=True),
    },
    hide_index=True, width="stretch", height=380, key="rows",
)

counts = edited.groupby("Sheet").size().to_dict()
ui.sheet_chips({spec.sheet: int(counts.get(spec.sheet, 0)) for spec in specs})


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #

ui.section(3, "Build", "totals, margins and the Summary are wired on the way out")

_cta, _ = st.columns([1, 2.1], gap="medium")
if _cta.button("Build workbook", type="primary", width="stretch"):
    keep = dict(zip(edited["Item Code"], edited["Sheet"]))
    selected = [item for item in items if item.code in keep]

    # Honour any sheet reassignment made in the table by rewriting the prefix
    # rules into explicit, full-length code rules for the moved items.
    moved = {code: sheet for code, sheet in keep.items()
             if sheet != assignment.get(code)}
    effective = [
        SheetSpec(
            spec.sheet, spec.layout, spec.source, spec.currency,
            spec.prefixes + [code for code, sheet in moved.items()
                             if sheet == spec.sheet],
        )
        for spec in specs
    ]

    stream, report = build_workbook(
        selected, effective,
        quantities=dict(zip(edited["Item Code"], edited["Qty"])),
        options=BuildOptions(freight_factor=freight_factor,
                             eur_factor=eur_factor,
                             vat_rate=vat_rate,
                             discounts=discounts),
    )

    st.success("Built {} rows across six offer sheets, from a {} table of {} "
               "items.".format(sum(report.rows_per_sheet.values()),
                               " / ".join(LOOKUP_TABLES),
                               sum(report.catalog_rows.values())))
    if report.unmatched:
        st.warning(
            "{} code(s) matched no prefix rule and went to the fallback sheet: "
            "{}".format(len(report.unmatched), ", ".join(report.unmatched[:20])))

    _dl, _ = st.columns([1, 2.1], gap="medium")
    _dl.download_button(
        "Download pricing workbook",
        data=stream.getvalue(),
        file_name="pricing tool {}.xlsx".format(date.today().isoformat()),
        mime=MIME_XLSX, type="primary", width="stretch",
    )
    st.caption(
        "Open in Excel and let it calculate — every derived cell is a live "
        "formula, so totals, margins and the Summary refresh on the spot."
    )
