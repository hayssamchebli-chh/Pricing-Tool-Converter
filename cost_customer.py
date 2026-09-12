"""Streamlit front end: customer sales history -> the Calcul offer sheet.

Upload a sales report and a catalogue extract, plus either a Calcul workbook to
fill or a BOQ to raise one from, and download the sheet with the three
``Last U.P. / Cur. / Date`` blocks, the descriptions and the Business Central
reference columns in place.

One of the pages wired up in ``home.py``; run that rather than this file.
"""

from __future__ import annotations

import io
from datetime import date

import streamlit as st

import ui
from catalog import best_quantity_sheet, read_catalog, read_quantities
from pricing import (
    DEFAULT_PRICE_FIELD, MAX_HISTORY, PRICE_FIELDS, BuildError, build_calcul,
    build_histories, catalog_index, fill_calcul, histories_to_frame,
    read_calcul_lines, read_sales,
)

st.set_page_config(page_title="Cost + Customer Pricing Tool", page_icon="🧮",
                   layout="wide")
ui.inject_theme()

MIME_XLSX = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")


# --------------------------------------------------------------------------- #
# cached parsing
# --------------------------------------------------------------------------- #

@st.cache_data(show_spinner=False)
def _load_sales(payload: bytes, filename: str, price_field: str):
    buffer = io.BytesIO(payload)
    buffer.name = filename          # the extension tells a CSV from a workbook
    return read_sales(buffer, price_field=price_field)


@st.cache_data(show_spinner=False)
def _load_catalog(payload: bytes):
    return read_catalog(io.BytesIO(payload))


@st.cache_data(show_spinner=False)
def _load_quantities(payload: bytes, known_codes: tuple = ()):
    return read_quantities(io.BytesIO(payload), known_codes)


# --------------------------------------------------------------------------- #
# sidebar
# --------------------------------------------------------------------------- #

st.sidebar.header("Bill of quantities")
boq_file = st.sidebar.file_uploader(
    "BOQ — optional", type=["xlsx", "xlsm"],
    help="Any sheet pairing a column of item codes with a quantity column.")
st.sidebar.caption(
    "Upload one and the offer is raised from its codes and quantities, so no "
    "Calcul file is needed. Leave it empty to fill a Calcul workbook you "
    "already have."
)

st.sidebar.divider()
st.sidebar.header("Sales history")
price_field = st.sidebar.selectbox(
    "Price basis", list(PRICE_FIELDS),
    index=list(PRICE_FIELDS).index(DEFAULT_PRICE_FIELD),
    format_func=lambda key: PRICE_FIELDS[key],
    help="What Last U.P. should mean. A report carrying a net price is taken "
         "at its word; one with only a unit price and a line discount has the "
         "discount applied.")
collapse = st.sidebar.checkbox(
    "Collapse repeated lines", value=True,
    help="Counts lines sharing a date, price and currency as one sale, so the "
         "three slots hold three distinct sales.")
include_returns = st.sidebar.checkbox(
    "Include credit memos", value=False,
    help="Off by default: a negative-quantity line reverses an invoice rather "
         "than recording a price the customer paid.")


# --------------------------------------------------------------------------- #
# uploads
# --------------------------------------------------------------------------- #

st.title("Cost + Customer Pricing Tool")

ui.section(1, "Source files", "sales report and catalogue are required")

left, middle, right = st.columns(3, gap="medium")
with left:
    with st.container(border=True):
        sales_file = st.file_uploader(
            "Customer sales report", type=["xlsx", "xlsm", "xls", "csv"],
            help="Posting Date · No. · Quantity · OC Net Price · Currency")
        st.caption("Fills Last U.P. / Cur. / Date, most recent sale first.")
with middle:
    with st.container(border=True):
        catalog_file = st.file_uploader(
            "Special Inquiry Worksheet", type=["xlsx", "xlsm"],
            help="Item No.1 · Description · Stock Available Quantity · Landed USD")
        st.caption("Fills Description, Stock AV (BC) and Landed usd (BC).")
with right:
    with st.container(border=True):
        if boq_file is None:
            calcul_file = st.file_uploader(
                "Calcul workbook", type=["xlsx", "xlsm"],
                help="The offer sheet to fill, matched on Item Code.")
            st.caption("Your codes, quantities and ex-works prices are kept; "
                       "the costing formulas are rewritten.")
        else:
            calcul_file = None
            st.markdown("**Calcul workbook**")
            st.caption("Not needed — the offer is being raised from the BOQ in "
                       "the sidebar, on the app's own template.")

if sales_file is None or catalog_file is None:
    st.info("Upload a sales report and a Special Inquiry Worksheet to start.")
    st.stop()
if boq_file is None and calcul_file is None:
    st.info("Upload a Calcul workbook to fill, or a BOQ in the sidebar to "
            "raise one from scratch.")
    st.stop()

try:
    sales = _load_sales(sales_file.getvalue(), sales_file.name, price_field)
    items, skipped = _load_catalog(catalog_file.getvalue())
except BuildError as problem:
    st.error(str(problem))
    st.stop()
except Exception as problem:                      # unreadable file, bad sheet
    st.error("Could not read that file: {}".format(problem))
    st.stop()

if not items:
    st.error(
        "No catalogue rows found. Expected a sheet whose header row carries an "
        "item code, a description and a unit price."
    )
    st.stop()

histories = build_histories(sales, collapse_duplicates=collapse,
                            include_returns=include_returns)
catalog = catalog_index(items)

# --------------------------------------------------------------------------- #
# the offer lines: from the BOQ, or from the Calcul workbook as it stands
# --------------------------------------------------------------------------- #

if boq_file is not None:
    per_sheet = _load_quantities(boq_file.getvalue(), tuple(sorted(catalog)))
    if not per_sheet:
        st.error(
            "No quantity column found in that BOQ. The app looks for a column "
            "headed *Qty* or *Quantity* beside a column of item codes."
        )
        st.stop()
    names = list(per_sheet)
    chosen = st.selectbox(
        "Quantity sheet", names, index=names.index(best_quantity_sheet(per_sheet)),
        help="The raw BOQ and its pivot usually both appear here and carry the "
             "same figures — pick either.") if len(names) > 1 else names[0]
    lines = [(code, qty, "") for code, qty in sorted(per_sheet[chosen].items())]
else:
    try:
        lines = read_calcul_lines(io.BytesIO(calcul_file.getvalue()))
    except BuildError as problem:
        st.error(str(problem))
        st.stop()
    if not lines:
        st.error("That Calcul workbook lists no item codes.")
        st.stop()

frame = histories_to_frame(lines, histories, catalog)


# --------------------------------------------------------------------------- #
# review
# --------------------------------------------------------------------------- #

ui.section(2, "Review", "three most recent sales per item, newest first")

with_history = int(frame["Last U.P.(1)"].notna().sum())
with_catalog = int(frame["Landed usd"].notna().sum())
returns = int((sales["quantity"].fillna(1) <= 0).sum())

metrics = st.columns(4)
metrics[0].metric("Offer lines", len(frame))
metrics[1].metric("With sales history", with_history)
metrics[2].metric("Found in catalogue", with_catalog)
metrics[3].metric("Sales lines read", len(sales))

st.caption("Last U.P. read as the **{}**.".format(sales.attrs.get(
    "price_used", "price as reported")))
if skipped:
    st.caption("Ignored, no catalogue header: " + ", ".join(skipped))
if returns and not include_returns:
    st.caption("{} credit-memo line(s) skipped. Turn on *Include credit memos* "
               "in the sidebar to count them.".format(returns))

missing_history = frame.loc[frame["Last U.P.(1)"].isna(), "Item Code"].tolist()
missing_catalog = frame.loc[frame["Landed usd"].isna(), "Item Code"].tolist()

if with_history == 0:
    st.warning(
        "Not one offer line was found in the sales report, so every "
        "Last U.P. column would come out empty. Check the report covers these "
        "items before building."
    )
if missing_history and with_history:
    with st.expander("{} line(s) with no sales history".format(len(missing_history))):
        st.write(", ".join(missing_history))
if missing_catalog:
    with st.expander("{} line(s) not in the catalogue".format(len(missing_catalog))):
        st.write(", ".join(missing_catalog))

st.dataframe(frame, width="stretch", hide_index=True, height=380)
st.caption(
    "Items sold fewer than {} times leave the unused slots empty. "
    "*U.P. Ex.* stays yours to key — every costed column follows from it."
    .format(MAX_HISTORY))


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #

ui.section(3, "Build", "U. Landed keeps the stock-versus-quantity condition")

_cta, _ = st.columns([1, 2.1], gap="medium")
if _cta.button("Build workbook", type="primary", width="stretch"):
    try:
        if boq_file is not None:
            payload = build_calcul(lines, histories, catalog)
            st.success(
                "Raised a {}-line offer from the BOQ; {} carry sales history."
                .format(len(lines), with_history))
        else:
            payload, matched, no_history, no_catalog = fill_calcul(
                io.BytesIO(calcul_file.getvalue()), histories, catalog)
            st.success(
                "Filled {} row(s): {} with sales history, {} found in the "
                "catalogue.".format(len(matched), len(matched) - len(no_history),
                                    len(matched) - len(no_catalog)))
    except BuildError as problem:
        st.error(str(problem))
        st.stop()

    _dl, _ = st.columns([1, 2.1], gap="medium")
    _dl.download_button(
        "Download Calcul workbook",
        data=payload,
        file_name="calcul {}.xlsx".format(date.today().isoformat()),
        mime=MIME_XLSX, type="primary", width="stretch",
    )
    st.caption(
        "Open in Excel and let it calculate. *U. Landed* takes the catalogue's "
        "landed cost while stock covers the quantity, and prices off *D.U.P. "
        "Ex.* grossed up by the sheet's own factors once it does not."
    )
