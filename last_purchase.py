"""Streamlit front end: purchase history -> the Order workbook's price columns.

Upload a purchase-history export (one row per invoice line) and download the
Order workbook with the three most recent unit prices per item written into
``Last U.P.(1..3) (EUR)`` and ``Date(1..3)``, most recent first.

One of the pages wired up in ``home.py``; run that rather than this file.
"""

from __future__ import annotations

import io
from datetime import date

import streamlit as st

import ui
from purchases import (
    MAX_HISTORY, TEMPLATE_PATH, BuildError, build_histories, build_order,
    fill_existing_order, histories_to_frame, read_purchases,
)

st.set_page_config(page_title="Last Purchase Price", page_icon="🧾",
                   layout="wide")
ui.inject_theme()

MIME_XLSX = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")

NEW_ORDER = "New Order from the template"
FILL_ORDER = "Fill an Order I already have"


# --------------------------------------------------------------------------- #
# cached parsing
# --------------------------------------------------------------------------- #

@st.cache_data(show_spinner=False)
def _load_purchases(payload: bytes, filename: str):
    # The name carries the extension, which is how a CSV is told from a
    # workbook; BytesIO alone would lose it.
    buffer = io.BytesIO(payload)
    buffer.name = filename
    return read_purchases(buffer)


# --------------------------------------------------------------------------- #
# sidebar
# --------------------------------------------------------------------------- #

st.sidebar.header("Order")
mode = st.sidebar.radio(
    "What to produce", [NEW_ORDER, FILL_ORDER],
    help="A new Order lists every item the history mentions, one row each. "
         "Filling one you already have leaves its rows alone and writes only "
         "the six price and date columns.")

template_file = order_file = None
if mode == NEW_ORDER:
    template_file = st.sidebar.file_uploader(
        "Order template", type=["xlsx", "xlsm"],
        help="Leave empty to use the template shipped with the app, "
             "{}.".format(TEMPLATE_PATH.name))
else:
    order_file = st.sidebar.file_uploader(
        "Order to fill", type=["xlsx", "xlsm"],
        help="Matched on Item Code. Rows whose code the history does not "
             "mention are left untouched and listed back to you.")

st.sidebar.divider()
st.sidebar.header("History")
collapse = st.sidebar.checkbox(
    "Collapse repeated lines", value=True,
    help="Counts rows sharing a date *and* a price as one purchase, so the "
         "three slots hold three distinct purchases. Off takes the last three "
         "rows as they come.")
st.sidebar.caption(
    "Where two purchases share a date at different prices, the one further "
    "down the file is taken as the more recent."
)


# --------------------------------------------------------------------------- #
# uploads
# --------------------------------------------------------------------------- #

st.title("Last Purchase Price")

ui.section(1, "Purchase history", "one row per invoice line")

with st.container(border=True):
    history_file = st.file_uploader(
        "Last purchases export", type=["xlsx", "xlsm", "xls", "csv"],
        help="Item code · description · posting date · unit cost")
    st.caption(
        "Headers are matched loosely — `No_` or `Item Code`, "
        "`InvLine_Posting_Date` or `Date`, `InvLine_Unit_Cost` or `Unit Price` "
        "all read correctly."
    )

if history_file is None:
    st.info("Upload a purchase history to start.")
    st.stop()

try:
    purchases = _load_purchases(history_file.getvalue(), history_file.name)
except BuildError as problem:
    st.error(str(problem))
    st.stop()
except Exception as problem:                      # unreadable file, bad sheet
    st.error("Could not read that file: {}".format(problem))
    st.stop()

if purchases.empty:
    st.error("No purchase lines found in that file — every row was missing an "
             "item code.")
    st.stop()

histories = build_histories(purchases, collapse_duplicates=collapse)
frame = histories_to_frame(histories)


# --------------------------------------------------------------------------- #
# review
# --------------------------------------------------------------------------- #

ui.section(2, "Review", "the three most recent prices per item, newest first")

undated = int(purchases["date"].isna().sum())
unpriced = int(purchases["price"].isna().sum())
complete = int(frame["Last U.P.({}) (EUR)".format(MAX_HISTORY)].notna().sum())

metrics = st.columns(4)
metrics[0].metric("Purchase lines", len(purchases))
metrics[1].metric("Items", len(histories))
metrics[2].metric("With {} prices".format(MAX_HISTORY), complete)
metrics[3].metric("Single purchase",
                  int((frame["Last U.P.(2) (EUR)"].isna()).sum()))

if undated or unpriced:
    st.warning(
        "Skipped {} line(s) with no usable date and {} with no usable price. "
        "Everything else was read.".format(undated, unpriced))

st.dataframe(frame, width="stretch", hide_index=True, height=380)
st.caption(
    "Items bought fewer than {} times leave the unused slots empty.".format(
        MAX_HISTORY))


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #

ui.section(3, "Build", "prices and dates land in columns L to Q")

if mode == FILL_ORDER and order_file is None:
    st.info("Upload the Order you want filled — the uploader is in the "
            "sidebar.")
    st.stop()

_cta, _ = st.columns([1, 2.1], gap="medium")
if _cta.button("Build Order", type="primary", width="stretch"):
    try:
        if mode == NEW_ORDER:
            payload = build_order(histories, template=template_file)
            matched, unmatched = list(histories), []
        else:
            payload, matched, unmatched = fill_existing_order(
                order_file, histories)
    except BuildError as problem:
        st.error(str(problem))
        st.stop()

    if mode == NEW_ORDER:
        st.success(
            "Built an Order of {} item(s); {} carry a full {} prices.".format(
                len(histories), complete, MAX_HISTORY))
    else:
        st.success("Filled {} row(s) from the purchase history.".format(
            len(matched)))
        if unmatched:
            with st.expander(
                    "{} row(s) the history does not cover".format(len(unmatched))):
                st.write(", ".join(unmatched))

    _dl, _ = st.columns([1, 2.1], gap="medium")
    _dl.download_button(
        "Download Order workbook",
        data=payload,
        file_name="order {}.xlsx".format(date.today().isoformat()),
        mime=MIME_XLSX, type="primary", width="stretch",
    )
    st.caption(
        "Qty, U.P., Agr.U.P. and the FOB columns are left for you; the "
        "template's own formulas and the Total row come down with the rows."
    )
