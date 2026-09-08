"""Entry point: the two tools, side by side as tabs.

Each tab is its own script and runs on its own — ``st.stop()`` in one leaves
the other untouched, which is what lets ``app.py`` stay exactly as it was.
``streamlit run app.py`` still works too, and gives the Pricing Tool Converter
alone with no tab bar.

Run with:  streamlit run home.py
"""

from __future__ import annotations

import streamlit as st

PAGES = [
    st.Page("app.py", title="Pricing Tool Converter", icon="📊",
            url_path="pricing-tool", default=True),
    st.Page("last_purchase.py", title="Last Purchase Price", icon="🧾",
            url_path="last-purchase-price"),
]

try:
    navigation = st.navigation(PAGES, position="top")
except Exception:
    # Streamlit predating top navigation: the sidebar list is accepted by
    # every version that has st.navigation at all, so the app still runs.
    navigation = st.navigation(PAGES)

navigation.run()
