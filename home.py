"""Entry point: the two tools, side by side as tabs.

Each tab is its own script and runs on its own — ``st.stop()`` in one leaves
the other untouched, which is what lets ``app.py`` stay exactly as it was.
``streamlit run app.py`` still works too, and gives the Pricing Tool Converter
alone with no tab bar.

Run with:  streamlit run home.py
"""

from __future__ import annotations

import streamlit as st

# The tab labels ship at .875rem, small enough beside a 2.4rem page title that
# the tabs read as toolbar furniture rather than the app's top-level switch.
# The only markdown Streamlit puts in the header is those labels, so scoping to
# stHeader leaves every other paragraph in the app alone.
TAB_CSS = """
<style>
[data-testid="stHeader"] [data-testid="stMarkdownContainer"] p {
  font-size: 1.05rem;
  font-weight: 560;
  letter-spacing: -.005em;
}
</style>
"""

PAGES = [
    st.Page("app.py", title="Pricing Tool Converter", icon="📊",
            url_path="pricing-tool", default=True),
    st.Page("cost_customer.py", title="Cost + Customer Pricing Tool", icon="🧮",
            url_path="cost-customer-pricing"),
    st.Page("last_purchase.py", title="Last Purchase Price", icon="🧾",
            url_path="last-purchase-price"),
]

try:
    navigation = st.navigation(PAGES, position="top")
except Exception:
    # Streamlit predating top navigation: the sidebar list is accepted by
    # every version that has st.navigation at all, so the app still runs.
    navigation = st.navigation(PAGES)

st.markdown(TAB_CSS, unsafe_allow_html=True)
navigation.run()
