"""Presentation layer: theme CSS and the small chrome pieces built on top of it.

Streamlit's own theme (see ``.streamlit/config.toml``) carries the palette,
type and radii.  What lives here is the handful of things the theme cannot
express: the masthead, numbered section headers, metric cards and the row of
sheet chips.

Streamlit exposes no CSS custom properties for its active theme and stamps no
marker on the DOM, so nothing here may hardcode a text or surface colour - a
light-mode literal turns invisible the moment the viewer is in dark mode.
Instead the chrome derives its tones from ``currentColor``, which Streamlit
does set correctly either way: mixing the inherited text colour into
transparency yields muted text on both grounds, and a few percent of it yields
a surface that darkens on white and lifts on near-black.  Only the brand accent
is a fixed hue, and it is pinned per render to the theme Streamlit reports.
"""

from __future__ import annotations

import html

import streamlit as st

FONT_IMPORT = (
    "@import url('https://fonts.googleapis.com/css2"
    "?family=Inter:wght@400;500;600;700"
    "&family=JetBrains+Mono:wght@400;500&display=swap');"
)

CSS = FONT_IMPORT + """
:root {
  --ptc-brand: #1E3A5F;
  --ptc-on-brand: #FFFFFF;
  --ptc-step-gap: 1.9rem;
}
/* First-paint default. inject_theme() appends an override keyed off the theme
   Streamlit actually resolved, which wins over this. */
@media (prefers-color-scheme: dark) {
  :root { --ptc-brand: #7CA9F0; --ptc-on-brand: #0B1220; }
}

/* Trim Streamlit's stock top padding so the masthead sits high. */
[data-testid="stMainBlockContainer"] { padding-top: 2.6rem; max-width: 1180px; }

/* Numerals line up in columns - item codes, prices, counts. */
[data-testid="stDataFrame"], [data-testid="stMetricValue"], .ptc-chip-count {
  font-variant-numeric: tabular-nums;
  font-feature-settings: "tnum" 1;
}

/* ---- masthead ---------------------------------------------------------- */
.ptc-masthead {
  display: flex; align-items: center; gap: .95rem;
  padding-bottom: 1.1rem; margin-bottom: .4rem;
  border-bottom: 1px solid rgba(128, 138, 157, .32);
  border-bottom-color: color-mix(in srgb, currentColor 16%, transparent);
}
.ptc-mark {
  flex: 0 0 auto; width: 40px; height: 40px; border-radius: 9px;
  background: var(--ptc-brand); color: var(--ptc-on-brand);
  display: flex; align-items: center; justify-content: center;
}
.ptc-mark svg { width: 21px; height: 21px; }
.ptc-title { font-size: 1.32rem; font-weight: 650; letter-spacing: -.021em;
             line-height: 1.2; color: inherit; margin: 0; }
.ptc-sub {
  font-size: .845rem; margin: .2rem 0 0; line-height: 1.45;
  color: rgba(128, 138, 157, .95);
  color: color-mix(in srgb, currentColor 66%, transparent);
}

/* ---- section headers --------------------------------------------------- */
.ptc-section { display: flex; align-items: center; gap: .6rem;
               margin: var(--ptc-step-gap) 0 .85rem; }
.ptc-step {
  flex: 0 0 auto; width: 21px; height: 21px; border-radius: 50%;
  background: color-mix(in srgb, var(--ptc-brand) 13%, transparent);
  border: 1px solid color-mix(in srgb, var(--ptc-brand) 32%, transparent);
  color: var(--ptc-brand);
  font-size: .68rem; font-weight: 650; line-height: 19px; text-align: center;
}
.ptc-section h2 { font-size: 1.0rem; font-weight: 620; letter-spacing: -.012em;
                  margin: 0; color: inherit; }
.ptc-section p {
  /* 64% measures 5.28:1 on light and 6.95:1 on dark. Lighter mixes fail small
     text on the light ground: 58% lands at 4.33, under the 4.5 floor. */
  font-size: .82rem; margin: 0;
  color: rgba(128, 138, 157, .9);
  color: color-mix(in srgb, currentColor 64%, transparent);
}
.ptc-rule {
  height: 1px; flex: 1 1 auto; margin-left: .35rem;
  background: rgba(128, 138, 157, .28);
  background: color-mix(in srgb, currentColor 14%, transparent);
}

/* ---- metric cards ------------------------------------------------------ */
[data-testid="stMetric"] {
  border-radius: 10px; padding: .85rem 1rem .95rem;
  background: color-mix(in srgb, currentColor 3.5%, transparent);
  border: 1px solid rgba(128, 138, 157, .28);
  border-color: color-mix(in srgb, currentColor 14%, transparent);
}
[data-testid="stMetricLabel"] p {
  font-size: .715rem !important; font-weight: 550; letter-spacing: .055em;
  text-transform: uppercase;
  color: color-mix(in srgb, currentColor 72%, transparent) !important;
}
[data-testid="stMetricValue"] { font-size: 1.5rem !important; font-weight: 600;
                                letter-spacing: -.02em; }

/* ---- sheet chips ------------------------------------------------------- */
.ptc-chips { display: flex; flex-wrap: wrap; gap: .45rem; margin: .2rem 0 .3rem; }
.ptc-chip {
  display: inline-flex; align-items: center; gap: .5rem;
  padding: .32rem .62rem .32rem .7rem; border-radius: 7px;
  font-size: .785rem; color: inherit; white-space: nowrap;
  background: color-mix(in srgb, currentColor 3.5%, transparent);
  border: 1px solid rgba(128, 138, 157, .28);
  border-color: color-mix(in srgb, currentColor 14%, transparent);
}
.ptc-chip-count {
  font-family: 'JetBrains Mono', ui-monospace, monospace;
  font-size: .74rem; font-weight: 500; padding: .05rem .4rem; border-radius: 5px;
  background: color-mix(in srgb, var(--ptc-brand) 14%, transparent);
  color: var(--ptc-brand);
}
.ptc-chip.is-empty { opacity: .48; }
.ptc-chip.is-empty .ptc-chip-count {
  background: color-mix(in srgb, currentColor 10%, transparent);
  color: inherit;
}

/* ---- controls ---------------------------------------------------------- */
.stButton button, [data-testid="stDownloadButton"] button {
  font-weight: 550; letter-spacing: .003em;
  transition: transform .12s ease-out, filter .12s ease-out;
}
.stButton button:hover, [data-testid="stDownloadButton"] button:hover {
  filter: brightness(1.08);
}
.stButton button:active, [data-testid="stDownloadButton"] button:active {
  transform: translateY(1px);
}
[data-testid="stFileUploaderDropzone"] { border-radius: 9px; }

[data-testid="stSidebar"] h2 {
  font-size: .74rem !important; font-weight: 600; letter-spacing: .07em;
  text-transform: uppercase;
  color: color-mix(in srgb, currentColor 62%, transparent) !important;
}

/* Someone reading this with reduced motion asked for less of it. */
@media (prefers-reduced-motion: reduce) {
  .stButton button, [data-testid="stDownloadButton"] button { transition: none; }
}

@media (max-width: 640px) {
  .ptc-masthead { gap: .7rem; }
  .ptc-title { font-size: 1.14rem; }
  .ptc-section p { display: none; }
}
"""

_MARK = (
    '<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" '
    'aria-hidden="true">'
    '<path d="M4 4.5h9.5L20 11v8.5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1v-14a1 1 0 0 1 1-1z" '
    'stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"/>'
    '<path d="M13 4.6V11h6.2" stroke="currentColor" stroke-width="1.6" '
    'stroke-linejoin="round"/>'
    '<path d="M7.6 16.4h8M7.6 13.2h4.2" stroke="currentColor" stroke-width="1.6" '
    'stroke-linecap="round"/></svg>'
)


BRAND = {"light": ("#1E3A5F", "#FFFFFF"), "dark": ("#7CA9F0", "#0B1220")}


def _resolved_accent() -> str:
    """Pin the accent to the theme Streamlit actually resolved.

    ``prefers-color-scheme`` alone is not enough: someone who picks a theme in
    Streamlit's settings menu against their system preference would otherwise
    get the navy accent on a near-black ground.  ``st.context.theme`` reports
    what Streamlit settled on, and can lag by one rerun on first load - which
    the media-query default above covers.
    """
    try:
        pair = BRAND.get(st.context.theme.type)
    except Exception:
        pair = None
    if not pair:
        return ""
    return ":root{{--ptc-brand:{};--ptc-on-brand:{};}}".format(*pair)


def inject_theme() -> None:
    """Load the fonts and custom chrome. Call once, right after page config."""
    st.markdown("<style>{}{}</style>".format(CSS, _resolved_accent()),
                unsafe_allow_html=True)


def masthead(title: str, subtitle: str) -> None:
    st.markdown(
        '<div class="ptc-masthead"><div class="ptc-mark">{mark}</div>'
        '<div><p class="ptc-title">{title}</p>'
        '<p class="ptc-sub">{sub}</p></div></div>'.format(
            mark=_MARK, title=html.escape(title), sub=html.escape(subtitle)),
        unsafe_allow_html=True,
    )


def section(step: int, title: str, note: str = "") -> None:
    st.markdown(
        '<div class="ptc-section"><span class="ptc-step">{step}</span>'
        '<h2>{title}</h2>{note}<span class="ptc-rule"></span></div>'.format(
            step=step, title=html.escape(title),
            note="<p>{}</p>".format(html.escape(note)) if note else ""),
        unsafe_allow_html=True,
    )


def sheet_chips(counts: dict) -> None:
    """One chip per offer sheet, dimmed where nothing was routed to it."""
    chips = "".join(
        '<span class="ptc-chip{empty}">{name}'
        '<span class="ptc-chip-count">{n}</span></span>'.format(
            empty="" if n else " is-empty", name=html.escape(str(name)), n=n)
        for name, n in counts.items()
    )
    st.markdown('<div class="ptc-chips">{}</div>'.format(chips),
                unsafe_allow_html=True)
