"""Shared UI theme for SPC Data Visualization.

Single source of truth for all CSS styling. Both app.py and Quick Test
import inject_theme() to get consistent styling without copy-paste drift.

Design: Claude Code aesthetic — monospace data, dense layout, sharp corners,
borders over shadows, teal accent on pure white.
"""

import streamlit as st

# -- Design tokens ----------------------------------------------------------
FONT_BODY = "'IBM Plex Sans', 'Barlow', system-ui, sans-serif"
FONT_MONO = "'JetBrains Mono', 'IBM Plex Mono', 'SF Mono', monospace"
FONT_HEADING = "'Barlow Condensed', 'Archivo Narrow', system-ui, sans-serif"

WHITE = "#FFFFFF"
BG_SUBTLE = "#FAFAFA"       # barely-there gray for sidebar / cards
BORDER = "#D4D4D4"          # neutral-400
BORDER_LIGHT = "#E5E5E5"    # neutral-300
TEXT_PRIMARY = "#111111"     # near-black
TEXT_SECONDARY = "#525252"   # neutral-600
TEXT_MUTED = "#737373"       # neutral-500
ACCENT = "#0D9488"           # teal-600 — distinctive, not default blue
ACCENT_HOVER = "#0F766E"    # teal-700
DANGER = "#DC2626"           # red-600
SUCCESS = "#16A34A"          # green-600
WARNING = "#D97706"          # amber-600

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

/* ================================================================
   RESET — pure white everywhere, kill Streamlit's default theming
   ================================================================ */
html, body, .stApp, .main, .main > div, .block-container,
[data-testid="stAppViewContainer"],
[data-testid="stAppViewBlockContainer"],
[data-testid="stVerticalBlock"],
[data-testid="stHorizontalBlock"],
[data-testid="column"],
[data-testid="stMainBlockContainer"],
.element-container, .stMarkdown,
div[data-testid] {{
    background-color: {WHITE} !important;
    color: {TEXT_PRIMARY} !important;
}}
.stApp div:not([data-testid="stPlotlyChart"] *) {{
    background-color: transparent;
}}
.stApp > div, .stApp > div > div,
[data-testid="stVerticalBlockBorderWrapper"],
[data-testid="stVerticalBlockBorderWrapper"] > div {{
    background-color: {WHITE} !important;
}}

/* ================================================================
   TYPOGRAPHY — body IBM Plex Sans, headings Barlow Condensed,
   data/metrics JetBrains Mono
   ================================================================ */
html, body, div, section, main, aside, header, footer, nav,
p, a, label, li, td, th,
input, textarea, select, button, small, strong, em,
[data-testid="stMarkdownContainer"],
[data-testid="stWidgetLabel"],
[data-baseweb="select"],
[data-baseweb="input"] {{
    font-family: {FONT_BODY} !important;
}}
h1, h2, h3, h4, h5, h6 {{
    font-family: {FONT_HEADING} !important;
    letter-spacing: -0.01em !important;
}}
h1 {{
    font-size: 1.5rem !important;
    font-weight: 700 !important;
    color: {TEXT_PRIMARY} !important;
    margin-bottom: 0.25rem !important;
}}
h2, h3 {{
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.05em !important;
    color: {TEXT_SECONDARY} !important;
}}
h4 {{
    font-size: 0.8rem !important;
    font-weight: 600 !important;
    color: {TEXT_SECONDARY} !important;
}}
p, span, label, li, td, th, div, small, strong, a {{
    color: {TEXT_PRIMARY} !important;
}}
p span, label span, h1 span, h2 span, h3 span, h4 span,
[data-testid="stMarkdownContainer"] span,
[data-testid="stWidgetLabel"] span {{
    font-family: {FONT_BODY} !important;
}}
code, pre, [data-testid="stMetricValue"],
[data-testid="stMetricValue"] * {{
    font-family: {FONT_MONO} !important;
    color: {TEXT_PRIMARY} !important;
}}
/* Preserve Material Symbols icon font */
span[style*="Material Symbols"] {{
    font-family: "Material Symbols Rounded" !important;
    color: {TEXT_MUTED} !important;
}}

/* ================================================================
   HEADER BAR — minimal top line
   ================================================================ */
header, header[data-testid="stHeader"], .stAppHeader,
[data-testid="stHeader"], [data-testid="stToolbar"] {{
    background-color: {WHITE} !important;
    border-bottom: 1px solid {BORDER_LIGHT} !important;
}}
[data-testid="stDecoration"] {{ display: none !important; }}
header button, [data-testid="stToolbar"] button {{
    color: {TEXT_MUTED} !important;
}}

/* ================================================================
   SIDEBAR — barely-there tint, dense labels
   ================================================================ */
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div,
section[data-testid="stSidebar"] > div > div,
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"],
section[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"],
section[data-testid="stSidebar"] .element-container {{
    background-color: {BG_SUBTLE} !important;
}}
section[data-testid="stSidebar"] h1 {{
    font-size: 1.1rem !important;
    font-weight: 700 !important;
    font-family: {FONT_HEADING} !important;
}}
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
    font-size: 0.7rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.08em !important;
    color: {TEXT_MUTED} !important;
    margin-top: 0.5rem !important;
    margin-bottom: 0.25rem !important;
}}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"],
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {{
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    color: {TEXT_SECONDARY} !important;
    font-family: {FONT_BODY} !important;
}}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] small,
section[data-testid="stSidebar"] li {{
    color: {TEXT_SECONDARY} !important;
}}
section[data-testid="stSidebar"] hr {{
    border-color: {BORDER_LIGHT} !important;
    margin: 0.5rem 0 !important;
}}
section[data-testid="stSidebar"] .stCaption,
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{
    font-size: 0.7rem !important;
    color: {TEXT_MUTED} !important;
}}

/* ================================================================
   INPUTS — sharp corners, thin borders, compact
   ================================================================ */
[data-baseweb="select"],
[data-baseweb="select"] > div,
[data-baseweb="select"] ul,
[data-baseweb="input"],
[data-baseweb="input"] > div {{
    background: {WHITE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
    color: {TEXT_PRIMARY} !important;
    font-size: 0.82rem !important;
}}
[data-baseweb="select"] *, [data-baseweb="input"] * {{
    color: {TEXT_PRIMARY} !important;
}}
[data-baseweb="tag"] {{
    background: {BORDER_LIGHT} !important;
    color: {TEXT_PRIMARY} !important;
    border-radius: 1px !important;
    font-size: 0.75rem !important;
    font-family: {FONT_MONO} !important;
}}
[data-testid="stNumberInput"] input {{
    background: {WHITE} !important;
    color: {TEXT_PRIMARY} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
    font-family: {FONT_MONO} !important;
    font-size: 0.82rem !important;
}}
[data-testid="stNumberInput"] button {{
    background: {BG_SUBTLE} !important;
    background-color: {BG_SUBTLE} !important;
    border: 1px solid {BORDER} !important;
    color: {TEXT_SECONDARY} !important;
    border-radius: 2px !important;
}}
[data-testid="stNumberInput"] button:hover {{
    background: {BORDER_LIGHT} !important;
    background-color: {BORDER_LIGHT} !important;
}}
[data-testid="stNumberInput"] button svg {{
    fill: {TEXT_SECONDARY} !important;
    color: {TEXT_SECONDARY} !important;
}}

/* ================================================================
   FILE UPLOADER
   ================================================================ */
[data-testid="stFileUploader"],
[data-testid="stFileUploader"] > div,
[data-testid="stFileUploader"] section,
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploaderDropzone"] > div {{
    background: {BG_SUBTLE} !important;
    border-color: {BORDER} !important;
    color: {TEXT_SECONDARY} !important;
}}
[data-testid="stFileUploaderDropzone"] {{
    border: 1px dashed {BORDER} !important;
    border-radius: 2px !important;
}}
[data-testid="stFileUploaderDropzone"] * {{
    color: {TEXT_MUTED} !important;
}}
[data-testid="stFileUploaderDropzone"] button,
[data-testid="baseButton-secondary"] {{
    background: {WHITE} !important;
    color: {ACCENT} !important;
    border: 1px solid {ACCENT} !important;
    border-radius: 2px !important;
}}
[data-testid="stFileUploaderFile"],
[data-testid="stFileUploaderFile"] > div,
[data-testid="stUploadedFile"],
[data-testid="stUploadedFile"] > div,
li[class*="uploadedFile"], li[class*="UploadedFile"],
div[class*="uploadedFile"], div[class*="UploadedFile"],
div[class*="UploadedFileInfo"] {{
    background: {BG_SUBTLE} !important;
    color: {TEXT_SECONDARY} !important;
    border: 1px solid {BORDER_LIGHT} !important;
    border-radius: 2px !important;
}}

/* ================================================================
   DROPDOWNS / POPOVER / LISTBOX
   ================================================================ */
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
[data-baseweb="popover"] > div > div,
[data-baseweb="popover"] > div > div > div,
[data-baseweb="popover"] ul, [data-baseweb="popover"] li,
[data-baseweb="menu"], [data-baseweb="menu"] > div,
[data-baseweb="listbox"], [data-baseweb="listbox"] > div,
div[role="listbox"], div[role="listbox"] > div,
div[role="listbox"] ul, div[role="listbox"] li {{
    background: {WHITE} !important;
    background-color: {WHITE} !important;
    border-color: {BORDER_LIGHT} !important;
}}
[data-baseweb="popover"] {{
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important;
}}
[data-baseweb="menu"] *, [data-baseweb="listbox"] *,
[data-baseweb="popover"] * {{
    color: {TEXT_PRIMARY} !important;
    font-size: 0.82rem !important;
}}
[data-baseweb="menu"] li:hover, [data-baseweb="listbox"] li:hover,
div[role="listbox"] li:hover,
[data-baseweb="menu"] li[aria-selected="true"],
[data-baseweb="listbox"] li[aria-selected="true"] {{
    background: {BG_SUBTLE} !important;
    background-color: {BG_SUBTLE} !important;
}}

/* ================================================================
   CHECKBOX — clean square, teal checked state
   ================================================================ */
[data-testid="stCheckbox"] {{
    background-color: transparent !important;
}}
[data-testid="stCheckbox"] label,
[data-testid="stCheckbox"] label > span,
[data-testid="stCheckbox"] label > div {{
    background-color: transparent !important;
}}
[data-testid="stCheckbox"] [role="checkbox"],
[data-testid="stCheckbox"] span[data-baseweb="checkbox"],
[data-testid="stCheckbox"] span[data-baseweb="checkbox"] > span,
[data-testid="stCheckbox"] span[data-baseweb="checkbox"] > div,
[data-testid="stCheckbox"] label > span:first-child,
[data-testid="stCheckbox"] label > span:first-child > span,
[data-testid="stCheckbox"] label > div:first-child,
[data-testid="stCheckbox"] label > div:first-child > div,
[data-testid="stCheckbox"] label > div:first-child > span {{
    background-color: {WHITE} !important;
    background: {WHITE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
}}
[data-testid="stCheckbox"] [role="checkbox"][aria-checked="true"],
[data-testid="stCheckbox"] span[data-baseweb="checkbox"][aria-checked="true"],
[data-testid="stCheckbox"] input:checked + span,
[data-testid="stCheckbox"] input:checked + div,
[data-testid="stCheckbox"] input:checked ~ span,
[data-testid="stCheckbox"] input:checked ~ div,
[data-testid="stCheckbox"] label:has(input:checked) > span:first-child,
[data-testid="stCheckbox"] label:has(input:checked) > span:first-child > span,
[data-testid="stCheckbox"] label:has(input:checked) > div:first-child,
[data-testid="stCheckbox"] label:has(input:checked) > div:first-child > div,
[data-testid="stCheckbox"] label:has(input:checked) > div:first-child > span {{
    background-color: {ACCENT} !important;
    background: {ACCENT} !important;
    border-color: {ACCENT} !important;
}}
[data-testid="stCheckbox"] [role="checkbox"]:hover,
[data-testid="stCheckbox"] span[data-baseweb="checkbox"]:hover {{
    border-color: {TEXT_MUTED} !important;
}}
[data-testid="stCheckbox"] label p,
[data-testid="stCheckbox"] label span,
[data-testid="stCheckbox"] [data-testid="stWidgetLabel"],
[data-testid="stCheckbox"] [data-testid="stWidgetLabel"] p,
[data-testid="stCheckbox"] [data-testid="stMarkdownContainer"],
[data-testid="stCheckbox"] [data-testid="stMarkdownContainer"] p {{
    color: {TEXT_SECONDARY} !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    font-family: {FONT_BODY} !important;
    line-height: 1.4 !important;
}}

/* ================================================================
   RADIO — horizontal tab bar, sharp, minimal
   ================================================================ */
[data-testid="stRadio"] > div {{
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    gap: 0 !important;
    background: transparent !important;
    border-bottom: 1px solid {BORDER} !important;
    border-radius: 0 !important;
    padding: 0 !important;
    display: flex !important;
    width: 100% !important;
    overflow-x: auto !important;
}}
[data-testid="stRadio"] > div > label {{
    background: transparent !important;
    border-radius: 0 !important;
    padding: 5px 6px !important;
    margin: 0 !important;
    cursor: pointer !important;
    font-family: {FONT_BODY} !important;
    font-size: 0.7rem !important;
    font-weight: 500 !important;
    color: {TEXT_MUTED} !important;
    border-bottom: 2px solid transparent !important;
    transition: color 0.1s, border-color 0.1s !important;
    white-space: nowrap !important;
    text-align: center !important;
}}
[data-testid="stRadio"] > div > label:has(input:checked) {{
    color: {ACCENT} !important;
    font-weight: 600 !important;
    border-bottom: 2px solid {ACCENT} !important;
    background: transparent !important;
}}
[data-testid="stRadio"] > div > label:hover:not(:has(input:checked)) {{
    color: {TEXT_SECONDARY} !important;
}}
[data-testid="stRadio"] > div > label > div:first-child {{
    display: none !important;
}}
/* Re-assert after catch-all */
.stApp [data-testid="stRadio"] > div {{
    background: transparent !important;
    background-color: transparent !important;
}}
.stApp [data-testid="stRadio"] > div > label {{
    background: transparent !important;
    background-color: transparent !important;
}}
.stApp [data-testid="stRadio"] > div > label:has(input:checked) {{
    background: transparent !important;
    background-color: transparent !important;
}}

/* ================================================================
   SLIDER
   ================================================================ */
[data-testid="stSlider"] * {{
    color: {TEXT_SECONDARY} !important;
}}
[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {{
    background: {ACCENT} !important;
    border-color: {ACCENT} !important;
}}

/* ================================================================
   METRIC CARDS — dense, bordered, monospace values
   ================================================================ */
[data-testid="stMetric"] {{
    background: {WHITE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
    padding: 8px 10px !important;
}}
[data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * {{
    font-family: {FONT_BODY} !important;
    font-size: 0.7rem !important;
    font-weight: 500 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.04em !important;
    color: {TEXT_MUTED} !important;
}}
[data-testid="stMetricValue"], [data-testid="stMetricValue"] * {{
    font-family: {FONT_MONO} !important;
    font-size: 1.1rem !important;
    font-weight: 500 !important;
    color: {TEXT_PRIMARY} !important;
}}

/* ================================================================
   ALERTS
   ================================================================ */
[data-testid="stAlert"], div[role="alert"],
[data-testid="stNotification"] {{
    background: {BG_SUBTLE} !important;
    border: 1px solid {BORDER_LIGHT} !important;
    border-radius: 2px !important;
    border-left: 3px solid {ACCENT} !important;
}}
[data-testid="stAlert"] *, div[role="alert"] * {{
    color: {TEXT_SECONDARY} !important;
    font-size: 0.82rem !important;
}}
[data-testid="stAlert"] svg {{ fill: {TEXT_MUTED} !important; }}

/* ================================================================
   EXPANDER — flat, bordered
   ================================================================ */
[data-testid="stExpander"],
[data-testid="stExpander"] > div,
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary {{
    background: {WHITE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
    color: {TEXT_PRIMARY} !important;
}}
[data-testid="stExpander"] summary {{
    font-family: {FONT_BODY} !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    padding: 8px 12px !important;
}}
[data-testid="stExpander"] * {{ color: {TEXT_PRIMARY} !important; }}

/* ================================================================
   TABS — underline style, condensed labels
   ================================================================ */
[data-testid="stTabs"] button[role="tab"] {{
    font-family: {FONT_BODY} !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    color: {TEXT_MUTED} !important;
    background: transparent !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    border-radius: 0 !important;
    padding: 6px 14px !important;
}}
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
    color: {ACCENT} !important;
    border-bottom: 2px solid {ACCENT} !important;
    font-weight: 600 !important;
    background: transparent !important;
}}
[data-testid="stTabs"] button[role="tab"]:hover {{
    color: {TEXT_SECONDARY} !important;
    background: transparent !important;
}}
[data-testid="stTabs"] [role="tablist"] {{
    border-bottom: 1px solid {BORDER_LIGHT} !important;
    background: transparent !important;
    gap: 0 !important;
}}

/* ================================================================
   PLOTLY CHART — white
   ================================================================ */
[data-testid="stPlotlyChart"],
[data-testid="stPlotlyChart"] > div,
[data-testid="stPlotlyChart"] iframe,
.stPlotlyChart, .stPlotlyChart > div,
iframe[title="streamlit_plotly_events"],
div[class*="plotly"], div[class*="chart"] {{
    background: {WHITE} !important;
}}

/* ================================================================
   IFRAMES / COMPONENTS
   ================================================================ */
iframe, [data-testid="stIFrame"],
[data-testid="stCustomComponentV1"],
[data-testid="stCustomComponentV1"] > div,
.stHtml, .stHtml > div {{
    background: {WHITE} !important;
    border: none !important;
}}

/* ================================================================
   TABLES / DATAFRAMES — dense, monospace data
   ================================================================ */
[data-testid="stDataFrame"], [data-testid="stTable"] {{
    background: {WHITE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
}}
[data-testid="stDataFrame"] *, [data-testid="stTable"] * {{
    background: {WHITE} !important;
    color: {TEXT_PRIMARY} !important;
    font-size: 0.78rem !important;
}}

/* ================================================================
   FOOTER
   ================================================================ */
footer, footer * {{ background: {WHITE} !important; color: {TEXT_MUTED} !important; }}

/* ================================================================
   SVG ICONS — muted, except Plotly and checkbox
   ================================================================ */
.stApp svg {{ fill: {TEXT_MUTED} !important; color: {TEXT_MUTED} !important; }}
[data-testid="stPlotlyChart"] svg,
.js-plotly-plot svg, .plot-container svg {{
    fill: unset !important; color: unset !important;
}}
.stApp [data-testid="stCheckbox"] svg {{
    fill: {WHITE} !important; color: {WHITE} !important; stroke: {WHITE} !important;
}}

/* ================================================================
   BUTTONS — flat, sharp, teal primary
   ================================================================ */
.stButton > button {{
    background: {WHITE} !important;
    color: {TEXT_SECONDARY} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 2px !important;
    font-family: {FONT_BODY} !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
}}
.stButton > button:hover {{
    background: {BG_SUBTLE} !important;
    border-color: {TEXT_MUTED} !important;
}}
[data-testid="baseButton-primary"],
.stButton > button[kind="primary"] {{
    background: {ACCENT} !important;
    color: {WHITE} !important;
    border: 1px solid {ACCENT} !important;
}}
[data-testid="baseButton-primary"]:hover,
.stButton > button[kind="primary"]:hover {{
    background: {ACCENT_HOVER} !important;
    border-color: {ACCENT_HOVER} !important;
}}
.stApp button svg {{ fill: {TEXT_MUTED} !important; }}

/* ================================================================
   COLOR PICKER — compact
   ================================================================ */
[data-testid="stColorPicker"] label {{
    font-size: 0.78rem !important;
    color: {TEXT_SECONDARY} !important;
}}
</style>
"""


def inject_theme():
    """Inject the shared CSS theme into the current Streamlit page."""
    st.markdown(CSS, unsafe_allow_html=True)
