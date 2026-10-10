import sys

# onnxruntime (OCR backend for the PDF to Excel tool) must be imported
# before pandas on Windows: importing pandas first leaves onnxruntime's
# native DLL unable to initialize ("DLL load failed ... initialization
# routine failed"), verified on onnxruntime 1.22 and 1.27. Windows only:
# on Linux (where the app is hosted) there's no such issue, and importing it
# eagerly would load the whole ONNX runtime into memory for every visitor,
# even though only OCR runs ever use it.
if sys.platform == "win32":
    try:
        import onnxruntime  # noqa: F401
    except ImportError:
        pass

import streamlit as st

from core.pages import EXTERNAL_TOOL_KEYS, TOOL_PAGES
from core.styles import CSS
from pages_registry import get_pages

st.set_page_config(page_title="DataForge", page_icon="assets/logo_icon.png", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)
st.logo("assets/wordmark.png", icon_image="assets/logo_icon.png")

nav = st.navigation(get_pages())

# Leaving a tool page discards its uploader widgets (Streamlit drops state
# for widgets that aren't rendered), so the results that tool kept in
# session state (built workbooks, zips, previews) can never be shown again
# but would stay in memory for the whole session. Drop them on navigation.
# Session keys are namespaced "<tool>__..." by core.session.ns_key, and the
# tool names are TOOL_PAGES' keys.
_page_tools = {entry["page"].url_path: tool for tool, entry in TOOL_PAGES.items() if "page" in entry}
_current_tool = _page_tools.get(nav.url_path)
if st.session_state.get("_active_tool", "") != _current_tool:
    for key in list(st.session_state.keys()):
        prefix = str(key).split("__", 1)[0]
        if prefix in _page_tools.values() and prefix != _current_tool:
            del st.session_state[key]
    st.session_state["_active_tool"] = _current_tool

# Rendered before nav.run(): tool pages call st.stop() until a file is
# uploaded, which ends the script right there, so anything placed after
# nav.run() never appeared on those pages.
with st.sidebar:
    # Hosted as their own separate apps, not this app's st.Page objects -
    # rendered as plain external links below the real nav, not mixed into it.
    for key in EXTERNAL_TOOL_KEYS:
        entry = TOOL_PAGES[key]
        st.page_link(entry["url"], label=entry["title"], icon=entry["icon"])
    if st.button("Reset everything", help="Clear all uploads, settings and results", width="stretch"):
        # Session state only. Parsed uploads live in a session-scoped cache
        # that is dropped with the session; clearing st.cache_data here used
        # to wipe the shared cache for every other visitor too.
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    st.caption("v1.0")

nav.run()
