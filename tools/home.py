import streamlit as st

from core.pages import TOOL_PAGES
from core.session import release_unused_uploads
from ui.shared import render_hero

TOOL_ORDER = ["merge", "split", "concat", "diff", "trim", "reduce", "pdf", "lcc", "cibil"]


def _render_tool_card(entry):
    # st.container(border=True) gives real DOM nesting (unlike opening/closing
    # a <div> across separate st.markdown calls, which the browser auto-closes
    # per call -- that produced an empty phantom box). core/styles.py picks
    # these cards out by the .tool-card-icon they contain (:has()), so other
    # st.container() calls elsewhere in the app aren't styled as cards.
    with st.container(border=True):
        st.markdown(
            f'<div class="tool-card-icon" style="background:{entry["color"]}26;color:{entry["color"]};">'
            f'{entry["icon"]}</div>',
            unsafe_allow_html=True,
        )
        # "url" entries are external tools hosted as separate apps (no
        # st.Page of their own); st.page_link renders a raw https:// string
        # as an external link automatically.
        if "url" in entry:
            st.page_link(entry["url"], label=entry["title"], width="stretch")
        else:
            st.page_link(entry["page"], label=entry["page"].title, width="stretch")
        st.markdown(f'<div class="tool-card-desc">{entry["description"]}</div>', unsafe_allow_html=True)


def render_home_page():
    # No uploads live on the home page: free anything parsed on a tool page.
    release_unused_uploads()

    # Side gutter columns keep the hero + card grid from stretching edge to
    # edge of the (deliberately wide) block-container -- real Streamlit
    # columns, not a CSS max-width hack, so there's no risk of the container
    # ambiguity that caused the earlier "whole page in one box" bug.
    _, main_col, _ = st.columns([1, 16, 1])
    with main_col:
        render_hero(
            "Schema-aware, asks when unsure",
            ["Your Data, Forged."],
            "Merge, split, diff, and clean Excel files - with a schema-aware matching engine "
            "that reconciles renamed and truncated headers instead of just concatenating blindly.",
        )

        # One row of 3 columns per 3 tools - zip against a fixed 6-column
        # list would silently drop every tool past the 6th.
        cols = []
        for _ in range(0, len(TOOL_ORDER), 3):
            cols.extend(st.columns(3, gap="medium"))
        for col, key in zip(cols, TOOL_ORDER):
            with col:
                _render_tool_card(TOOL_PAGES[key])
