import json
import re
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from typing import Optional, Tuple
from db.connection import fetch_companies, fetch_projects_by_company, fetch_type_last_edit_dates


def _normalize_last_edit(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.isna(parsed):
            return text
        return parsed.strftime("%d-%m")
    except Exception:
        return text


def _normalize_last_edit_full(value):
    """Same as _normalize_last_edit but keeps the full year (DD-MM-YYYY),
    used for the 'اخر تعديل' caption shown under a dropdown after a
    selection is made."""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.isna(parsed):
            return text
        return parsed.strftime("%d-%m-%Y")
    except Exception:
        return text


def _inject_dropdown_styles(data_map: dict):
    """
    data_map: {option_name: 'DD-MM' or None}

    Instead of encoding the date into the option label ("name\\ndate")
    and trying to re-split that text in JS, we keep labels clean
    (just the name) and pass a name -> date lookup map to the browser.
    The JS then APPENDS a small badge next to each row/closed-input,
    rather than parsing/rewriting the row's own text. This works for:
      - <li>/<div role="option"> rows, even when BaseWeb wraps part of
        the text in extra spans for search-match highlighting
      - the closed select's <input>, which can't contain child nodes,
        so its badge is positioned on the wrapper instead
    """
    clean_map = {k: v for k, v in data_map.items() if v}

    components.html(
        f"""
        <script>
        (function() {{
            const doc = window.parent.document;
            window.__ddMap = Object.assign(window.__ddMap || {{}}, {json.dumps(clean_map)});

            function ensureStyle() {{
                if (doc.getElementById('dd-style')) return;
                const style = doc.createElement('style');
                style.id = 'dd-style';
                style.textContent = `
                    ul[role="listbox"], div[data-baseweb="popover"] ul {{
                        padding: 6px !important;
                    }}
                    li[role="option"], div[role="option"] {{
                        position: relative !important;
                        overflow: visible !important;
                        text-overflow: clip !important;
                        white-space: normal !important;
                        text-align: right !important;
                        direction: rtl !important;
                        font-weight: 600 !important;
                        font-size: 0.98rem !important;
                        letter-spacing: 0.1px !important;
                        color: #eef0f4 !important;
                        min-height: 44px !important;
                        height: auto !important;
                        padding: 10px 14px 14px 14px !important;
                        margin: 2px 0 !important;
                        border-radius: 10px !important;
                        border: 1px solid transparent !important;
                        transition: background-color 160ms ease, border-color 160ms ease, transform 120ms ease !important;
                    }}
                    li[role="option"]:not(:last-child), div[role="option"]:not(:last-child) {{
                        border-bottom: 1px solid rgba(255,255,255,0.05) !important;
                    }}
                    li[role="option"]:hover, div[role="option"]:hover {{
                        background-color: rgba(255,255,255,0.07) !important;
                        border-color: rgba(255,255,255,0.08) !important;
                        transform: translateX(-1px) !important;
                    }}
                    li[role="option"][aria-selected="true"], div[role="option"][aria-selected="true"] {{
                        background-color: rgba(255, 180, 84, 0.10) !important;
                        border-color: rgba(255, 180, 84, 0.25) !important;
                    }}
                    [data-baseweb="select"] {{
                        position: relative !important;
                    }}
                    [data-baseweb="select"] > div {{
                        transition: box-shadow 160ms ease, border-color 160ms ease !important;
                    }}
                    .dd-opt-date {{
                        position: absolute;
                        left: 8px;
                        bottom: 6px;
                        display: inline-flex;
                        align-items: center;
                        gap: 2px;
                        font-weight: 800;
                        font-size: 8px;
                        letter-spacing: 0.2px;
                        color: #ffcf8a;
                        background: linear-gradient(135deg, rgba(255,180,84,0.22), rgba(255,150,60,0.10));
                        border: 1px solid rgba(255, 180, 84, 0.30);
                        padding: 1.5px 5px;
                        border-radius: 999px;
                        box-shadow: 0 1px 2px rgba(0,0,0,0.22), inset 0 1px 0 rgba(255,255,255,0.05);
                        direction: ltr;
                        unicode-bidi: isolate;
                        z-index: 5;
                        pointer-events: none;
                        transition: transform 160ms ease, box-shadow 160ms ease;
                    }}
                    .dd-opt-date::before {{
                        content: '';
                        width: 3px;
                        height: 3px;
                        border-radius: 50%;
                        background: #ffb454;
                        box-shadow: 0 0 3px rgba(255,180,84,0.8);
                        flex-shrink: 0;
                    }}
                    li[role="option"]:hover .dd-opt-date, div[role="option"]:hover .dd-opt-date {{
                        transform: scale(1.04);
                    }}
                    [data-baseweb="select"] > div > .dd-opt-date {{
                        bottom: 6px;
                        left: 8px;
                    }}
                `;
                doc.head.appendChild(style);
            }}

            function norm(t) {{ return (t || '').trim(); }}

            function addBadge(container, date) {{
                if (!container || container.querySelector(':scope > .dd-opt-date')) return;
                const badge = doc.createElement('span');
                badge.className = 'dd-opt-date';
                badge.setAttribute('dir', 'ltr');
                badge.textContent = date;
                container.appendChild(badge);
            }}

            function styleOptionRow(el) {{
                const date = window.__ddMap[norm(el.textContent)];
                if (date) addBadge(el, date);
            }}

            function styleClosedInput(inputEl) {{
                const date = window.__ddMap[norm(inputEl.value)];
                const wrapper = inputEl.closest('[data-baseweb="select"]');
                if (date && wrapper) addBadge(wrapper, date);
            }}

            function scan() {{
                ensureStyle();
                doc.querySelectorAll('li[role="option"], div[role="option"]').forEach(styleOptionRow);
                doc.querySelectorAll('[data-baseweb="select"] input').forEach(styleClosedInput);
            }}

            if (!window.__ddObserverInit) {{
                window.__ddObserverInit = true;
                const observer = new MutationObserver(scan);
                observer.observe(doc.body, {{ childList: true, subtree: true, characterData: true }});
            }}
            scan();
        }})();
        </script>
        """,
        height=0,
    )


def _inject_global_dropdown_polish():
    """One-time page-level CSS (not the popover styling above): makes the
    'اختر الشركة' / 'اختر المشروع' widget labels small+bold, and defines
    the '.dd-lastedit-*' classes used by _render_last_edit_caption."""
    if st.session_state.get("_dd_global_polish_injected"):
        return
    st.session_state["_dd_global_polish_injected"] = True

    st.markdown(
        """
        <style>
        div[data-testid="stWidgetLabel"] p {
            font-size: 12.5px !important;
            font-weight: 700 !important;
            color: #c7cbd4 !important;
            letter-spacing: 0.15px !important;
        }
        div[data-baseweb="select"] > div {
            border-radius: 10px !important;
            transition: box-shadow 160ms ease, border-color 160ms ease !important;
        }
        div[data-baseweb="select"]:focus-within > div {
            box-shadow: 0 0 0 3px rgba(255, 180, 84, 0.15) !important;
            border-color: rgba(255, 180, 84, 0.45) !important;
        }
        ul[role="listbox"],
        div[role="listbox"],
        div[data-baseweb="popover"] ul,
        div[data-baseweb="popover"] div[role="listbox"] {
            max-height: 280px !important;
            overflow-y: auto !important;
            overflow-x: hidden !important;
            scrollbar-width: thin !important;
            scrollbar-color: rgba(255, 207, 138, 0.45) transparent !important;
        }
        ul[role="listbox"]::-webkit-scrollbar,
        div[role="listbox"]::-webkit-scrollbar,
        div[data-baseweb="popover"] ul::-webkit-scrollbar,
        div[data-baseweb="popover"] div[role="listbox"]::-webkit-scrollbar {
            width: 7px !important;
        }
        ul[role="listbox"]::-webkit-scrollbar-thumb,
        div[role="listbox"]::-webkit-scrollbar-thumb,
        div[data-baseweb="popover"] ul::-webkit-scrollbar-thumb,
        div[data-baseweb="popover"] div[role="listbox"]::-webkit-scrollbar-thumb {
            background: rgba(255, 207, 138, 0.45) !important;
            border-radius: 999px !important;
        }
        .dd-lastedit-wrap {
            display: flex;
            justify-content: flex-end;
            align-items: center;
            gap: 6px;
            margin-top: 4px;
            margin-bottom: 2px;
            direction: rtl;
            animation: dd-fade-in 180ms ease;
        }
        .dd-lastedit-label {
            font-size: 11px;
            font-weight: 700;
            color: #9aa0ab;
        }
        .dd-lastedit-date {
            font-size: 11px;
            font-weight: 800;
            color: #ffcf8a;
            background: linear-gradient(135deg, rgba(255,180,84,0.20), rgba(255,150,60,0.08));
            border: 1px solid rgba(255, 180, 84, 0.28);
            padding: 2px 9px;
            border-radius: 999px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.20);
            direction: ltr;
            unicode-bidi: isolate;
        }
        @keyframes dd-fade-in {
            from { opacity: 0; transform: translateY(-2px); }
            to { opacity: 1; transform: translateY(0); }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_last_edit_caption(date_full: Optional[str]):
    """Small bold 'اخر تعديل: DD-MM-YYYY' line placed right under a dropdown."""
    if not date_full:
        return
    st.markdown(
        f"""
        <div class="dd-lastedit-wrap">
            <span class="dd-lastedit-label">اخر تعديل:</span>
            <span class="dd-lastedit-date">{date_full}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def create_factory_dropdown() -> Optional[str]:
    display_to_factory = {
        "الكل": None,
        "التجمع": "التجمع",
        "بدر": "بدر",
    }
    display_choice = st.selectbox(
        "اسم المصنع",
        options=list(display_to_factory.keys()),
        index=0,
        key="factory_select",
        help="اختر المصنع أو اترك الكل لعرض جميع الشركات",
    )
    return display_to_factory.get(display_choice)


def create_company_dropdown(conn, factory_name: Optional[str] = None):
    _inject_global_dropdown_polish()

    companies_df = fetch_companies(conn, factory_name=factory_name)
    if companies_df.empty or "اسم الشركة" not in companies_df.columns:
        return None

    companies_df = companies_df.copy()
    companies_df["آخر تعديل"] = companies_df.get("آخر تعديل", None)
    companies_df["آخر تعديل"] = companies_df["آخر تعديل"].replace({pd.NaT: None}).astype(object)

    rows = companies_df.loc[companies_df["اسم الشركة"].notna(), ["اسم الشركة", "آخر تعديل"]]
    rows = rows.drop_duplicates(subset=["اسم الشركة"], keep="last")
    rows = rows.sort_values("اسم الشركة", key=lambda s: s.str.lower())
    options = [(row["اسم الشركة"], row["آخر تعديل"]) for _, row in rows.iterrows()]

    if not options:
        return None

    query = st.text_input(
        "🔍 اكتب بداية اسم الشركة",
        value="",
        placeholder="اكتب بداية اسم الشركة ...",
        key="company_search",
    )
    if query:
        q = str(query).strip().lower()
        filtered = [opt for opt in options if opt[0].lower().startswith(q)]
    else:
        filtered = options

    if not filtered:
        return None

    data_map = {name: _normalize_last_edit(date) for name, date in filtered}
    full_date_map = {name: _normalize_last_edit_full(date) for name, date in filtered}
    _inject_dropdown_styles(data_map)

    selected = st.selectbox(
        "اختر الشركة",
        options=filtered,
        index=0 if filtered else None,
        format_func=lambda x: x[0],  # clean name only - date is injected separately
        key="company_select",
    )
    _render_last_edit_caption(full_date_map.get(selected[0]))
    return selected[0]


def create_project_dropdown(conn, company_name: str):
    _inject_global_dropdown_polish()

    if not company_name:
        return None
    projects_df = fetch_projects_by_company(conn, company_name)
    projects_df = projects_df.copy()
    projects_df["آخر تعديل"] = projects_df.get("آخر تعديل", None)
    projects_df["آخر تعديل"] = projects_df["آخر تعديل"].replace({pd.NaT: None}).astype(object)

    rows = projects_df.loc[projects_df["اسم المشروع"].notna(), ["اسم المشروع", "آخر تعديل"]]
    rows = rows.drop_duplicates(subset=["اسم المشروع"], keep="first")
    rows = rows.sort_values("اسم المشروع", key=lambda s: s.str.lower())
    options = [(row["اسم المشروع"], row["آخر تعديل"]) for _, row in rows.iterrows()]

    if not options:
        return None

    data_map = {name: _normalize_last_edit(date) for name, date in options}
    full_date_map = {name: _normalize_last_edit_full(date) for name, date in options}
    _inject_dropdown_styles(data_map)

    selected = st.selectbox(
        "اختر المشروع",
        options=options,
        index=0,
        format_func=lambda x: x[0],  # clean name only - date is injected separately
        placeholder="— اختر —",
        key="project_select",
    )
    _render_last_edit_caption(full_date_map.get(selected[0]))
    return selected[0]


def create_type_dropdown(conn) -> Tuple[Optional[str], Optional[str]]:
    _inject_global_dropdown_polish()

    display_to_key = {
        "تقرير مالي": "financial_report",
        "العقد": "contract",
        "خطابات و شيكات ضمان": "guarantee",
        "المستخلصات": "invoice",
        "الشيكات و التحويلات": "checks",
        "مواد اوليه و مقاولين باطن": "supplier_costs",
        "شهادات تامينات": "social_insurance_certificate",
    }
    options = [
        (display_name, key)
        for display_name, key in display_to_key.items()
    ]
    selected = st.selectbox(
        "اختر نوع البيانات",
        options=options,
        index=0 if options else None,
        format_func=lambda x: x[0],
        key="type_select",
    )
    return selected[0], selected[1]


def create_column_search(df: pd.DataFrame):
    if df.empty:
        return None, None
    col = st.selectbox("اختَر عمودًا للبحث", options=df.columns.tolist(), index=0)
    term = st.text_input("كلمة/عبارة للبحث")
    return col, term


def create_date_range():
    c1, c2 = st.columns(2)
    with c1:
        d_from = st.date_input("من تاريخ", value=None, format="YYYY-MM-DD")
    with c2:
        d_to = st.date_input("إلى تاريخ", value=None, format="YYYY-MM-DD")
    d_from = pd.to_datetime(d_from).date() if d_from else None
    d_to = pd.to_datetime(d_to).date() if d_to else None
    return d_from, d_to