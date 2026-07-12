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
                    li[role="option"], div[role="option"] {{
                        position: relative !important;
                        overflow: visible !important;
                        text-overflow: clip !important;
                        white-space: normal !important;
                        text-align: right !important;
                        direction: rtl !important;
                        font-weight: 500 !important;
                        font-size: 1rem !important;
                        color: #f5f6f8 !important;
                        min-height: 52px !important;
                        height: auto !important;
                        padding-top: 10px !important;
                        padding-bottom: 18px !important;
                        border-radius: 8px !important;
                    }}
                    li[role="option"]:hover, div[role="option"]:hover {{
                        background-color: rgba(255,255,255,0.08) !important;
                    }}
                    [data-baseweb="select"] {{
                        position: relative !important;
                    }}
                    .dd-opt-date {{
                        position: absolute;
                        left: 10px;
                        bottom: 8px;
                        font-weight: 800;
                        font-size: 9px;
                        letter-spacing: 0.2px;
                        color: #ffb454;
                        background: rgba(255, 180, 84, 0.14);
                        padding: 2px 6px;
                        border-radius: 5px;
                        direction: ltr;
                        unicode-bidi: isolate;
                        z-index: 5;
                        pointer-events: none;
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
    companies_df = fetch_companies(conn, factory_name=factory_name)
    if companies_df.empty or "اسم الشركة" not in companies_df.columns:
        message = "لا توجد شركات مطابقة للمصنع المحدد." if factory_name else "لا توجد شركات."
        st.info(message)
        return None

    companies_df = companies_df.copy()
    companies_df["آخر تعديل"] = companies_df.get("آخر تعديل", None)
    companies_df["آخر تعديل"] = companies_df["آخر تعديل"].replace({pd.NaT: None}).astype(object)

    rows = companies_df.loc[companies_df["اسم الشركة"].notna(), ["اسم الشركة", "آخر تعديل"]]
    rows = rows.drop_duplicates(subset=["اسم الشركة"], keep="last")
    rows = rows.sort_values("اسم الشركة", key=lambda s: s.str.lower())
    options = [(row["اسم الشركة"], row["آخر تعديل"]) for _, row in rows.iterrows()]

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
        st.info(f"لا توجد شركات تبدأ بـ «{query}»") if query else st.info("لا توجد شركات.")
        return None

    data_map = {name: _normalize_last_edit(date) for name, date in filtered}
    _inject_dropdown_styles(data_map)

    selected = st.selectbox(
        "اختر الشركة",
        options=filtered,
        index=0 if filtered else None,
        format_func=lambda x: x[0],  # clean name only - date is injected separately
        key="company_select",
    )
    return selected[0]


def create_project_dropdown(conn, company_name: str):
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
    _inject_dropdown_styles(data_map)

    selected = st.selectbox(
        "اختر المشروع",
        options=options,
        index=0,
        format_func=lambda x: x[0],  # clean name only - date is injected separately
        placeholder="— اختر —",
        key="project_select",
    )
    return selected[0]


def create_type_dropdown(conn) -> Tuple[Optional[str], Optional[str]]:
    display_to_key = {
        "تقرير مالي": "financial_report",
        "العقود": "contract",
        "خطابات الضمان": "guarantee",
        "المستخلصات": "invoice",
        "الشيكات / التحويلات": "checks",
        "مواد اوليه و مقاولين باطن": "supplier_costs",
        "شهادة تامينات": "social_insurance_certificate",
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