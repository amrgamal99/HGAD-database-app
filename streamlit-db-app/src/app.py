# app.py
import os
import re
import base64
from io import BytesIO
from pathlib import Path
from typing import Optional, Tuple, Dict, List

import pandas as pd
import streamlit as st

# ReportLab (PDF) + Arabic
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image as RLImage,
    PageBreak,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display

try:
    from PIL import Image as PILImage
except Exception:
    PILImage = None

from db.connection import (
    get_db_connection,
    fetch_data,
    fetch_financial_flow_view,
    fetch_contract_summary_view,
)
from components.filters import (
    create_company_dropdown,
    create_project_dropdown,
    create_type_dropdown,
    create_column_search,
)

# =========================================================
# Paths / Assets
# =========================================================
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

LOGO_CANDIDATES = [ASSETS_DIR / "logo.png"]
WIDE_LOGO_CANDIDATES = [
    ASSETS_DIR / "logo_wide.png",
    ASSETS_DIR / "logo_wide.jpg",
    ASSETS_DIR / "logo_wide.jpeg",
    ASSETS_DIR / "logo_wide.webp",
]

AR_FONT_CANDIDATES = [
    ASSETS_DIR / "Cairo-Regular.ttf",
    ASSETS_DIR / "Amiri-Regular.ttf",
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
]

_AR_RE = re.compile(r"[\u0600-\u06FF]")  # Arabic block


def _first_existing(paths) -> Optional[Path]:
    for p in paths:
        pth = Path(p)
        if pth.exists() and pth.is_file() and (pth.stat().st_size > 0):
            return pth
    return None


def _image_size(path: Path) -> Tuple[int, int]:
    if PILImage:
        try:
            with PILImage.open(path) as im:
                return im.size  # (w,h) px
        except Exception:
            pass
    return (600, 120)


def _img_to_data_uri(path: Path) -> Optional[str]:
    try:
        ext = path.suffix.lower().lstrip(".") or "png"
        mime = f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
        b64 = base64.b64encode(path.read_bytes()).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return None


def _site_logo_path() -> Optional[Path]:
    return _first_existing(LOGO_CANDIDATES)


def _wide_logo_path() -> Optional[Path]:
    return _first_existing(WIDE_LOGO_CANDIDATES)


def _first_existing_font_path() -> Optional[str]:
    p = _first_existing(AR_FONT_CANDIDATES)
    return str(p) if p else None


def register_arabic_font() -> Tuple[str, bool]:
    p = _first_existing_font_path()
    if p:
        name = os.path.splitext(os.path.basename(p))[0]
        try:
            pdfmetrics.registerFont(TTFont(name, p))
            return name, True
        except Exception:
            pass
    return "Helvetica", False


def looks_arabic(s: str) -> bool:
    return bool(_AR_RE.search(str(s or "")))


def shape_arabic(s: str) -> str:
    try:
        return get_display(arabic_reshaper.reshape(str(s)))
    except Exception:
        return str(s)


def _safe_filename(s: str) -> str:
    return (
        (s or "")
        .replace("/", "-").replace("\\", "-").replace(":", "-")
        .replace("*", "-").replace("?", "-").replace('"', "'")
        .replace("<", "(").replace(">", ")").replace("|", "-")
    )


# =========================================================
# Streamlit Page Config
# =========================================================
st.set_page_config(
    page_title="قاعدة البيانات والتقارير المالية | HGAD",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# Global Styles (CSS)
# =========================================================
st.markdown(
    """
<style>
/* Sidebar always open */
[data-testid="stSidebar"] { transform:none !important; visibility:visible !important; width:340px !important; min-width:340px !important; }
[data-testid="stSidebar"][aria-expanded="false"] { transform:none !important; visibility:visible !important; }
[data-testid="collapsedControl"], button[kind="header"],
button[title="Expand sidebar"], button[title="Collapse sidebar"],
[data-testid="stSidebarCollapseButton"] { display:none !important; }

/* RTL root */
html, body {
    direction: rtl !important;
    text-align: right !important;
    font-family: "Cairo","Noto Kufi Arabic","Segoe UI",Tahoma,sans-serif !important;
    white-space: normal !important;
    word-wrap: break-word !important;
    overflow-x: hidden !important;
}

/* DataFrame readability */
[data-testid="stDataFrame"] thead tr th {
    position: sticky; top: 0; background: #1f2937; color: #f9fafb; z-index: 2;
    font-weight: 700; font-size: 16px;
}
[data-testid="stDataFrame"] div[role="row"] { font-size: 15px; }
[data-testid="stDataFrame"] div[role="row"]:nth-child(even) { background-color: rgba(255,255,255,0.04); }

/* === Date inputs in MAIN area === */
.date-box {
    border: 1px solid #334155; border-radius: 12px; padding: 12px; background:#0b1220; margin-bottom: 10px;
}
.date-row { display:flex; gap: 10px; flex-wrap: wrap; align-items: center; }
.date-row > div { min-width: 200px; }
[data-testid="stDateInput"] input {
    background:#0f172a !important; color:#e5e7eb !important;
    border:1px solid #334155 !important; border-radius:10px !important;
    text-align:center !important; height:44px !important; min-width: 180px !important;
}
[data-testid="stDateInput"] label { color:#cbd5e1 !important; font-weight:700; }
.stPopover, div[role="dialog"] { z-index: 99999 !important; }

/* Header */
.fin-head {
    display:flex; justify-content: space-between; align-items:center;
    border: 1px dashed #1e3a8a55; border-radius: 14px;
    padding: 14px 18px; margin: 6px 0 12px 0; background: #0b1220;
}
.fin-head .line { font-size: 24px; font-weight: 900; color: #e5e7eb; }
.badge { display:inline-block; background:#1e3a8a; color:white; font-weight:700; padding:6px 12px; border-radius:999px; }

/* Two-column financial panel (RTL) */
.fin-panel { display:grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 10px; }
.fin-table { width:100%; border-collapse: collapse; table-layout: fixed; }
.fin-table th, .fin-table td {
    border: 1px solid #334155; padding: 10px; font-size: 14px;
    white-space: normal; word-wrap: break-word;
}
/* VALUE narrower (left), LABEL wider (right) */
.fin-table td.value { background:#111827; color:#e5e7eb; font-weight:800; text-align:center; width: 32%; }
.fin-table td.label { background:#0b1220; color:#e5e7eb; font-weight:700; text-align:right; width: 68%; }

/* Section title */
.hsec { color:#1E3A8A; font-weight:800; margin:0.2rem 0 0.6rem 0; font-size: 20px; }
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# Header (inline base64 logo)
# =========================================================
def _logo_html() -> str:
    p = _first_existing(LOGO_CANDIDATES)
    if not p:
        return ""
    ext = p.suffix.lower().lstrip(".") or "png"
    mime = f"image/{'jpeg' if ext in ('jpg','jpeg') else ext}"
    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    return f'<img src="data:{mime};base64,{b64}" width="64" />'

c_logo, c_title = st.columns([1, 6], gap="small")
with c_logo:
    st.markdown(_logo_html(), unsafe_allow_html=True)
with c_title:
    st.markdown(
        """
<h1 style="color:#1E3A8A; font-weight:800; margin:0;">
    قاعدة البيانات والتقارير المالية
    <span style="font-size:20px; color:#4b5563;">| HGAD Company</span>
</h1>
""",
        unsafe_allow_html=True,
    )
st.markdown('<hr style="border:0; height:2px; background:linear-gradient(to left, transparent, #1E3A8A, transparent);"/>', unsafe_allow_html=True)

# =========================================================
# Excel helpers (with WIDE LOGO)
# =========================================================
def _pick_excel_engine() -> Optional[str]:
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        pass
    try:
        import openpyxl  # noqa: F401
        return "openpyxl"
    except Exception:
        return None


def _insert_wide_logo(ws, workbook, start_row: int = 0, col: int = 0) -> int:
    """
    Insert the wide logo (if exists) into the worksheet using xlsxwriter.
    Returns the next empty row after the image (to start writing the table).
    """
    wlp = _wide_logo_path()
    if not wlp:
        return start_row
    try:
        # Reasonable scaling to keep it slim
        options = {"x_scale": 0.6, "y_scale": 0.6}
        ws.insert_image(start_row, col, str(wlp), options)
        return start_row + 6  # leave some space below the logo
    except Exception:
        return start_row


def _auto_excel_sheet(writer, df: pd.DataFrame, sheet_name: str, put_logo: bool = True):
    engine = writer.engine
    df_x = df.copy()
    safe_name = (sheet_name or "Sheet1")[:31]

    if engine == "xlsxwriter":
        wb = writer.book
        ws = wb.add_worksheet(safe_name)
        writer.sheets[safe_name] = ws

        # Optional logo
        row0 = 0
        if put_logo:
            row0 = _insert_wide_logo(ws, wb, start_row=0, col=0)

        hdr_fmt = wb.add_format({"align": "right", "bold": True})
        fmt_text = wb.add_format({"align": "right"})
        fmt_date = wb.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
        fmt_num  = wb.add_format({"align": "right", "num_format": "#,##0.00"})
        fmt_link = wb.add_format({"font_color": "blue", "underline": 1, "align": "right"})

        # Auto widths
        for idx, col in enumerate(df_x.columns):
            series = df_x[col]
            max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
            width_chars = min(max_len + 4, 60)
            if pd.api.types.is_datetime64_any_dtype(series):
                ws.set_column(idx, idx, max(14, width_chars), fmt_date)
            elif pd.api.types.is_numeric_dtype(series):
                ws.set_column(idx, idx, max(14, width_chars), fmt_num)
            elif "رابط" in str(col):
                ws.set_column(idx, idx, max(20, width_chars), fmt_link)
            else:
                ws.set_column(idx, idx, width_chars, fmt_text)

        # Header row
        for col_num, col_name in enumerate(df_x.columns):
            ws.write(row0, col_num, col_name, hdr_fmt)

        # Body
        for idx, col in enumerate(df_x.columns):
            series = df_x[col]
            if "رابط" in str(col):
                for r, val in enumerate(series, start=row0 + 1):
                    sval = "" if pd.isna(val) else str(val)
                    if sval.startswith(("http://", "https://")):
                        ws.write_url(r, idx, sval, fmt_link, string="فتح الرابط")
                    else:
                        ws.write(r, idx, sval, fmt_text)
            elif pd.api.types.is_datetime64_any_dtype(series):
                for r, val in enumerate(series, start=row0 + 1):
                    if pd.notna(val):
                        ws.write_datetime(r, idx, pd.to_datetime(val), fmt_date)
                    else:
                        ws.write_blank(r, idx, None, fmt_text)
            elif pd.api.types.is_numeric_dtype(series):
                for r, val in enumerate(series, start=row0 + 1):
                    if pd.notna(val):
                        ws.write_number(r, idx, float(val), fmt_num)
                    else:
                        ws.write_blank(r, idx, None, fmt_text)
            else:
                for r, val in enumerate(series, start=row0 + 1):
                    ws.write(r, idx, "" if pd.isna(val) else str(val), fmt_text)
    else:
        # openpyxl fallback (no image support here)
        df_x.to_excel(writer, index=False, sheet_name=safe_name)


def make_excel_bytes(df: pd.DataFrame, sheet_name: str = "البيانات", put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        _auto_excel_sheet(writer, df, sheet_name, put_logo=put_logo)
    buf.seek(0)
    return buf.getvalue()


def make_excel_combined_two_sheets(dfs: Dict[str, pd.DataFrame], put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        for sheet, df in dfs.items():
            _auto_excel_sheet(writer, df, sheet, put_logo=put_logo)
    buf.seek(0)
    return buf.getvalue()


def make_excel_single_sheet_stacked(dfs: Dict[str, pd.DataFrame], sheet_name="تقرير_موحد", put_logo: bool = True) -> Optional[bytes]:
    engine = _pick_excel_engine()
    if engine is None:
        return None

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=engine) as writer:
        if writer.engine == "xlsxwriter":
            wb = writer.book
            ws = wb.add_worksheet(sheet_name[:31])
            writer.sheets[sheet_name[:31]] = ws

            title_fmt = wb.add_format({"bold": True, "align": "right", "font_size": 12})
            hdr_fmt = wb.add_format({"align": "right", "bold": True})
            fmt_text = wb.add_format({"align": "right"})
            fmt_date = wb.add_format({"align": "right", "num_format": "yyyy-mm-dd"})
            fmt_num  = wb.add_format({"align": "right", "num_format": "#,##0.00"})
            fmt_link = wb.add_format({"font_color": "blue", "underline": 1, "align": "right"})

            row_offset = 0
            if put_logo:
                row_offset = _insert_wide_logo(ws, wb, start_row=row_offset, col=0)

            for title, df in dfs.items():
                # Title
                ws.write(row_offset, 0, title, title_fmt)
                row_offset += 1

                # Headers
                for c_idx, col in enumerate(df.columns):
                    ws.write(row_offset, c_idx, col, hdr_fmt)
                row_offset += 1

                # Body
                for r in range(len(df)):
                    for c_idx, col in enumerate(df.columns):
                        val = df.iloc[r, c_idx]
                        if "رابط" in str(col):
                            sval = "" if pd.isna(val) else str(val)
                            if sval.startswith(("http://", "https://")):
                                ws.write_url(row_offset, c_idx, sval, fmt_link, string="فتح الرابط")
                            else:
                                ws.write(row_offset, c_idx, sval, fmt_text)
                        elif pd.api.types.is_datetime64_any_dtype(df[col]):
                            if pd.notna(val):
                                ws.write_datetime(row_offset, c_idx, pd.to_datetime(val), fmt_date)
                            else:
                                ws.write_blank(row_offset, c_idx, None, fmt_text)
                        elif pd.api.types.is_numeric_dtype(df[col]):
                            if pd.notna(val):
                                ws.write_number(row_offset, c_idx, float(val), fmt_num)
                            else:
                                ws.write_blank(row_offset, c_idx, None, fmt_text)
                        else:
                            ws.write(row_offset, c_idx, "" if pd.isna(val) else str(val), fmt_text)
                    row_offset += 1

                # Auto width
                for c_idx, col in enumerate(df.columns):
                    series = df[col]
                    max_len = max([len(str(col))] + [len(str(v)) for v in series.values])
                    ws.set_column(c_idx, c_idx, min(max_len + 4, 60))
                row_offset += 2
        else:
            # Fallback: concat with separators
            out = []
            first = True
            for title, df in dfs.items():
                if first and put_logo and _wide_logo_path():
                    # openpyxl fallback can't embed image; just write a title marker
                    logo_note = pd.DataFrame([[f"[LOGO: {_wide_logo_path().name}]"] + [""] * (len(df.columns) - 1)], columns=df.columns)
                    out.append(logo_note)
                    first = False
                title_row = pd.DataFrame([[title] + [""] * (len(df.columns) - 1)], columns=df.columns)
                out.append(title_row)
                out.append(df)
                out.append(pd.DataFrame([[""] * len(df.columns)], columns=df.columns))
            big = pd.concat(out, ignore_index=True)
            big.to_excel(writer, index=False, sheet_name=sheet_name[:31])

    buf.seek(0)
    return buf.getvalue()


def make_csv_utf8(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================================================
# PDF helpers (auto-fit wide tables)
# =========================================================
def _format_numbers_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_numeric_dtype(out[c]):
            out[c] = out[c].map(lambda x: "" if pd.isna(x) else f"{float(x):,.2f}")
        else:
            def _fmt_cell(v):
                s = str(v)
                try:
                    if s.strip().endswith("%"):
                        return s
                    fv = float(s.replace(",", ""))
                    return f"{fv:,.2f}"
                except Exception:
                    return s
            out[c] = out[c].map(_fmt_cell)
    return out


def _pdf_table(
    df: pd.DataFrame,
    title: str = "",
    max_col_width: int = 120,
    font_size: float = 7.0,
    avail_width: Optional[float] = None
) -> list:
    font_name, _ = register_arabic_font()
    hdr_style = ParagraphStyle(
        name="Hdr", fontName=font_name, fontSize=font_size+0.6,
        textColor=colors.whitesmoke, alignment=1, leading=font_size+1.8
    )
    cell_rtl  = ParagraphStyle(
        name="CellR", fontName=font_name, fontSize=font_size,
        leading=font_size+1.5, alignment=2, wordWrap='CJK'
    )
    cell_ltr  = ParagraphStyle(
        name="CellL", fontName=font_name, fontSize=font_size,
        leading=font_size+1.5, alignment=0, wordWrap='CJK'
    )

    blocks = []
    if title:
        tstyle = ParagraphStyle(
            name="Sec", fontName=font_name, fontSize=font_size+2,
            alignment=2, textColor=colors.HexColor("#1E3A8A")
        )
        blocks += [Paragraph(shape_arabic(title), tstyle), Spacer(1, 4)]

    headers = [Paragraph(shape_arabic(c) if looks_arabic(c) else str(c), hdr_style) for c in df.columns]
    rows = [headers]
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            sval = "" if pd.isna(r[c]) else str(r[c])
            is_ar = looks_arabic(sval)
            cells.append(Paragraph(shape_arabic(sval) if is_ar else sval, cell_rtl if is_ar else cell_ltr))
        rows.append(cells)

    col_widths = []
    for c in df.columns:
        max_len = max(len(str(c)), df[c].astype(str).map(len).max())
        col_widths.append(min(max_len * 6.2, max_col_width))

    if avail_width:
        total = sum(col_widths)
        if total > avail_width:
            factor = avail_width / total
            col_widths = [w * factor for w in col_widths]

    table = Table(rows, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), font_name),
        ("FONTSIZE", (0,0), (-1,-1), font_size),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1E3A8A")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("BOTTOMPADDING", (0,0), (-1,0), 3),
        ("TOPPADDING", (0,1), (-1,-1), 2),
        ("BOTTOMPADDING", (0,1), (-1,-1), 2),
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("WORDWRAP", (0,0), (-1,-1), True),
    ]))
    blocks.append(table)
    return blocks


def _choose_pdf_font(df: pd.DataFrame) -> Tuple[int, float]:
    n = len(df.columns)
    if n >= 12:
        return 110, 6.6
    if n >= 9:
        return 120, 7.0
    return 140, 7.6


def make_pdf_bytes(df: pd.DataFrame, pdf_name: str = "") -> bytes:
    buf = BytesIO()
    font_name, arabic_ok = register_arabic_font()

    page = landscape(A4)
    left, right, top, bottom = 14, 14, 18, 14
    doc = SimpleDocTemplate(
        buf, pagesize=page, rightMargin=right, leftMargin=left,
        topMargin=top, bottomMargin=bottom
    )

    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=13.5, leading=15, alignment=1)
    base_title = "التقرير المالي"
    title_text = f"{base_title} ({pdf_name})" if pdf_name else base_title
    if arabic_ok:
        title_text = shape_arabic(title_text)

    elements = []
    wlp = _wide_logo_path()
    avail_w = page[0] - left - right
    if wlp and wlp.exists():
        try:
            if PILImage:
                w_px, h_px = _image_size(wlp)
                ratio = h_px / float(w_px) if w_px else 0.2
                img_h = max(22, avail_w * ratio * 0.55)
            else:
                img_h = 36
            logo_img = RLImage(str(wlp), hAlign="CENTER")
            logo_img.drawWidth = avail_w
            logo_img.drawHeight = img_h
            elements.append(logo_img)
            elements.append(Spacer(1, 4))
        except Exception:
            pass

    elements.append(Paragraph(title_text, title_style))
    elements.append(Spacer(1, 5))

    max_col_width, base_font = _choose_pdf_font(df)
    elements += _pdf_table(
        df,
        max_col_width=max_col_width,
        font_size=base_font,
        avail_width=avail_w
    )
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


def make_pdf_combined(summary_df: pd.DataFrame, flow_df: pd.DataFrame, header_text: str = "") -> bytes:
    buf = BytesIO()
    font_name, arabic_ok = register_arabic_font()

    page = landscape(A4)
    left, right, top, bottom = 14, 14, 18, 14
    doc = SimpleDocTemplate(
        buf, pagesize=page, rightMargin=right, leftMargin=left,
        topMargin=top, bottomMargin=bottom
    )

    title_style = ParagraphStyle(name="Title", fontName=font_name, fontSize=13.5, leading=15, alignment=1)
    head_style  = ParagraphStyle(name="Head",  fontName=font_name, fontSize=11, leading=14, alignment=2, textColor=colors.HexColor("#1E3A8A"))

    base_title = "التقرير المالي"
    if arabic_ok:
        base_title = shape_arabic(base_title)
        header_text = shape_arabic(header_text)

    elements = []
    wlp = _wide_logo_path()
    avail_w = page[0] - left - right
    if wlp and wlp.exists():
        try:
            if PILImage:
                w_px, h_px = _image_size(wlp)
                ratio = h_px / float(w_px) if w_px else 0.2
                img_h = max(22, avail_w * ratio * 0.55)
            else:
                img_h = 36
            logo_img = RLImage(str(wlp), hAlign="CENTER")
            logo_img.drawWidth = avail_w
            logo_img.drawHeight = img_h
            elements.append(logo_img)
            elements.append(Spacer(1, 4))
        except Exception:
            pass

    elements.append(Paragraph(base_title, title_style))
    if header_text:
        elements.append(Spacer(1, 4))
        elements.append(Paragraph(header_text, head_style))
    elements.append(Spacer(1, 6))

    max_w_s, f_s = _choose_pdf_font(summary_df)
    elements += _pdf_table(summary_df, title="ملخص المشروع", max_col_width=max_w_s, font_size=f_s, avail_width=avail_w)
    elements.append(PageBreak())
    max_w_f, f_f = _choose_pdf_font(flow_df)
    elements += _pdf_table(flow_df, title="دفتر التدفق", max_col_width=max_w_f, font_size=f_f, avail_width=avail_w)
    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# Helpers
# =========================================================
def fin_panel_two_tables(left_items: List[Tuple[str, str]], right_items: List[Tuple[str, str]]):
    """
    Render two side-by-side tables with reversed cells:
    [ value | label ]  -> label (Arabic) on the RIGHT, value on the LEFT.
    """
    def _table_html(items):
        rows = []
        for label, value in items:
            rows.append(f'<tr><td class="value">{value}</td><td class="label">{label}</td></tr>')
        return f'<table class="fin-table">{"".join(rows)}</table>'

    html = f'<div class="fin-panel"><div>{_table_html(right_items)}</div><div>{_table_html(left_items)}</div></div>'
    st.markdown(html, unsafe_allow_html=True)


def _apply_date_filter(df: pd.DataFrame, dfrom, dto) -> pd.DataFrame:
    if df is None or df.empty or (not dfrom and not dto):
        return df
    date_cols = [c for c in df.columns if any(k in str(c) for k in ["تاريخ", "إصدار", "date", "تعاقد"])]
    if not date_cols:
        return df
    out = df.copy()
    for col in date_cols:
        try:
            dseries = pd.to_datetime(out[col], errors="coerce").dt.date
            if dfrom:
                out = out[dseries >= dfrom]
            if dto:
                out = out[dseries <= dto]
        except Exception:
            pass
    return out


def _fmt_value(v) -> str:
    try:
        if isinstance(v, str) and v.strip().endswith("%"):
            return v
        f = float(str(v).replace(",", ""))
        return f"{f:,.2f}"
    except Exception:
        return "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v)


def _row_to_pairs_from_data(row: pd.Series) -> List[Tuple[str, str]]:
    """
    Build (label,value) pairs DIRECTLY from df_summary row.
    Removes technical columns and empty values.
    """
    ignore_substrings = {"id", "ID", "companyid", "contractid"}
    pairs = []
    for col, val in row.items():
        if any(k in str(col).lower() for k in ignore_substrings):
            continue
        sval = _fmt_value(val)
        if sval == "" or sval.lower() == "nan":
            continue
        pairs.append((str(col), sval))
    return pairs


def _split_pairs_two_columns(pairs: List[Tuple[str, str]]) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
    n = len(pairs)
    mid = (n + 1) // 2
    right = pairs[:mid]
    left = pairs[mid:]
    return left, right


# =========================================================
# Main App
# =========================================================
def main() -> None:
    conn = get_db_connection()
    if conn is None:
        st.error("فشل الاتصال بقاعدة البيانات. يرجى مراجعة بيانات الاتصال والتأكد من تشغيل الخادم.")
        return

    # Sidebar: ONLY selectors
    with st.sidebar:
        st.title("عوامل التصفية")
        company_name = create_company_dropdown(conn)
        project_name = create_project_dropdown(conn, company_name)
        type_label, type_key = create_type_dropdown()

    if not company_name or not project_name or not type_key:
        st.info("برجاء اختيار الشركة والمشروع ونوع البيانات من الشريط الجانبي لعرض النتائج.")
        return

    # === Global date filters in MAIN area (for ALL data types) ===
    g_date_from, g_date_to = None, None
    with st.container():
        st.markdown('<div class="date-box"><div class="date-row">', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 1], gap="small")
        with c1:
            g_date_from = st.date_input("من تاريخ", value=None, key="g_from", format="YYYY-MM-DD")
        with c2:
            g_date_to = st.date_input("إلى تاريخ", value=None, key="g_to", format="YYYY-MM-DD")
        st.markdown('</div></div>', unsafe_allow_html=True)

    # =======================
    # Financial Report Mode
    # =======================
    if type_key == "financial_report":
        # Summary (single row)
        df_summary = fetch_contract_summary_view(conn, company_name, project_name)
        if df_summary.empty:
            st.warning("لم يتم العثور على ملخص العقد لهذا المشروع.")
            return
        row = df_summary.iloc[0]

        header_company = company_name or "—"
        header_project = project_name or "—"
        header_date = str(row.get("تاريخ التعاقد", "—"))
        st.markdown(
            f"""
            <div class="fin-head">
                <div class="line">
                    <strong>الشركة:</strong> {header_company}
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong>المشروع:</strong> {header_project}
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong>تاريخ التعاقد:</strong> {header_date}
                </div>
                <span class="badge">تقرير مالي</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Build pairs from data only
        summary_pairs = _row_to_pairs_from_data(row)
        if summary_pairs:
            left_items, right_items = _split_pairs_two_columns(summary_pairs)
            st.markdown('<h3 class="hsec">ملخص المشروع</h3>', unsafe_allow_html=True)
            fin_panel_two_tables(left_items=left_items, right_items=right_items)
        else:
            st.info("لا توجد حقول قابلة للعرض في ملخص العقد.")

        # Downloads (summary only) — with wide logo in Excel
        df_summary_out = df_summary.copy()
        xlsx_sum = make_excel_bytes(df_summary_out, sheet_name="ملخص", put_logo=True)
        if xlsx_sum:
            st.download_button(
                label="تنزيل الملخص كـ Excel",
                data=xlsx_sum,
                file_name=_safe_filename(f"ملخص_{company_name}_{project_name}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        pdf_sum = make_pdf_bytes(_format_numbers_for_display(df_summary_out))
        st.download_button(
            label="تنزيل الملخص كـ PDF (مُصغّر واضح)",
            data=pdf_sum,
            file_name=_safe_filename(f"ملخص_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )

        st.markdown("---")

        # Ledger (v_financial_flow) with GLOBAL dates
        st.markdown('<h3 class="hsec">دفتر التدفق (v_financial_flow)</h3>', unsafe_allow_html=True)
        df_flow = fetch_financial_flow_view(conn, company_name, project_name, g_date_from, g_date_to)
        if df_flow.empty:
            st.info("لا توجد حركات مطابقة ضمن النطاق المحدد.")
            return

        col_search, term = create_column_search(df_flow)
        if col_search and term:
            df_flow = df_flow[df_flow[col_search].astype(str).str.contains(str(term), case=False, na=False)]
            if df_flow.empty:
                st.info("لا توجد نتائج بعد تطبيق البحث.")
                return

        # On-screen: hide IDs
        df_flow_display = df_flow.drop(columns=["companyid", "contractid"], errors="ignore")
        st.dataframe(df_flow_display, use_container_width=True, hide_index=True)

        # Individual downloads (EXCEL + wide logo)
        xlsx_flow = make_excel_bytes(df_flow_display, sheet_name="دفتر_التدفق", put_logo=True)
        if xlsx_flow:
            st.download_button(
                label="تنزيل الدفتر كـ Excel",
                data=xlsx_flow,
                file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        csv_flow = make_csv_utf8(df_flow_display)
        st.download_button(
            label="تنزيل الدفتر كـ CSV",
            data=csv_flow,
            file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.csv"),
            mime="text/csv",
        )
        pdf_flow = make_pdf_bytes(_format_numbers_for_display(df_flow_display))
        st.download_button(
            label="تنزيل الدفتر كـ PDF (مُصغّر واضح)",
            data=pdf_flow,
            file_name=_safe_filename(f"دفتر_التدفق_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )

        # Combined downloads (Excel: both with wide logo at top)
        st.markdown("### تنزيل تقرير موحّد")
        excel_two_sheets = make_excel_combined_two_sheets({
            "ملخص": df_summary_out,
            "دفتر_التدفق": df_flow_display,
        }, put_logo=True)
        if excel_two_sheets:
            st.download_button(
                label="Excel موحّد (ورقتان: ملخص + دفتر)",
                data=excel_two_sheets,
                file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}_ورقتين.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        excel_one_sheet = make_excel_single_sheet_stacked({
            "ملخص": df_summary_out,
            "دفتر_التدفق": df_flow_display,
        }, sheet_name="تقرير_موحد", put_logo=True)
        if excel_one_sheet:
            st.download_button(
                label="Excel موحّد (ورقة واحدة)",
                data=excel_one_sheet,
                file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}_ورقة_واحدة.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        header_line = f"الشركة: {company_name} | المشروع: {project_name} | تاريخ التعاقد: {row.get('تاريخ التعاقد','—')}"
        pdf_all = make_pdf_combined(
            _format_numbers_for_display(df_summary_out),
            _format_numbers_for_display(df_flow_display),
            header_text=header_line,
        )
        st.download_button(
            label="PDF موحّد (ملخص + دفتر) – مُصغّر واضح",
            data=pdf_all,
            file_name=_safe_filename(f"تقرير_مالي_{company_name}_{project_name}.pdf"),
            mime="application/pdf",
        )
        return

    # =======================
    # Classic table mode (contracts/guarantees/invoices/checks)
    # =======================
    df = fetch_data(conn, company_name, project_name, type_key)
    if df.empty:
        st.warning("لا توجد بيانات مطابقة للاختيارات المحددة.")
        return

    # Apply GLOBAL date filter to any date-like columns
    df = _apply_date_filter(df, g_date_from, g_date_to)

    search_column, search_term = create_column_search(df)
    if search_column and search_term:
        df = df[df[search_column].astype(str).str.contains(str(search_term), case=False, na=False)]
        if df.empty:
            st.info("لا توجد نتائج بعد تطبيق معيار البحث.")
            return

    column_config = {}
    for col in df.columns:
        if "رابط" in str(col):
            column_config[col] = st.column_config.LinkColumn(label=col, display_text="فتح الرابط")

    st.markdown('<h3 class="hsec">البيانات</h3>', unsafe_allow_html=True)
    st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

    xlsx_bytes = make_excel_bytes(df, sheet_name="البيانات", put_logo=True)
    if xlsx_bytes is not None:
        st.download_button(
            label="تنزيل كـ Excel (XLSX) – مُوصى به",
            data=xlsx_bytes,
            file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    csv_bytes = make_csv_utf8(df)
    st.download_button(
        label="تنزيل كـ CSV (UTF-8)",
        data=csv_bytes,
        file_name=_safe_filename(f"{type_key}_{company_name}_{project_name}.csv"),
        mime="text/csv",
    )

    pdf_title = _safe_filename(f"{type_key}_{company_name}_{project_name}")
    pdf_bytes = make_pdf_bytes(_format_numbers_for_display(df))
    st.download_button(
        label="تنزيل كـ PDF (مُصغّر واضح)",
        data=pdf_bytes,
        file_name=f"{pdf_title}.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    main()
