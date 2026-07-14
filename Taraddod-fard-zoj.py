#!/usr/bin/env python3
"""
Excel Processor App
───────────────────
1. Flatten multi-level column headers
2. Handle merged cells (forward-fill)
3. Merge columns with space separator
4. Filter incomplete entry/exit rows
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ═══════════════════════════════════════════════════════════
# Page Config — MUST be first Streamlit command
# ═══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="پردازشگر فایل اکسل",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════
# RTL Styling
# ═══════════════════════════════════════════════════════════
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;700&display=swap');

    html, body, [class*="st-"]  {
        font-family: 'Vazirmatn', 'Tahoma', 'Arial', sans-serif !important;
    }

    .stApp                        { direction: rtl; }
    .stMarkdown, .stCaption,
    .stSubheader, .stHeader,
    h1, h2, h3, p, span, label,
    .stSelectbox label,
    .stMultiselect label,
    .stNumberInput label,
    .stTextInput label,
    .stFileUploader label,
    .stRadio label                 { direction: rtl; text-align: right; }

    .stDataFrame                   { direction: ltr; text-align: left; }
    .stButton > button             { font-family: 'Vazirmatn', 'Tahoma', sans-serif; width: 100%; }
    [data-testid="stSidebar"]      { direction: rtl; text-align: right; }

    /* fix download button direction */
    .stDownloadButton > button     { direction: rtl; font-family: 'Vazirmatn', 'Tahoma', sans-serif; }
</style>
""",
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════
# Utility Functions
# ═══════════════════════════════════════════════════════════

def flatten_multiindex_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    """
    Remove parent column headers, promote sub-column names to top level.

    Before:  ('a','aa') | ('a','bb') | ('a','cc') | ('b',NaN) | ('c',NaN)
    After :  'aa'       | 'bb'       | 'cc'       | 'b'       | 'c'
    """
    if not isinstance(df.columns, pd.MultiIndex):
        return df, False

    new_columns: list[str] = []
    for col in df.columns:
        parts = [
            str(c).strip()
            for c in col
            if pd.notna(c) and str(c).strip()
        ]
        if len(parts) > 1:
            # Has sub-column → use lowest-level name
            new_columns.append(parts[-1])
        elif len(parts) == 1:
            new_columns.append(parts[0])
        else:
            new_columns.append(f"col_{len(new_columns)}")

    # Resolve duplicate names
    seen: dict[str, int] = {}
    final: list[str] = []
    for c in new_columns:
        if c in seen:
            seen[c] += 1
            final.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            final.append(c)

    df.columns = final
    return df, True


def safe_forward_fill(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Forward-fill selected columns (handles merged-cell NaN patterns)."""
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = df[col].astype(object).ffill()
    return df


def safe_merge_columns(
    df: pd.DataFrame,
    columns_to_merge: list[str],
    new_name: str,
    separator: str = " ",
) -> pd.DataFrame:
    """
    Merge multiple columns into one.
    Example:  "علی" + "حسینی"  →  "علی حسینی"
    NaN / empty values are skipped so no double-spaces appear.
    """
    df = df.copy()
    str_parts: list[pd.Series] = []

    for col in columns_to_merge:
        if col in df.columns:
            s = df[col].astype(str)
            # Normalise all NaN representations to empty string
            s = s.replace({"nan": "", "NaT": "", "None": "", "none": ""})
            str_parts.append(s)
        else:
            str_parts.append(pd.Series([""] * len(df), index=df.index))

    merged = str_parts[0]
    for part in str_parts[1:]:
        merged = merged.str.cat(part, sep=separator, na_rep="")

    # Trim leading/trailing separators, turn empty → NaN
    merged = merged.str.strip(separator)
    merged = merged.replace("", np.nan)

    # Drop originals, insert new column
    cols_to_drop = [c for c in columns_to_merge if c in df.columns]
    df = df.drop(columns=cols_to_drop)
    df.insert(0, new_name, merged)          # put merged col at the front
    return df


def filter_incomplete_rows(
    df: pd.DataFrame, entry_col: str, exit_col: str
) -> pd.DataFrame:
    """
    Keep only rows where *either* entry or exit is missing / empty.
    """
    if entry_col not in df.columns or exit_col not in df.columns:
        raise ValueError("یکی یا هر دو ستون انتخابی در داده‌ها وجود ندارند.")

    def _is_empty(series: pd.Series) -> pd.Series:
        s_str = series.astype(str).str.strip()
        return series.isna() | s_str.isin(["", "nan", "None", "NaT", "none"])

    mask = _is_empty(df[entry_col]) | _is_empty(df[exit_col])
    return df[mask].reset_index(drop=True)


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Export DataFrame to in-memory .xlsx bytes."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1", na_rep="")
    buf.seek(0)
    return buf.getvalue()


def read_excel_safely(file_obj, header_rows: int) -> pd.DataFrame:
    """Try openpyxl first (xlsx), fall back to xlrd (xls)."""
    header_arg = list(range(header_rows)) if header_rows > 1 else 0
    try:
        return pd.read_excel(file_obj, header=header_arg, engine="openpyxl")
    except Exception:
        file_obj.seek(0)
        return pd.read_excel(file_obj, header=header_arg, engine="xlrd")


# ═══════════════════════════════════════════════════════════
# Session State
# ═══════════════════════════════════════════════════════════
if "df" not in st.session_state:
    st.session_state.df = None
if "original_df" not in st.session_state:
    st.session_state.original_df = None
if "history" not in st.session_state:
    st.session_state.history: list[pd.DataFrame] = []
if "step_log" not in st.session_state:
    st.session_state.step_log: list[str] = []


def _push_history(label: str):
    """Record current df in history & log the step."""
    st.session_state.history.append(st.session_state.df.copy())
    st.session_state.step_log.append(label)


# ═══════════════════════════════════════════════════════════
# Sidebar — All Controls
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.header("⚙️ پنل عملیات")

    # ── Upload ──────────────────────────────────────────
    st.subheader("📤 آپلود فایل")
    uploaded = st.file_uploader(
        "فایل اکسل را انتخاب کنید",
        type=["xlsx", "xls"],
        key="uploader",
    )
    header_rows = st.number_input(
        "تعداد سطرهای سرستون",
        min_value=1, max_value=5, value=1,
        help="اگر فایل سرستون چندسطحی دارد (ستون اصلی + زیرستون)، تعداد سطرهای سرستون را وارد کنید.",
        key="n_header",
    )
    if uploaded and st.button("📂 خواندن فایل", use_container_width=True, key="btn_read"):
        try:
            with st.spinner("در حال خواندن…"):
                df = read_excel_safely(uploaded, header_rows)
                st.session_state.original_df = df.copy()
                st.session_state.df = df.copy()
                st.session_state.history = [df.copy()]
                st.session_state.step_log = ["خواندن فایل اصلی"]
            st.success("✅ فایل با موفقیت خوانده شد!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ خطا در خواندن فایل:\n`{e}`")

    # ── Step 1: Flatten ─────────────────────────────────
    st.divider()
    st.subheader("۱️⃣ یکسان‌سازی سرستون‌ها")
    if st.session_state.df is not None:
        if isinstance(st.session_state.df.columns, pd.MultiIndex):
            st.info("🔒 سرستون‌های چندسطحی شناسایی شد!")
            if st.button("یکسان‌سازی", use_container_width=True, key="btn_flat"):
                st.session_state.df, _ = flatten_multiindex_columns(st.session_state.df)
                _push_history("یکسان‌سازی سرستون‌ها")
                st.success("✅ انجام شد!")
                st.rerun()
        else:
            st.success("✅ سرستون‌ها یکسان هستند.")
    else:
        st.caption("ابتدا فایل را آپلود کنید.")

    # ── Step 2: Forward-Fill ────────────────────────────
    st.divider()
    st.subheader("۲️⃣ پر کردن سلول‌های ادغام‌شده")
    if st.session_state.df is not None:
        st.caption(
            "ستون‌هایی که در اکسل ادغام (Merge) شده‌اند را انتخاب کنید. "
            "مقدار سلول بالایی در ردیف‌های خالی تکرار می‌شود."
        )
        ff_cols = st.multiselect(
            "انتخاب ستون‌ها",
            st.session_state.df.columns.tolist(),
            key="ff_cols",
        )
        if st.button("پر کردن مقادیر خالی", use_container_width=True, key="btn_ff"):
            if ff_cols:
                st.session_state.df = safe_forward_fill(st.session_state.df, ff_cols)
                _push_history(f"پر کردن: {', '.join(ff_cols)}")
                st.success("✅ انجام شد!")
                st.rerun()
            else:
                st.warning("حداقل یک ستون انتخاب کنید.")
    else:
        st.caption("ابتدا فایل را آپلود کنید.")

    # ── Step 3: Merge Columns ───────────────────────────
    st.divider()
    st.subheader("۳️⃣ ادغام ستون‌ها")
    if st.session_state.df is not None:
        st.caption(
            "مثلاً ستون «نام» و «نام‌خانوادگی» را انتخاب کنید تا «علی حسینی» ساخته شود."
        )
        mg_cols = st.multiselect(
            "ستون‌ها برای ادغام (حداقل ۲)",
            st.session_state.df.columns.tolist(),
            key="mg_cols",
        )
        mg_name = st.text_input(
            "نام ستون جدید", placeholder="نام_کامل", key="mg_name"
        )
        if st.button("ادغام", use_container_width=True, key="btn_mg"):
            if len(mg_cols) < 2:
                st.warning("حداقل دو ستون انتخاب کنید.")
            elif not mg_name.strip():
                st.warning("نام ستون جدید را وارد کنید.")
            else:
                try:
                    st.session_state.df = safe_merge_columns(
                        st.session_state.df, mg_cols, mg_name.strip()
                    )
                    _push_history(f"ادغام → {mg_name.strip()}")
                    st.success("✅ انجام شد!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ خطا: {e}")
    else:
        st.caption("ابتدا فایل را آپلود کنید.")

    # ── Step 4: Filter Incomplete ───────────────────────
    st.divider()
    st.subheader("۴️⃣ فیلتر ردیف‌های ناقص")
    if st.session_state.df is not None:
        st.caption(
            "ردیف‌هایی که **ورود** یا **خروج** آن‌ها خالی / ناقص است فیلتر شده و فقط همان‌ها در خروجی باقی می‌مانند."
        )
        opts = st.session_state.df.columns.tolist()

        # Auto-detect entry / exit column names
        entry_idx, exit_idx = 0, 0
        for i, c in enumerate(opts):
            if "ورود" in str(c):
                entry_idx = i + 1
            if "خروج" in str(c):
                exit_idx = i + 1

        entry_col = st.selectbox("ستون «ورود»", [""] + opts, index=entry_idx, key="sel_in")
        exit_col = st.selectbox("ستون «خروج»", [""] + opts, index=exit_idx, key="sel_out")

        if st.button("فیلتر ردیف‌های ناقص", use_container_width=True, key="btn_filt"):
            if not entry_col or not exit_col:
                st.warning("هر دو ستون را انتخاب کنید.")
            else:
                try:
                    filtered = filter_incomplete_rows(st.session_state.df, entry_col, exit_col)
                    st.session_state.df = filtered
                    _push_history(f"فیلتر ناقص ({len(filtered)} ردیف)")
                    st.success(f"✅ {len(filtered)} ردیف ناقص باقی ماند.")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ خطا: {e}")
    else:
        st.caption("ابتدا فایل را آپلود کنید.")

    # ── Undo / Reset ────────────────────────────────────
    st.divider()
    st.subheader("🔧 ابزارها")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("↩️ بازگشت", use_container_width=True, key="btn_undo"):
            if len(st.session_state.history) > 1:
                st.session_state.history.pop()
                st.session_state.step_log.pop()
                st.session_state.df = st.session_state.history[-1].copy()
                st.rerun()
            else:
                st.info("مرحله‌ای نیست.")
    with c2:
        if st.button("🔄 شروع مجدد", use_container_width=True, key="btn_reset"):
            if st.session_state.original_df is not None:
                st.session_state.df = st.session_state.original_df.copy()
                st.session_state.history = [st.session_state.original_df.copy()]
                st.session_state.step_log = ["بازنشانی به فایل اصلی"]
                st.rerun()


# ═══════════════════════════════════════════════════════════
# Main Area — Preview & Download
# ═══════════════════════════════════════════════════════════
st.title("📋 پردازشگر فایل اکسل")
st.caption(
    "یکسان‌سازی سرستون‌ها  ·  پر کردن سلول‌های ادغام‌شده  "
    "·  ادغام ستون‌ها  ·  فیلتر ردیف‌های ناقص ورود/خروج"
)

if st.session_state.df is not None:
    df = st.session_state.df

    # ── Metrics ─────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("تعداد سطرها", f"{len(df):,}")
    m2.metric("تعداد ستون‌ها", len(df.columns))
    m3.metric("مقادیر خالی", f"{int(df.isna().sum().sum()):,}")
    m4.metric("مراحل انجام‌شده", len(st.session_state.step_log) - 1)

    # ── Step log ────────────────────────────────────────
    if len(st.session_state.step_log) > 1:
        with st.expander("📜 تاریخچه عملیات"):
            for i, step in enumerate(st.session_state.step_log):
                st.markdown(f"**{i}**. {step}")

    # ── Tabs: Current vs Original ───────────────────────
    tab_cur, tab_orig = st.tabs(["📊 داده‌های فعلی", "📄 فایل اصلی"])

    with tab_cur:
        st.dataframe(df, use_container_width=True, height=480)

    with tab_orig:
        if st.session_state.original_df is not None:
            st.dataframe(st.session_state.original_df, use_container_width=True, height=480)

    # ── Download ────────────────────────────────────────
    st.divider()
    st.subheader("📥 دانلود فایل نهایی")

    if len(df) == 0:
        st.warning(
            "⚠️ داده‌ای برای دانلود وجود ندارد — ممکن است همه ردیف‌ها فیلتر شده باشند. "
            "از دکمه «بازگشت» استفاده کنید."
        )
    else:
        xlsx_bytes = df_to_excel_bytes(df)
        csv_bytes = df.to_csv(index=False, na_rep="").encode("utf-8-sig")

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                label="📥 دانلود Excel (.xlsx)",
                data=xlsx_bytes,
                file_name="processed_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                label="📥 دانلود CSV (.csv)",
                data=csv_bytes,
                file_name="processed_data.csv",
                mime="text/csv",
                use_container_width=True,
            )

else:
    # ── Empty State ─────────────────────────────────────
    st.markdown(
        """
        <div style="
            text-align: center;
            padding: 100px 20px;
            border: 2px dashed #d0d0d0;
            border-radius: 16px;
            margin-top: 60px;
            background: #fafafa;
        ">
            <p style="font-size: 4rem; margin-bottom: 0;">📂</p>
            <h2 style="color: #888; margin-top: 10px;">
                فایل اکسل خود را از پنل سمت چپ آپلود کنید
            </h2>
            <p style="color: #aaa; margin-top: 8px;">
                فرمت‌های پشتیبانی‌شده: <b>.xlsx</b> و <b>.xls</b>
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
