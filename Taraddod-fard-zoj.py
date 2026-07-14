import streamlit as st
import pandas as pd
import numpy as np
import io
import warnings
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=UserWarning)

# ═══════════════════════════════════════════════════════════════
# ترتیب دلخواه ستون‌ها در خروجی نهایی
# ═══════════════════════════════════════════════════════════════
DESIRED_COLUMN_ORDER = [
    "غیبت روزانه",
    "پایان شیفت",
    "شروع شیفت",
    "عنوان گروه کاری",
    "روز هفته",
    "تاریخ",
    "نام خانوادگی",
    "نام",
    "وضعیت",
    "مدت زمان",
    "خروج",
    "ورود",
    "عنوان واحد سازمانی",
]


# ═══════════════════════════════════════════════════════════════
# بخش ۱: پارس و تمیزکاری اکسل
# ═══════════════════════════════════════════════════════════════

def _find_header_row(df: pd.DataFrame) -> int | None:
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).values
        row_text = " ".join(row_values)
        if "غیبت روزانه" in row_text and "نام خانوادگی" in row_text:
            return i
    return None


def _find_column_index(row: pd.Series, search_text: str) -> int | None:
    for idx, val in enumerate(row):
        if pd.notna(val) and search_text.strip() in str(val).strip():
            return idx
    return None


def _build_column_map(header_row1: pd.Series, header_row2: pd.Series) -> dict:
    col_map: dict[str, int] = {}
    main_columns = [
        "غیبت روزانه", "پایان شیفت", "شروع شیفت", "عنوان گروه کاری",
        "روز هفته", "تاریخ", "نام خانوادگی", "نام", "روز",
        "کد پرسنلی کارمند", "عنوان واحد سازمانی",
    ]
    for col_name in main_columns:
        idx = _find_column_index(header_row1, col_name)
        if idx is not None:
            col_map[col_name] = idx

    info_col_start = _find_column_index(header_row1, "اطلاعات تردد")
    if info_col_start is not None:
        for offset in range(8):
            col_idx = info_col_start + offset
            if col_idx >= len(header_row2):
                break
            val = header_row2.iloc[col_idx]
            if pd.notna(val):
                val_str = str(val).strip()
                if val_str and val_str not in ("*", "nan", "None", ""):
                    if val_str not in col_map:
                        col_map[val_str] = col_idx

    if not col_map:
        raise ValueError("هیچ ستونی تشخیص داده نشد. ساختار فایل را بررسی کنید.")
    return col_map


def _extract_columns(df_data: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    sorted_cols = sorted(col_map.items(), key=lambda x: x[1])
    df_clean = pd.DataFrame()
    for col_name, col_idx in sorted_cols:
        if col_idx < len(df_data.columns):
            df_clean[col_name] = df_data[col_idx].values
    return df_clean


def _clean_values(df: pd.DataFrame) -> pd.DataFrame:
    junk_values = {"*", "", "nan", "None", "none", "NaN"}
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: (
                np.nan
                if pd.isna(x) or str(x).strip() in junk_values
                else str(x).strip()
            )
        )
    return df


def _remove_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    name_cols = [c for c in ["نام", "نام خانوادگی"] if c in df.columns]
    if name_cols:
        before = len(df)
        df = df.dropna(subset=name_cols, how="all")
        removed = before - len(df)
        if removed > 0:
            print(f"[INFO] {removed} ردیف بدون نام حذف شد.")
    return df


def _drop_unwanted_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols_to_drop = [c for c in ["روز", "کد پرسنلی کارمند"] if c in df.columns]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)
    return df


def _forward_fill_personnel_code(df: pd.DataFrame) -> pd.DataFrame:
    required = ["عنوان واحد سازمانی", "نام", "نام خانوادگی"]
    if all(c in df.columns for c in required):
        df["عنوان واحد سازمانی"] = df.groupby(
            ["نام", "نام خانوادگی"], sort=False
        )["عنوان واحد سازمانی"].ffill()
        df["عنوان واحد سازمانی"] = df.groupby(
            ["نام", "نام خانوادگی"], sort=False
        )["عنوان واحد سازمانی"].bfill()
    return df


def _reorder_columns(df: pd.DataFrame) -> pd.DataFrame:
    ordered = [c for c in DESIRED_COLUMN_ORDER if c in df.columns]
    extra = [c for c in df.columns if c not in DESIRED_COLUMN_ORDER]
    return df[ordered + extra]


def parse_and_clean_excel(file) -> pd.DataFrame:
    df_raw = pd.read_excel(file, header=None)
    if df_raw.empty or len(df_raw) < 7:
        raise ValueError("فایل خالی است یا دارای داده‌های کافی نیست.")

    header_row_idx = _find_header_row(df_raw)
    if header_row_idx is None:
        raise ValueError(
            "ساختار فایل تشخیص داده نشد.\n"
            "فایل باید شامل ستون‌های 'غیبت روزانه' و 'نام خانوادگی' باشد."
        )

    header_row1 = df_raw.iloc[header_row_idx]
    header_row2 = df_raw.iloc[header_row_idx + 1]
    data_start_row = header_row_idx + 2

    col_index_map = _build_column_map(header_row1, header_row2)
    df_data = df_raw.iloc[data_start_row:].reset_index(drop=True)
    df_clean = _extract_columns(df_data, col_index_map)
    df_clean = _clean_values(df_clean)
    df_clean = _remove_empty_rows(df_clean)
    df_clean = _drop_unwanted_columns(df_clean)
    df_clean = _forward_fill_personnel_code(df_clean)
    df_clean = _reorder_columns(df_clean)
    df_clean = df_clean.reset_index(drop=True)
    return df_clean


# ═══════════════════════════════════════════════════════════════
# بخش ۲: منطق آنالیز و جستجو
# ═══════════════════════════════════════════════════════════════

def get_unique_persons(df: pd.DataFrame) -> list[str]:
    if "نام" not in df.columns or "نام خانوادگی" not in df.columns:
        return []
    persons = (
        df[["نام خانوادگی", "نام"]]
        .drop_duplicates()
        .apply(lambda r: f"{r['نام خانوادگی']} - {r['نام']}", axis=1)
        .tolist()
    )
    return persons


def get_person_dates(df: pd.DataFrame, person_label: str) -> list[str]:
    parts = person_label.split(" - ")
    if len(parts) != 2:
        return []
    family, name = parts[0].strip(), parts[1].strip()
    if "تاریخ" not in df.columns:
        return []
    mask = (df["نام خانوادگی"] == family) & (df["نام"] == name)
    dates = df.loc[mask, "تاریخ"].dropna().unique().tolist()
    return dates


def get_person_rows_for_date(
    df: pd.DataFrame, person_label: str, date: str
) -> pd.DataFrame:
    parts = person_label.split(" - ")
    if len(parts) != 2:
        return pd.DataFrame()
    family, name = parts[0].strip(), parts[1].strip()
    mask = (
        (df["نام خانوادگی"] == family)
        & (df["نام"] == name)
        & (df["تاریخ"] == date)
    )
    return df.loc[mask].reset_index(drop=True)


def find_similar_records(
    df: pd.DataFrame,
    reference_rows: pd.DataFrame,
    match_columns: list[str] | None = None,
) -> pd.DataFrame:
    if match_columns is None:
        match_columns = ["ورود", "خروج"]

    missing = [c for c in match_columns if c not in df.columns]
    if missing:
        raise ValueError(f"ستون‌های {missing} در داده‌ها یافت نشدند.")

    patterns: list[tuple] = []
    for _, row in reference_rows.iterrows():
        entry_val = row.get("ورود", np.nan)
        exit_val = row.get("خروج", np.nan)
        if pd.isna(entry_val) and pd.isna(exit_val):
            continue
        patterns.append((entry_val, exit_val))

    if not patterns:
        return pd.DataFrame()

    combined_mask = pd.Series(False, index=df.index)
    for entry_val, exit_val in patterns:
        row_mask = pd.Series(True, index=df.index)
        if pd.isna(entry_val):
            row_mask &= df["ورود"].isna()
        else:
            row_mask &= df["ورود"] == entry_val
        if pd.isna(exit_val):
            row_mask &= df["خروج"].isna()
        else:
            row_mask &= df["خروج"] == exit_val
        combined_mask |= row_mask

    result = df.loc[combined_mask].reset_index(drop=True)
    return result


# ═══════════════════════════════════════════════════════════════
# بخش ۳: توابع کمکی UI
# ═══════════════════════════════════════════════════════════════

def _reset_session():
    keys_to_keep = {"file_upload"}
    for key in list(st.session_state.keys()):
        if key not in keys_to_keep:
            del st.session_state[key]


def _get_active_renames() -> dict[str, str]:
    renames = st.session_state.get("column_renames", {})
    return {k: v for k, v in renames.items() if k != v and v.strip()}


def _apply_renames(df: pd.DataFrame) -> pd.DataFrame:
    actual = _get_active_renames()
    if actual:
        return df.rename(columns=actual)
    return df.copy()


def _export_to_excel(df: pd.DataFrame, filename: str = "result.xlsx") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="نتیجه")
        worksheet = writer.sheets["نتیجه"]
        for i, col in enumerate(df.columns):
            col_data = df[col].astype(str)
            max_len = max(col_data.apply(len).max(), len(str(col))) + 3
            max_len = min(max_len, 55)
            col_letter = get_column_letter(i + 1)
            worksheet.column_dimensions[col_letter].width = max_len
    buffer.seek(0)
    return buffer.getvalue()


# ═══════════════════════════════════════════════════════════════
# بخش ۴: اپلیکیشن اصلی
# ═══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="آنالیزور غیبت و تردد",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DARK_THEME_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;600;700&display=swap');
    * {
        font-family: 'Vazirmatn', 'Tahoma', 'Arial', sans-serif !important;
        direction: rtl;
        text-align: right;
    }
    .stApp { background-color: #0E1117; color: #FAFAFA; }
    h1, h2, h3, h4, h5, h6 {
        color: #E8E8E8 !important;
        font-family: 'Vazirmatn', sans-serif !important;
        text-align: right !important;
    }
    p, span, label, div { text-align: right; }
    .stButton > button[kind="primary"],
    .stButton > button:not([kind]) {
        background: linear-gradient(135deg, #FF4B4B, #E53935);
        color: white; font-weight: 600; border: none; border-radius: 10px;
        padding: 12px 28px; font-size: 15px; transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(255, 75, 75, 0.25);
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #FF6B6B, #FF5252);
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(255, 75, 75, 0.45);
    }
    .stButton > button:active { transform: translateY(0); }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #00C853, #00E676);
        color: #0E1117; font-weight: 700; border: none; border-radius: 10px;
        padding: 12px 28px; font-size: 15px; transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(0, 200, 83, 0.25);
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #69F0AE, #B9F6CA);
        color: #1B5E20; transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0, 200, 83, 0.45);
    }
    div[data-testid="stSidebar"] {
        background-color: #161822; border-left: 1px solid #2A2D35;
    }
    div[data-testid="stSidebar"] * { color: #E0E0E0; }
    div[data-testid="stMetricValue"] {
        color: #FFFFFF !important; font-size: 28px !important; font-weight: 700 !important;
    }
    div[data-testid="stMetricLabel"] {
        color: #9E9E9E !important; font-size: 13px !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px; background-color: #161822; border-radius: 12px; padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px; padding: 10px 20px; color: #9E9E9E;
        font-weight: 500; background-color: transparent;
    }
    .stTabs [aria-selected="true"] {
        background-color: #262730 !important; color: #FFFFFF !important; font-weight: 700;
    }
    .stSelectbox label, .stTextInput label { color: #E0E0E0 !important; font-weight: 600; }
    div[data-baseweb="select"] { background-color: #262730 !important; }
    div[data-baseweb="input"] { background-color: #262730 !important; }
    .stDataFrame {
        border-radius: 10px; overflow: hidden; border: 1px solid #2A2D35;
    }
    .dataframe th {
        background-color: #1E2028 !important; color: #E0E0E0 !important;
        font-weight: 700 !important; border-bottom: 2px solid #FF4B4B !important;
        text-align: right !important; padding: 10px 12px !important;
    }
    .dataframe td {
        color: #E8E8E8 !important; text-align: right !important;
        padding: 8px 12px !important; border-bottom: 1px solid #1E2028 !important;
    }
    .dataframe tr:hover { background-color: #262730 !important; }
    [data-testid="stFileUploader"] section {
        background-color: #1A1D23; border: 2px dashed #3A3D45;
        border-radius: 12px; padding: 30px; transition: all 0.3s ease;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #FF4B4B; background-color: #1E2028;
    }
    hr { border-color: #2A2D35; margin: 20px 0; }
    .stCaption { color: #757575 !important; }
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #1A1D23; }
    ::-webkit-scrollbar-thumb { background: #3A3D45; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #4A4D55; }
</style>
"""

st.markdown(DARK_THEME_CSS, unsafe_allow_html=True)


def main():
    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 10px;">
            <h1 style="margin-bottom: 4px;">📊 آنالیزور غیبت و تردد</h1>
            <p style="color: #757575; font-size: 14px; margin-top: 0;">
                تمیزکاری، استانداردسازی و تحلیل فایل‌های اکسل حضور و غیاب پرسنل
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<hr style='border-color: #2A2D35;'>", unsafe_allow_html=True)

    # ── آپلود فایل ──
    st.subheader("📁 آپلود فایل اکسل")
    uploaded_file = st.file_uploader(
        "فایل اکسل (.xlsx یا .xls) خود را انتخاب کنید",
        type=["xlsx", "xls"],
        key="file_upload",
        label_visibility="collapsed",
    )

    if uploaded_file is None:
        st.markdown(
            """
            <div style="
                background: linear-gradient(135deg, #0D47A1, #1565C0);
                color: white; padding: 16px 24px; border-radius: 10px; margin: 16px 0;
            ">
                <strong>📌 راهنمای استفاده:</strong><br><br>
                ۱. فایل اکسل غیبت را آپلود کنید<br>
                ۲. در تب «پیش‌نمایش» داده‌های تمیز شده را ببینید<br>
                ۳. در تب «تغییر نام» نام ستون‌ها را اصلاح کنید<br>
                ۴. در تب «آنالیز» فرد و تاریخ را انتخاب و جستجو کنید
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # ── پردازش فایل ──
    if (
        "df_clean" not in st.session_state
        or st.session_state.get("_last_file") != uploaded_file.name
    ):
        _reset_session()
        with st.spinner("در حال خواندن و تمیزکاری فایل..."):
            try:
                df_clean = parse_and_clean_excel(uploaded_file)
                st.session_state["df_clean"] = df_clean
                st.session_state["_last_file"] = uploaded_file.name
                st.session_state["column_renames"] = {
                    col: col for col in df_clean.columns
                }
            except ValueError as e:
                st.error(f"❌ خطا در پردازش فایل:\n\n{str(e)}")
                return
            except Exception as e:
                st.error(f"❌ خطای غیرمنتظره:\n\n{str(e)}")
                return

    df_clean = st.session_state["df_clean"]

    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #1B5E20, #2E7D32);
            color: white; padding: 14px 24px; border-radius: 10px; margin: 8px 0 16px 0;
        ">
            ✅ فایل <strong>{uploaded_file.name}</strong> با موفقیت پردازش شد —
            <strong>{len(df_clean):,}</strong> ردیف و
            <strong>{len(df_clean.columns)}</strong> ستون استخراج شد
        </div>
        """,
        unsafe_allow_html=True,
    )

    mc1, mc2, mc3, mc4 = st.columns(4)
    with mc1:
        st.metric("تعداد ردیف‌ها", f"{len(df_clean):,}")
    with mc2:
        unique_persons = df_clean[["نام", "نام خانوادگی"]].drop_duplicates()
        st.metric("تعداد افراد", f"{len(unique_persons):,}")
    with mc3:
        if "تاریخ" in df_clean.columns:
            st.metric("تعداد تاریخ‌ها", f"{df_clean['تاریخ'].dropna().nunique():,}")
    with mc4:
        if "وضعیت" in df_clean.columns:
            present_count = df_clean["وضعیت"].dropna().str.contains("حضور").sum()
            st.metric("ردیف‌های با حضور", f"{present_count:,}")

    st.markdown("<hr style='border-color: #2A2D35;'>", unsafe_allow_html=True)

    # ── تب‌ها ──
    tab_preview, tab_rename, tab_analyze = st.tabs(
        ["📋 پیش‌نمایش داده‌ها", "✏️ تغییر نام ستون‌ها", "🔍 آنالیز و جستجو"]
    )

    # ────── تب ۱: پیش‌نمایش ──────
    with tab_preview:
        st.subheader("داده‌های تمیز و استاندارد شده")
        display_df = _apply_renames(df_clean)
        st.dataframe(display_df, use_container_width=True, height=480, hide_index=True)
        st.download_button(
            label="📥 دانلود اکسل تمیز شده",
            data=_export_to_excel(display_df, "cleaned_data.xlsx"),
            file_name="cleaned_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ────── تب ۲: تغییر نام ──────
    with tab_rename:
        st.subheader("تغییر نام ستون‌ها")
        st.caption("نام دلخواه خود را وارد کنید. فیلدهای خالی نام اصلی را حفظ می‌کنند.")
        current_renames = st.session_state.get("column_renames", {c: c for c in df_clean.columns})
        new_renames: dict[str, str] = {}
        columns_list = list(df_clean.columns)
        cols_per_row = 3
        for i in range(0, len(columns_list), cols_per_row):
            row_cols = st.columns(cols_per_row)
            for j, col in enumerate(columns_list[i : i + cols_per_row]):
                with row_cols[j]:
                    st.markdown(
                        f"<span style='color:#9E9E9E; font-size:12px;'>"
                        f"نام اصلی: <strong>{col}</strong></span>",
                        unsafe_allow_html=True,
                    )
                    new_name = st.text_input(
                        label=f"نام جدید برای {col}",
                        value=current_renames.get(col, col),
                        key=f"rename_{col}",
                        label_visibility="collapsed",
                        placeholder=col,
                    )
                    new_renames[col] = new_name
        bc1, bc2 = st.columns([1, 4])
        with bc1:
            if st.button("💾 اعمال تغییرات", type="primary"):
                st.session_state["column_renames"] = new_renames
                st.success("✅ تغییرات ذخیره شد.")
                st.rerun()

    # ────── تب ۳: آنالیز ──────
    with tab_analyze:
        st.subheader("آنالیز و جستجوی موارد مشابه")
        st.markdown(
            """
            <div style="
                background: linear-gradient(135deg, #0D47A1, #1565C0);
                color: white; padding: 16px 24px; border-radius: 10px; margin-bottom: 20px;
            ">
                <strong>🔎 روش کار:</strong><br><br>
                ۱. یک نفر را از لیست انتخاب کنید<br>
                ۲. تاریخ مورد نظر را انتخاب کنید<br>
                ۳. رکوردهای تردد آن شخص در آن تاریخ نمایش داده می‌شود<br>
                ۴. با کلیک بر دکمه «جستجوی موارد مشابه»، تمام ردیف‌هایی در کل
                فایل که <strong>ورود و خروج یکسانی</strong> دارند پیدا می‌شوند<br>
                ۵. نتایج قابل دانلود هستند
            </div>
            """,
            unsafe_allow_html=True,
        )

        persons = get_unique_persons(df_clean)
        if not persons:
            st.warning("⚠️ هیچ فردی در داده‌ها یافت نشد.")
            return

        cs1, cs2 = st.columns(2)
        with cs1:
            selected_person = st.selectbox(
                "👤 انتخاب فرد:", options=persons, key="person_select",
                index=None, placeholder="-- فرد مورد نظر را انتخاب کنید --",
            )

        selected_date = None
        if selected_person:
            dates = get_person_dates(df_clean, selected_person)
            if dates:
                with cs2:
                    selected_date = st.selectbox(
                        "📅 انتخاب تاریخ:", options=dates, key="date_select",
                        index=None, placeholder="-- تاریخ را انتخاب کنید --",
                    )
            else:
                st.warning("⚠️ هیچ تاریخی برای این فرد یافت نشد.")

        if selected_person and selected_date:
            person_rows = get_person_rows_for_date(df_clean, selected_person, selected_date)
            if person_rows.empty:
                st.warning("⚠️ رکوردی برای این فرد در تاریخ انتخاب شده یافت نشد.")
            else:
                display_rows = _apply_renames(person_rows)
                st.markdown("#### رکوردهای تردد فرد در تاریخ انتخاب شده")
                st.dataframe(
                    display_rows, use_container_width=True,
                    height=min(200, 50 + len(display_rows) * 35), hide_index=True,
                )

                if "ورود" in person_rows.columns and "خروج" in person_rows.columns:
                    patterns = person_rows[["ورود", "خروج"]].dropna(how="all")
                    if not patterns.empty:
                        st.markdown("#### الگوهای ورود / خروج")
                        html_parts = []
                        for _, row in patterns.iterrows():
                            ev = row["ورود"] if pd.notna(row["ورود"]) else "—"
                            xv = row["خروج"] if pd.notna(row["خروج"]) else "—"
                            html_parts.append(
                                f"<span style='display:inline-block; background:#262730; "
                                f"padding:8px 18px; border-radius:8px; margin:4px 4px 4px 0; "
                                f"border-right: 3px solid #FF4B4B;'>"
                                f"⏰ ورود: <strong>{ev}</strong>&nbsp;&nbsp;|&nbsp;&nbsp;"
                                f"⏰ خروج: <strong>{xv}</strong></span>"
                            )
                        st.markdown("".join(html_parts), unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
                search_clicked = st.button(
                    "🔍 جستجوی موارد مشابه در کل فایل",
                    type="primary", use_container_width=True,
                )
                st.markdown("</div>", unsafe_allow_html=True)

                if search_clicked:
                    with st.spinner("در حال جستجوی ردیف‌های مشابه..."):
                        try:
                            results = find_similar_records(df_clean, person_rows)
                            st.session_state["analysis_results"] = results
                        except ValueError as e:
                            st.error(f"❌ {str(e)}")
                            return

                    if results.empty:
                        st.warning("⚠️ هیچ ردیف مشابهی در کل فایل یافت نشد.")
                    else:
                        st.session_state["analysis_results"] = results

        # ── نمایش نتایج ──
        if "analysis_results" in st.session_state:
            results = st.session_state["analysis_results"]
            if not results.empty:
                display_results = _apply_renames(results)
                st.markdown("---")
                st.markdown(
                    f"#### 📊 نتایج جستجو — "
                    f"<span style='color:#FF4B4B;'>{len(results):,} ردیف مشابه</span> یافت شد",
                    unsafe_allow_html=True,
                )
                rc1, rc2, rc3 = st.columns(3)
                with rc1:
                    st.metric("ردیف‌های مشابه", f"{len(results):,}")
                with rc2:
                    if "نام" in results.columns and "نام خانوادگی" in results.columns:
                        mp = results[["نام", "نام خانوادگی"]].drop_duplicates()
                        st.metric("افراد مشابه", f"{len(mp):,}")
                with rc3:
                    if "تاریخ" in results.columns:
                        st.metric("تاریخ‌های مختلف", f"{results['تاریخ'].nunique():,}")

                st.dataframe(display_results, use_container_width=True, height=450, hide_index=True)
                st.download_button(
                    label="📥 دانلود نتایج جستجو (Excel)",
                    data=_export_to_excel(display_results, "analysis_results.xlsx"),
                    file_name="analysis_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                if st.button("🗑️ پاک کردن نتایج جستجو"):
                    del st.session_state["analysis_results"]
                    st.rerun()


if __name__ == "__main__":
    main()
