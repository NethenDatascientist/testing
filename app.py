import io
import os
import pickle
import re
import threading
import time
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION (MOBILE PORTRAIT OPTIMIZED)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MI Overdue Works — CE Mobile View",
    page_icon="📱",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Mobile-first Portrait CSS (2x2 KPI cards, tight padding, thumb-friendly UI)
st.markdown(
    """
    <style>
        .block-container {
            max-width: 680px !important;
            padding-left: 0.65rem !important;
            padding-right: 0.65rem !important;
            padding-top: 1.8rem !important;
            padding-bottom: 2.5rem !important;
        }
        /* Compact Title for Mobile */
        .mobile-title {
            font-size: 1.25rem;
            font-weight: 800;
            line-height: 1.3;
            margin-bottom: 0.2rem;
        }
        .mobile-subtitle {
            font-size: 0.78rem;
            opacity: 0.75;
            margin-bottom: 0.75rem;
        }
        /* 2x2 Responsive KPI Grid for Portrait Phones */
        .kpi-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin-top: 6px;
            margin-bottom: 12px;
        }
        .kpi-card {
            background: rgba(148, 163, 184, 0.10);
            border: 1px solid rgba(148, 163, 184, 0.25);
            border-radius: 10px;
            padding: 10px 12px;
        }
        .kpi-card-red {
            background: rgba(239, 68, 68, 0.14);
            border: 1px solid rgba(239, 68, 68, 0.45);
            border-radius: 10px;
            padding: 10px 12px;
        }
        .kpi-card-green {
            background: rgba(74, 222, 128, 0.12);
            border: 1px solid rgba(74, 222, 128, 0.40);
            border-radius: 10px;
            padding: 10px 12px;
        }
        .kpi-label {
            font-size: 0.74rem;
            font-weight: 600;
            opacity: 0.85;
            margin-bottom: 2px;
        }
        .kpi-val {
            font-size: 1.22rem;
            font-weight: 800;
            line-height: 1.2;
        }
        /* Legend pill bar */
        .legend-bar {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            font-size: 0.78rem;
            font-weight: 600;
            margin: 6px 0 4px 0;
            padding: 6px 10px;
            background: rgba(148, 163, 184, 0.08);
            border-radius: 8px;
        }
        .dot-red { color: #EF4444; }
        .dot-green { color: #4ADE80; }
        .dot-grey { color: #94A3B8; }
        /* Make radio & select widgets easier to tap on mobile */
        div[role="radiogroup"] {
            gap: 0.4rem !important;
        }
        [data-testid="stPlotlyChart"] {
            width: 100% !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# CONSTANTS & TARGET SHEETS
# -----------------------------------------------------------------------------
TARGET_SHEETS = [
    "4702_MOT",
    "4702-New Tanks",
    "4702-LIS",
    "4702_A&P",
    "PMKSY-SMI-RRR",
    "4702_Paschima Vahini",
    "4711__Other Flood Control",
    "4702-Kharland",
    "NABARD - New Tank",
    "SDMF",
    "KMERC",
    "NABARD - A&P",
    "ATAL BHUJAL",
    "NABARD_MOT",
    "NABARD - LIS",
    "other dept funds",
    "SCP (LIS)",
    "SCP (Borewell)",
    "SCP (Checkdam)",
    "TSP (Checkdam)",
    "TSP (Borewell)",
    "TSP (LIS)",
    "NABARD - Other Flood Control",
]

COL_SO = "ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು"
COL_AEE = "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು"
COL_TARGET_DATE = "ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ"
COL_ACTUAL_DATE = "ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ"
COL_STATUS_STD = "CP/CF/P/TBS/TP/ETS/OTH"

DEFAULT_SHEET_URL = st.secrets.get(
    "SHEET_URL",
    "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869",
)

DISK_BACKUP_FILE = "/tmp/mi_overdue_shared_snapshot_v3.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK (Instant load on CE's phone)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_shared_state() -> dict:
    state = {
        "lock": threading.Lock(),
        "data": None,
        "last_synced": None,
        "source_label": None,
    }
    if os.path.exists(DISK_BACKUP_FILE):
        try:
            with open(DISK_BACKUP_FILE, "rb") as f:
                saved = pickle.load(f)
                state["data"] = saved.get("data")
                state["last_synced"] = saved.get("last_synced")
                state["source_label"] = saved.get("source_label")
        except Exception:
            pass
    return state


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def normalize_text(val) -> str:
    if val is None or pd.isna(val):
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def format_cell_for_display(val) -> str:
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d-%m-%Y")
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return normalize_text(val)


def extract_spreadsheet_key(sheet_url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("Invalid Google Sheet URL.")
    return match.group(1)


def download_xlsx_bytes(sheet_url: str) -> bytes:
    key = extract_spreadsheet_key(sheet_url)
    urls = [
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx",
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx&id={key}",
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Linux; Android 14; Pixel 8) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/128.0.0.0 Mobile Safari/537.36"
        ),
    }

    last_err = None
    for attempt in range(3):
        url = urls[attempt % len(urls)]
        if attempt > 0:
            url = f"{url}&t={int(time.time())}"
        try:
            resp = requests.get(url, headers=headers, timeout=120)
            resp.raise_for_status()
            if len(resp.content) > 1000:
                return resp.content
        except Exception as e:
            last_err = e
            time.sleep(2 * (attempt + 1))

    raise RuntimeError(f"Could not download workbook after 3 attempts: {last_err}")


def find_matching_column_indices(headers: list[str]) -> dict:
    col_idx_map = {}
    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_upper = re.sub(r"\s+", "", upper_col)

        if (COL_SO in norm_col or "ಶಾಖಾಧಿಕಾರಿ" in norm_col) and "so" not in col_idx_map:
            col_idx_map["so"] = idx
        elif (
            COL_AEE in norm_col or "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್" in norm_col
        ) and "aee" not in col_idx_map:
            col_idx_map["aee"] = idx
        elif (
            COL_TARGET_DATE in norm_col or "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ" in norm_col
        ) and "target_date" not in col_idx_map:
            col_idx_map["target_date"] = idx
        elif (
            COL_ACTUAL_DATE in norm_col
            and "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ" not in norm_col
            and "actual_date" not in col_idx_map
        ):
            col_idx_map["actual_date"] = idx
        elif "CP/CF" in compact_upper and "status" not in col_idx_map:
            col_idx_map["status"] = idx

    return col_idx_map


def build_clean_unique_headers(raw_headers: list[str], col_idx_map: dict) -> list[str]:
    standardized = list(raw_headers)
    if "so" in col_idx_map:
        standardized[col_idx_map["so"]] = COL_SO
    if "aee" in col_idx_map:
        standardized[col_idx_map["aee"]] = COL_AEE
    if "target_date" in col_idx_map:
        standardized[col_idx_map["target_date"]] = COL_TARGET_DATE
    if "actual_date" in col_idx_map:
        standardized[col_idx_map["actual_date"]] = COL_ACTUAL_DATE
    if "status" in col_idx_map:
        standardized[col_idx_map["status"]] = COL_STATUS_STD

    seen = {}
    unique_headers = []
    for idx, h in enumerate(standardized):
        base = h if h else f"Unnamed_Col_{idx + 1}"
        if base not in seen:
            seen[base] = 1
            unique_headers.append(base)
        else:
            seen[base] += 1
            unique_headers.append(f"{base} ({seen[base]})")
    return unique_headers


def parse_indian_date(series: pd.Series) -> pd.Series:
    def _clean_date(val):
        if val is None or pd.isna(val):
            return pd.NaT
        if isinstance(val, (datetime, pd.Timestamp)):
            return pd.Timestamp(val.date())
        s = normalize_text(val)
        if not s or s.upper() in {"-", "--", "NIL", "NA", "N/A", "NONE", ""}:
            return pd.NaT
        s = re.sub(r"[./]", "-", s)
        return s

    cleaned = series.map(_clean_date)
    return pd.to_datetime(cleaned, dayfirst=True, errors="coerce")


def is_empty_date(series: pd.Series) -> pd.Series:
    def _check_empty(val):
        if val is None or pd.isna(val):
            return True
        s = normalize_text(val)
        return s.upper() in {"", "-", "--", "NIL", "NA", "N/A", "NONE", "NAT"}

    return series.map(_check_empty)


def is_valid_officer_name(name: str) -> bool:
    if not name or len(name) > 65:
        return False
    upper = name.upper()
    invalid_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in invalid_tokens or name.isdigit():
        return False
    if "ಒಟ್ಟು" in name or "TOTAL" in upper or COL_SO in name or COL_AEE in name:
        return False
    return True


def shorten_mobile_label(name: str, max_len: int = 16) -> str:
    """Shorten Y-axis officer names on portrait phones so bars have maximum width."""
    clean = re.sub(r"^(ಶ್ರೀ|ಶ್ರೀಮತಿ|ಕುಮಾರಿ)\.?\s*", "", name).strip()
    return clean if len(clean) <= max_len else clean[: max_len - 1] + "…"


def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Overdue_Works")
    return output.getvalue()


def parse_xlsx_workbook(workbook_bytes: bytes):
    progress = st.progress(0, text="Opening workbook...")

    try:
        excel_file = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="calamine")
    except Exception:
        excel_file = pd.ExcelFile(io.BytesIO(workbook_bytes), engine="openpyxl")

    available_sheets = {
        normalize_text(name).lower(): name for name in excel_file.sheet_names
    }

    today_ist = pd.Timestamp(datetime.now(ZoneInfo("Asia/Kolkata")).date())

    compiled_rows = []
    full_overdue_rows = []
    sheet_diagnostics = []

    total_targets = len(TARGET_SHEETS)
    for idx_sheet, target in enumerate(TARGET_SHEETS, start=1):
        progress.progress(
            idx_sheet / total_targets,
            text=f"Reading {idx_sheet}/{total_targets}: {target}...",
        )

        norm_target = normalize_text(target).lower()
        actual_sheet_name = available_sheets.get(norm_target)

        if not actual_sheet_name:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Missing",
                    "Rows": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )
            continue

        try:
            raw_df = pd.read_excel(excel_file, sheet_name=actual_sheet_name, header=None)
            if raw_df.empty:
                continue

            header_idx = None
            for idx in range(min(30, len(raw_df))):
                row_text = " | ".join(normalize_text(v) for v in raw_df.iloc[idx].values)
                if (
                    COL_SO in row_text
                    or COL_AEE in row_text
                    or "ಶಾಖಾಧಿಕಾರಿ" in row_text
                ):
                    header_idx = idx
                    break

            if header_idx is None:
                continue

            raw_headers = [normalize_text(col) for col in raw_df.iloc[header_idx].values]
            col_idx_map = find_matching_column_indices(raw_headers)

            required_keys = {"so", "aee", "target_date", "actual_date", "status"}
            if (required_keys - set(col_idx_map.keys())) and (header_idx + 1 < len(raw_df)):
                next_row = [normalize_text(col) for col in raw_df.iloc[header_idx + 1].values]
                combined_headers = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(raw_headers, next_row)
                ]
                alt_map = find_matching_column_indices(combined_headers)
                if len(alt_map) > len(col_idx_map):
                    raw_headers = combined_headers
                    col_idx_map = alt_map

            if required_keys - set(col_idx_map.keys()):
                continue

            df = raw_df.iloc[header_idx + 1 :].reset_index(drop=True).copy()

            s_so = df.iloc[:, col_idx_map["so"]].map(normalize_text)
            s_aee = df.iloc[:, col_idx_map["aee"]].map(normalize_text)
            s_status = df.iloc[:, col_idx_map["status"]].map(normalize_text)
            s_status_upper = s_status.str.upper().str.strip()
            s_target_raw = df.iloc[:, col_idx_map["target_date"]]
            s_actual_raw = df.iloc[:, col_idx_map["actual_date"]]

            target_dt = parse_indian_date(s_target_raw)
            actual_empty = is_empty_date(s_actual_raw)

            is_status_p = s_status_upper == "P"
            is_status_ets = s_status_upper == "ETS"

            is_overdue = (
                is_status_p
                & target_dt.notna()
                & (target_dt < today_ist)
                & actual_empty
            )

            valid_row_mask = (s_status != "") | (s_so != "") | (s_aee != "")

            sub = pd.DataFrame(
                {
                    "sheet_name": target,
                    "so_name": s_so,
                    "aee_name": s_aee,
                    "status_raw": s_status,
                    "target_date_raw": s_target_raw.map(format_cell_for_display),
                    "actual_date_raw": s_actual_raw.map(format_cell_for_display),
                    "target_dt": target_dt,
                    "actual_empty": actual_empty,
                    "is_status_p": is_status_p,
                    "is_status_ets": is_status_ets,
                    "is_overdue": is_overdue,
                }
            )
            sub = sub[valid_row_mask].reset_index(drop=True)
            compiled_rows.append(sub)

            overdue_mask = valid_row_mask & is_overdue
            if overdue_mask.any():
                unique_headers = build_clean_unique_headers(raw_headers, col_idx_map)
                full_df = df.loc[overdue_mask].copy()
                full_df.columns = unique_headers

                unnamed_cols = [
                    c for c in full_df.columns if c.startswith("Unnamed_Col_")
                ]
                for u_col in unnamed_cols:
                    if full_df[u_col].map(format_cell_for_display).eq("").all():
                        full_df.drop(columns=[u_col], inplace=True)

                for col in full_df.columns:
                    full_df[col] = full_df[col].map(format_cell_for_display)

                full_df.insert(0, "Worksheet / ಯೋಜನೆ", target)
                full_df["_so_name"] = s_so.loc[overdue_mask].values
                full_df["_aee_name"] = s_aee.loc[overdue_mask].values

                full_overdue_rows.append(full_df.reset_index(drop=True))

            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Loaded",
                    "Rows": len(sub),
                    "Overdue ('P')": int(overdue_mask.sum()),
                    "ETS": int((valid_row_mask & is_status_ets).sum()),
                }
            )

        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "Rows": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )

    progress.empty()

    master_df = (
        pd.concat(compiled_rows, ignore_index=True)
        if compiled_rows
        else pd.DataFrame()
    )
    overdue_full_df = (
        pd.concat(full_overdue_rows, ignore_index=True).fillna("")
        if full_overdue_rows
        else pd.DataFrame()
    )
    diag_df = pd.DataFrame(sheet_diagnostics)
    return master_df, overdue_full_df, diag_df, today_ist.strftime("%d-%m-%Y")


def save_shared_state(state: dict, data_tuple, source_label: str):
    now_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %I:%M %p")
    state["data"] = data_tuple
    state["last_synced"] = now_str
    state["source_label"] = source_label
    try:
        with open(DISK_BACKUP_FILE, "wb") as f:
            pickle.dump(
                {
                    "data": data_tuple,
                    "last_synced": now_str,
                    "source_label": source_label,
                },
                f,
            )
    except Exception:
        pass


# -----------------------------------------------------------------------------
# MOBILE HEADER & DATA LOAD
# -----------------------------------------------------------------------------
st.markdown(
    '<div class="mobile-title">📊 Minor Irrigation — Overdue Works</div>',
    unsafe_allow_html=True,
)

shared_state = get_shared_state()

if shared_state["data"] is None:
    with shared_state["lock"]:
        if shared_state["data"] is None:
            with st.spinner("Loading 23 worksheets from Google Sheets..."):
                try:
                    xlsx_bytes = download_xlsx_bytes(DEFAULT_SHEET_URL)
                    parsed = parse_xlsx_workbook(xlsx_bytes)
                    save_shared_state(shared_state, parsed, "Live Google Sheet")
                except Exception as e:
                    st.error(f"Unable to load data right now. Please tap Refresh. ({e})")
                    st.stop()

master_df, overdue_full_df, diag_df, today_str = shared_state["data"]

st.markdown(
    f'<div class="mobile-subtitle">📅 Today: <b>{today_str}</b> &nbsp;|&nbsp; '
    f'🔄 Synced: {shared_state["last_synced"]}</div>',
    unsafe_allow_html=True,
)

if master_df.empty:
    st.warning("No matching data rows found.")
    st.stop()

# -----------------------------------------------------------------------------
# TOUCH-FRIENDLY TOP CONTROLS FOR CHIEF ENGINEER
# -----------------------------------------------------------------------------
view_choice = st.radio(
    "👤 ವೀಕ್ಷಣೆ ಆಯ್ಕೆಮಾಡಿ / Select Officer Level:",
    options=[COL_SO, COL_AEE],
    horizontal=True,
)

c1, c2 = st.columns([1.2, 1])
with c1:
    sort_choice = st.radio(
        "⬇️ Sort Ranking By:",
        options=[
            "Highest Overdue Count",
            "Highest Overdue %",
        ],
        horizontal=True,
    )

with c2:
    filter_mode = st.selectbox(
        "🔍 Show Officers:",
        options=[
            "Overdue > 0 Only",
            "Top 20 Officers",
            "Top 35 Officers",
            "All Officers",
        ],
        index=0,
    )

group_col = "so_name" if view_choice == COL_SO else "aee_name"
full_filter_col = "_so_name" if view_choice == COL_SO else "_aee_name"

valid_df = master_df[master_df[group_col].map(is_valid_officer_name)].copy()

summary_df = (
    valid_df.groupby(group_col, as_index=False)
    .agg(
        total_works=("is_overdue", "count"),
        ongoing_p_works=("is_status_p", "sum"),
        ets_works=("is_status_ets", "sum"),
        overdue_works=("is_overdue", "sum"),
    )
    .rename(columns={group_col: view_choice})
)

summary_df["rest_works"] = (
    summary_df["total_works"] - summary_df["overdue_works"] - summary_df["ets_works"]
).clip(lower=0)

summary_df["overdue_pct"] = (
    (summary_df["overdue_works"] / summary_df["total_works"]) * 100
).round(1).fillna(0.0)

if sort_choice == "Highest Overdue %":
    summary_df = summary_df.sort_values(
        by=["overdue_pct", "overdue_works", "total_works"],
        ascending=[False, False, False],
    ).reset_index(drop=True)
else:
    summary_df = summary_df.sort_values(
        by=["overdue_works", "overdue_pct", "total_works"],
        ascending=[False, False, False],
    ).reset_index(drop=True)

# -----------------------------------------------------------------------------
# 2x2 MOBILE PORTRAIT KPI GRID
# -----------------------------------------------------------------------------
total_assigned_all = int(summary_df["total_works"].sum())
total_overdue_all = int(summary_df["overdue_works"].sum())
total_ets_all = int(summary_df["ets_works"].sum())
overall_overdue_pct = (
    round((total_overdue_all / total_assigned_all) * 100, 1)
    if total_assigned_all > 0
    else 0.0
)

st.markdown(
    f"""
    <div class="kpi-grid">
        <div class="kpi-card-red">
            <div class="kpi-label">🔴 Overdue Works ('P')</div>
            <div class="kpi-val">{total_overdue_all:,} ({overall_overdue_pct}%)</div>
        </div>
        <div class="kpi-card-green">
            <div class="kpi-label">🟢 Total 'ETS' Works</div>
            <div class="kpi-val">{total_ets_all:,}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">📋 Total Works Assigned</div>
            <div class="kpi-val">{total_assigned_all:,}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">👥 Total Officers</div>
            <div class="kpi-val">{len(summary_df):,}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# MOBILE PORTRAIT HORIZONTAL BAR CHART (RED = OVERDUE, GREEN = ETS, GREY = REST)
# -----------------------------------------------------------------------------
if filter_mode == "Overdue > 0 Only":
    chart_source_df = summary_df[summary_df["overdue_works"] > 0].copy()
elif filter_mode == "Top 20 Officers":
    chart_source_df = summary_df.head(20).copy()
elif filter_mode == "Top 35 Officers":
    chart_source_df = summary_df.head(35).copy()
else:
    chart_source_df = summary_df.copy()

st.markdown(
    """
    <div class="legend-bar">
        <span><span class="dot-red">■</span> Overdue ('P')</span>
        <span><span class="dot-green">■</span> ETS Works</span>
        <span><span class="dot-grey">■</span> Rest of Works</span>
    </div>
    """,
    unsafe_allow_html=True,
)

if chart_source_df.empty:
    st.info("No officers match the current filter.")
else:
    plot_df = chart_source_df.iloc[::-1].copy()
    plot_df["y_short"] = plot_df[view_choice].map(lambda s: shorten_mobile_label(s, 15))
    plot_df["overdue_plus_ets"] = plot_df["overdue_works"] + plot_df["ets_works"]

    # Concise label so it fits cleanly on narrow portrait screens
    if sort_choice == "Highest Overdue %":
        plot_df["bar_label"] = plot_df.apply(
            lambda r: (
                f" <b>{r['overdue_pct']:.0f}%</b> "
                f"({int(r['overdue_works'])} OD | {int(r['ets_works'])} ETS / {int(r['total_works'])})"
            ),
            axis=1,
        )
    else:
        plot_df["bar_label"] = plot_df.apply(
            lambda r: (
                f" <b>{int(r['overdue_works'])} OD</b> | "
                f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
                f"{int(r['total_works'])} ({r['overdue_pct']:.0f}%)"
            ),
            axis=1,
        )

    fig = go.Figure()

    # 1. Grey Bar: Total Works (visible tail = Rest of Works)
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["total_works"],
            name="Rest of Works",
            orientation="h",
            marker=dict(
                color="rgba(148, 163, 184, 0.28)",
                line=dict(color="#64748B", width=1),
            ),
            text=plot_df["bar_label"],
            textposition="outside",
            textfont=dict(size=11),
            cliponaxis=False,
            customdata=plot_df[
                [
                    view_choice,
                    "total_works",
                    "overdue_works",
                    "ets_works",
                    "rest_works",
                    "overdue_pct",
                ]
            ],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "🔴 Overdue ('P'): <b>%{customdata[2]}</b> (%{customdata[5]:.1f}%)<br>"
                "🟢 ETS Works: <b>%{customdata[3]}</b><br>"
                "⚪ Rest of Works: <b>%{customdata[4]}</b><br>"
                "📋 Total Works: <b>%{customdata[1]}</b><extra></extra>"
            ),
        )
    )

    # 2. Light Green Bar: Overdue + ETS (visible middle = ETS Works)
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["overdue_plus_ets"],
            name="ETS Works",
            orientation="h",
            marker=dict(color="#4ADE80"),
            customdata=plot_df[
                [
                    view_choice,
                    "total_works",
                    "overdue_works",
                    "ets_works",
                    "rest_works",
                    "overdue_pct",
                ]
            ],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "🟢 ETS Works: <b>%{customdata[3]}</b><br>"
                "🔴 Overdue ('P'): <b>%{customdata[2]}</b> (%{customdata[5]:.1f}%)<br>"
                "📋 Total Works: <b>%{customdata[1]}</b><extra></extra>"
            ),
        )
    )

    # 3. Red Bar: Overdue Works ('P')
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["overdue_works"],
            name="Overdue Works ('P')",
            orientation="h",
            marker=dict(color="#EF4444"),
            customdata=plot_df[
                [
                    view_choice,
                    "total_works",
                    "overdue_works",
                    "ets_works",
                    "rest_works",
                    "overdue_pct",
                ]
            ],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "🔴 Overdue ('P'): <b>%{customdata[2]}</b> (%{customdata[5]:.1f}%)<br>"
                "🟢 ETS Works: <b>%{customdata[3]}</b><br>"
                "📋 Total Works: <b>%{customdata[1]}</b><extra></extra>"
            ),
        )
    )

    max_x = max(int(plot_df["total_works"].max()), 1)
    x_range_max = max_x * 1.52
    chart_height = max(360, len(plot_df) * 32 + 50)

    fig.update_layout(
        barmode="overlay",
        height=chart_height,
        bargap=0.24,
        showlegend=False,
        xaxis=dict(
            title="",
            range=[0, x_range_max],
            gridcolor="rgba(200, 200, 200, 0.15)",
            zeroline=True,
            fixedrange=True,  # Prevents accidental zoom while scrolling on phone
            tickfont=dict(size=10),
        ),
        yaxis=dict(
            title="",
            automargin=False,
            tickmode="array",
            tickvals=plot_df[view_choice],
            ticktext=plot_df["y_short"],
            tickfont=dict(size=11),
            fixedrange=True,  # Prevents accidental pan while scrolling on phone
        ),
        margin=dict(l=118, r=10, t=10, b=25),
    )

    # Disable Plotly floating toolbar for clean mobile touch scrolling
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False, "scrollZoom": False},
    )

st.divider()

# -----------------------------------------------------------------------------
# OFFICER DRILL-DOWN (MOBILE CARDS + FULL SPREADSHEET TABLE + DOWNLOAD)
# -----------------------------------------------------------------------------
st.markdown("##### 📂 Officer Overdue Works Drill-Down & Download")

officer_options = summary_df[view_choice].tolist()
officer_stats_map = {
    row[view_choice]: (
        int(row["overdue_works"]),
        int(row["ets_works"]),
        int(row["total_works"]),
        float(row["overdue_pct"]),
    )
    for _, row in summary_df.iterrows()
}

selected_officer = st.selectbox(
    f"Tap to select {view_choice}:",
    options=officer_options,
    index=0 if officer_options else None,
    format_func=lambda name: (
        f"{name} ({officer_stats_map[name][0]} OD | "
        f"{officer_stats_map[name][1]} ETS / {officer_stats_map[name][2]} Total — "
        f"{officer_stats_map[name][3]:.1f}%)"
    ),
)

if selected_officer:
    if overdue_full_df.empty:
        st.info("No overdue works found across the 23 worksheets.")
    else:
        officer_overdue_df = overdue_full_df[
            overdue_full_df[full_filter_col] == selected_officer
        ].copy()

        internal_cols = [
            c for c in ["_so_name", "_aee_name"] if c in officer_overdue_df.columns
        ]
        officer_overdue_df.drop(columns=internal_cols, inplace=True)

        extra_unnamed = [
            c
            for c in officer_overdue_df.columns
            if c.startswith("Unnamed_Col_") and officer_overdue_df[c].eq("").all()
        ]
        if extra_unnamed:
            officer_overdue_df.drop(columns=extra_unnamed, inplace=True)

        officer_overdue_df = officer_overdue_df.reset_index(drop=True)
        officer_overdue_df.index = officer_overdue_df.index + 1
        officer_overdue_df.index.name = "Sl. No."

        if officer_overdue_df.empty:
            st.success(f"🎉 **{selected_officer}** has **0** overdue works!")
        else:
            sheets_involved = officer_overdue_df["Worksheet / ಯೋಜನೆ"].unique().tolist()
            st.caption(
                f"**Overdue Works:** `{len(officer_overdue_df)}` &nbsp;|&nbsp; "
                f"**Sheets:** `{', '.join(sheets_involved)}`"
            )

            safe_name = (
                re.sub(r"[^\w\-]+", "_", selected_officer).strip("_") or "Officer"
            )
            excel_bytes = convert_df_to_excel_bytes(officer_overdue_df.reset_index())
            csv_bytes = (
                officer_overdue_df.reset_index().to_csv(index=False).encode("utf-8-sig")
            )

            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    label="📥 Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Overdue_{safe_name}_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl2:
                st.download_button(
                    label="📥 CSV (.csv)",
                    data=csv_bytes,
                    file_name=f"Overdue_{safe_name}_{today_str}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            st.dataframe(
                officer_overdue_df,
                use_container_width=True,
                height=min(480, max(200, (len(officer_overdue_df) + 1) * 40)),
            )

# -----------------------------------------------------------------------------
# COLLAPSIBLE SUMMARY TABLE & ADMIN SYNC AT THE BOTTOM
# -----------------------------------------------------------------------------
with st.expander("📋 View Full Officer-Wise Summary Table"):
    display_table = summary_df[
        [
            view_choice,
            "overdue_works",
            "ets_works",
            "total_works",
            "overdue_pct",
        ]
    ].rename(
        columns={
            "overdue_works": "Overdue ('P')",
            "ets_works": "ETS",
            "total_works": "Total",
            "overdue_pct": "Overdue %",
        }
    )
    display_table.index = display_table.index + 1
    st.dataframe(
        display_table.style.format({"Overdue %": "{:.1f}%"}),
        use_container_width=True,
    )

with st.expander("⚙️ Sync / Refresh Live Data from Google Sheets"):
    sheet_url_input = st.text_input(
        "Google Spreadsheet Link",
        value=DEFAULT_SHEET_URL,
    )
    if st.button(
        "🔄 Sync Latest Data Now",
        use_container_width=True,
        type="primary",
    ):
        with shared_state["lock"]:
            with st.spinner("Syncing latest changes from Google Sheets..."):
                xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                parsed = parse_xlsx_workbook(xlsx_bytes)
                save_shared_state(shared_state, parsed, "Live Google Sheet")
                st.rerun()
    st.dataframe(diag_df, use_container_width=True)
