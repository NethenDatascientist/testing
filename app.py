import io
import os
import pickle
import re
import threading
import time
from datetime import date, datetime
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION (MOBILE PORTRAIT OPTIMIZED FOR CE)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="MI Overdue Works — CE Mobile View",
    page_icon="📱",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# Mobile-first Portrait CSS (unclipped top heading, spaced control cards, 2x2 KPI grid)
st.markdown(
    """
    <style>
        header[data-testid="stHeader"] {
            background: transparent !important;
        }
        .block-container {
            max-width: 680px !important;
            padding-left: 0.65rem !important;
            padding-right: 0.65rem !important;
            padding-top: 3.8rem !important;
            padding-bottom: 2.5rem !important;
        }
        /* Compact Title for Mobile */
        .mobile-title {
            font-size: 1.25rem;
            font-weight: 800;
            line-height: 1.35;
            margin-top: 0.2rem;
            margin-bottom: 0.2rem;
        }
        .mobile-subtitle {
            font-size: 0.76rem;
            opacity: 0.78;
            margin-bottom: 0.75rem;
        }
        /* Spaced bordered control cards on mobile */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 10px !important;
            padding: 0.3rem 0.45rem !important;
            margin-bottom: 0.45rem !important;
            background-color: rgba(148, 163, 184, 0.05);
        }
        /* 2x2 Responsive KPI Grid for Portrait Phones */
        .kpi-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            margin-top: 8px;
            margin-bottom: 8px;
        }
        .kpi-card {
            background: rgba(148, 163, 184, 0.10);
            border: 1px solid rgba(148, 163, 184, 0.25);
            border-radius: 10px;
            padding: 9px 11px;
        }
        .kpi-card-red {
            background: rgba(239, 68, 68, 0.14);
            border: 1px solid rgba(239, 68, 68, 0.45);
            border-radius: 10px;
            padding: 9px 11px;
        }
        .kpi-card-green {
            background: rgba(74, 222, 128, 0.12);
            border: 1px solid rgba(74, 222, 128, 0.40);
            border-radius: 10px;
            padding: 9px 11px;
        }
        .kpi-label {
            font-size: 0.73rem;
            font-weight: 600;
            opacity: 0.85;
            margin-bottom: 2px;
        }
        .kpi-val {
            font-size: 1.18rem;
            font-weight: 800;
            line-height: 1.2;
        }
        /* Legend pill bar */
        .legend-bar {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            font-size: 0.76rem;
            font-weight: 600;
            margin: 4px 0 4px 0;
            padding: 6px 10px;
            background: rgba(148, 163, 184, 0.08);
            border-radius: 8px;
        }
        .dot-red { color: #EF4444; }
        .dot-green { color: #4ADE80; }
        .dot-grey { color: #94A3B8; }
        div[role="radiogroup"] {
            gap: 0.45rem !important;
        }
        hr {
            margin-top: 0.35rem !important;
            margin-bottom: 0.45rem !important;
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

COL_DIV = "ವಿಭಾಗ"
COL_SUBDIV = "ಉಪವಿಭಾಗ"
COL_SO = "ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು"
COL_AEE = "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು"
COL_TARGET_DATE = "ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ"
COL_ACTUAL_DATE = "ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ"
COL_STATUS_STD = "CP/CF/P/TBS/TP/ETS/OTH"

ALL_DIVISIONS_OPTION = "ALL (ಎಲ್ಲಾ ವಿಭಾಗಗಳು)"
ALL_SUBDIVISIONS_OPTION = "ALL (ಎಲ್ಲಾ ಉಪವಿಭಾಗಗಳು)"
UNSPECIFIED_SUBDIV_LABEL = "ಉಪವಿಭಾಗ ನಮೂದಿಸಿಲ್ಲ / Not Specified (Empty)"
UNASSIGNED_LABEL = "ಹೆಸರು ನಮೂದಿಸಿಲ್ಲ / Unassigned (Empty)"

# Shrink bars above 250 works so normal bars stay wide and full labels never clip on phone
BAR_SHRINK_THRESHOLD = 250
BAR_SHRINK_MAX_EXTRA = 22.0

DEFAULT_SHEET_URL = st.secrets.get(
    "SHEET_URL",
    "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869",
)

SCHEMA_VERSION = "v9_upavibhaga_logic"
DISK_BACKUP_FILE = "/tmp/mi_overdue_shared_snapshot_v9.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK (Instant load on CE's phone)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_shared_state_v9() -> dict:
    state = {
        "version": SCHEMA_VERSION,
        "lock": threading.Lock(),
        "data": None,
        "last_synced": None,
        "source_label": None,
    }
    if os.path.exists(DISK_BACKUP_FILE):
        try:
            with open(DISK_BACKUP_FILE, "rb") as f:
                saved = pickle.load(f)
                if saved.get("version") == SCHEMA_VERSION:
                    state["data"] = saved.get("data")
                    state["last_synced"] = saved.get("last_synced")
                    state["source_label"] = saved.get("source_label")
        except Exception:
            pass
    return state


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def compress_bar_length(val: float, threshold: int = BAR_SHRINK_THRESHOLD) -> float:
    """
    Keep values <= 250 at full 1:1 scale.
    For values > 250, smoothly compress the excess into a tight 250..272 band
    so 250+ bars still look longer than other bars while leaving room for text labels.
    """
    v = float(val)
    t = float(threshold)
    if v <= t:
        return v
    excess = v - t
    return t + BAR_SHRINK_MAX_EXTRA * (excess / (excess + 120.0))


def to_safe_python_date(val):
    """
    Convert any datetime/Timestamp/date object into a plain Python datetime.date.
    Supports years 0001 to 9999 and avoids Pandas nanosecond overflow & timezone errors.
    """
    if val is None:
        return None
    if isinstance(val, pd.Timestamp):
        return date(val.year, val.month, val.day)
    if isinstance(val, datetime):
        return date(val.year, val.month, val.day)
    if isinstance(val, date):
        return val
    return val


def sanitize_dataframe_cells(df: pd.DataFrame) -> pd.DataFrame:
    cleaned_rows = [
        [to_safe_python_date(cell) for cell in row]
        for row in df.to_numpy().tolist()
    ]
    return pd.DataFrame(cleaned_rows, dtype=object)


def read_sheet_safely(
    excel_file: pd.ExcelFile, workbook_bytes: bytes, sheet_name: str
) -> pd.DataFrame:
    try:
        raw_df = pd.read_excel(
            excel_file, sheet_name=sheet_name, header=None, dtype=object
        )
        return sanitize_dataframe_cells(raw_df)
    except Exception:
        wb = openpyxl.load_workbook(
            io.BytesIO(workbook_bytes), data_only=True, read_only=True
        )
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            cleaned_row = [to_safe_python_date(cell) for cell in row]
            if any(c is not None and str(c).strip() != "" for c in cleaned_row):
                rows.append(cleaned_row)
            elif len(rows) < 35:
                rows.append(cleaned_row)
        wb.close()
        return pd.DataFrame(rows, dtype=object)


def normalize_text(val) -> str:
    val = to_safe_python_date(val)
    if val is None:
        return ""
    if isinstance(val, date):
        return f"{val.day:02d}-{val.month:02d}-{val.year:04d}"
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(val)).strip()


def format_cell_for_display(val) -> str:
    val = to_safe_python_date(val)
    if val is None:
        return ""
    if isinstance(val, date):
        return f"{val.day:02d}-{val.month:02d}-{val.year:04d}"
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
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
        compact_col = re.sub(r"[\s\-_]+", "", norm_col)
        if norm_col in {"ವಿಭಾಗ", "ವಿಭಾಗದ ಹೆಸರು"} and "division" not in col_idx_map:
            col_idx_map["division"] = idx
        elif (
            compact_col in {"ಉಪವಿಭಾಗ", "ಉಪವಿಭಾಗದಹೆಸರು"}
            and "subdivision" not in col_idx_map
        ):
            col_idx_map["subdivision"] = idx

    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_upper = re.sub(r"\s+", "", upper_col)

        if (
            ("ಉಪವಿಭಾಗ" in norm_col or "ಉಪ ವಿಭಾಗ" in norm_col or "ಉಪ-ವಿಭಾಗ" in norm_col)
            and "subdivision" not in col_idx_map
        ):
            col_idx_map["subdivision"] = idx
        elif (
            "ವಿಭಾಗ" in norm_col
            and "ಉಪ" not in norm_col
            and "SUB" not in upper_col
            and "division" not in col_idx_map
        ):
            col_idx_map["division"] = idx
        elif (COL_SO in norm_col or "ಶಾಖಾಧಿಕಾರಿ" in norm_col) and "so" not in col_idx_map:
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
    if "division" in col_idx_map:
        standardized[col_idx_map["division"]] = COL_DIV
    if "subdivision" in col_idx_map:
        standardized[col_idx_map["subdivision"]] = COL_SUBDIV
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


def parse_indian_date_list(series: pd.Series) -> list:
    parsed_dates = []
    for raw_val in series.tolist():
        val = to_safe_python_date(raw_val)
        if val is None:
            parsed_dates.append(None)
            continue

        if isinstance(val, date):
            parsed_dates.append(val)
            continue

        s = normalize_text(val)
        if not s or s.upper() in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", ""}:
            parsed_dates.append(None)
            continue

        s = re.sub(r"[./]", "-", s).strip(" -")
        match = re.search(r"(\d{1,4})-(\d{1,2})-(\d{1,4})", s)
        if match:
            p1, p2, p3 = int(match.group(1)), int(match.group(2)), int(match.group(3))
            try:
                if p1 > 31:
                    parsed_dates.append(date(max(1, min(9999, p1)), p2, p3))
                else:
                    yr = p3 + 2000 if p3 < 100 else p3
                    parsed_dates.append(date(max(1, min(9999, yr)), p2, p1))
                continue
            except Exception:
                pass

        try:
            dt = pd.to_datetime(s, dayfirst=True, utc=True, errors="coerce")
            if pd.isna(dt):
                parsed_dates.append(None)
            else:
                parsed_dates.append(date(dt.year, dt.month, dt.day))
        except Exception:
            parsed_dates.append(None)

    return parsed_dates


def is_empty_date(series: pd.Series) -> pd.Series:
    results = []
    for raw_val in series.tolist():
        val = to_safe_python_date(raw_val)
        if val is None:
            results.append(True)
        elif isinstance(val, date):
            results.append(False)
        else:
            s = normalize_text(val)
            results.append(
                s.upper() in {"", "-", "--", "---", "NIL", "NA", "N/A", "NONE", "NAT"}
            )
    return pd.Series(results, index=series.index, dtype=bool)


def is_valid_division_text(val: str) -> bool:
    if not val:
        return False
    upper = val.upper().strip()
    if upper in {"-", "--", "---", "NIL", "NA", "N/A", "NONE", "NAN", "NAT", "NULL", "0"}:
        return False
    if re.fullmatch(r"[\d\W_]+", val):
        return False
    if val in {COL_DIV, "ವಿಭಾಗದ ಹೆಸರು", "ಒಟ್ಟು", "ಒಟ್ಟು ಮೊತ್ತ", "ಮಹಾ ಒಟ್ಟು"}:
        return False
    if upper in {"DIVISION", "DIVISION NAME", "TOTAL", "GRAND TOTAL", "SUB TOTAL"}:
        return False
    if val.startswith("ಒಟ್ಟು ") or val.endswith(" ಒಟ್ಟು") or upper.startswith("TOTAL "):
        return False
    return True


def map_subdivision_or_unspecified(name: str) -> str:
    if not name:
        return UNSPECIFIED_SUBDIV_LABEL
    upper = name.upper().strip()
    empty_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in empty_tokens or re.fullmatch(r"[\d\W_]+", name):
        return UNSPECIFIED_SUBDIV_LABEL
    if (
        len(name) > 65
        or "ಒಟ್ಟು" in name
        or "TOTAL" in upper
        or name in {COL_SUBDIV, "ಉಪ ವಿಭಾಗ", "ಉಪವಿಭಾಗದ ಹೆಸರು"}
    ):
        return UNSPECIFIED_SUBDIV_LABEL
    return name


def map_officer_or_unassigned(name: str) -> str:
    if not name:
        return UNASSIGNED_LABEL
    upper = name.upper()
    empty_tokens = {
        "NAN", "NONE", "-", "--", "---", "NIL", "NA", "N/A", "0", "ಇಲ್ಲ", "ಖಾಲಿ"
    }
    if upper in empty_tokens or name.isdigit():
        return UNASSIGNED_LABEL
    if (
        len(name) > 65
        or "ಒಟ್ಟು" in name
        or "TOTAL" in upper
        or COL_SO in name
        or COL_AEE in name
    ):
        return UNASSIGNED_LABEL
    return name


def shorten_mobile_label(name: str, max_len: int = 15) -> str:
    """Shorten Y-axis officer names on portrait phones so bars have maximum width."""
    if name == UNASSIGNED_LABEL:
        return "⚠️ ಹೆಸರು ಖಾಲಿ"
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

    today_ist_date = datetime.now(ZoneInfo("Asia/Kolkata")).date()

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
                    "ಉಪವಿಭಾಗ": "N/A",
                    "Rows (ವಿಭಾಗ)": 0,
                    "Overdue ('P')": 0,
                    "ETS": 0,
                }
            )
            continue

        try:
            raw_df = read_sheet_safely(excel_file, workbook_bytes, actual_sheet_name)
            if raw_df.empty:
                continue

            header_idx = None
            for idx in range(min(30, len(raw_df))):
                row_text = " | ".join(
                    normalize_text(v) for v in raw_df.iloc[idx].tolist()
                )
                if (
                    COL_SO in row_text
                    or COL_AEE in row_text
                    or "ಶಾಖಾಧಿಕಾರಿ" in row_text
                ):
                    header_idx = idx
                    break

            if header_idx is None:
                continue

            raw_headers = [
                normalize_text(col) for col in raw_df.iloc[header_idx].tolist()
            ]
            col_idx_map = find_matching_column_indices(raw_headers)

            required_keys = {
                "division",
                "so",
                "aee",
                "target_date",
                "actual_date",
                "status",
            }
            if (
                (required_keys - set(col_idx_map.keys()))
                or ("subdivision" not in col_idx_map)
            ) and (header_idx + 1 < len(raw_df)):
                next_row = [
                    normalize_text(col) for col in raw_df.iloc[header_idx + 1].tolist()
                ]
                combined_headers = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(raw_headers, next_row)
                ]
                alt_map = find_matching_column_indices(combined_headers)
                if len(alt_map) > len(col_idx_map):
                    raw_headers = combined_headers
                    col_idx_map = alt_map

            if (
                ("division" not in col_idx_map or "subdivision" not in col_idx_map)
                and (header_idx > 0)
            ):
                prev_row = [
                    normalize_text(col) for col in raw_df.iloc[header_idx - 1].tolist()
                ]
                combined_prev = [
                    f"{h1} {h2}".strip() for h1, h2 in zip(prev_row, raw_headers)
                ]
                alt_prev_map = find_matching_column_indices(combined_prev)
                if "division" not in col_idx_map and "division" in alt_prev_map:
                    col_idx_map["division"] = alt_prev_map["division"]
                    raw_headers[col_idx_map["division"]] = COL_DIV
                if "subdivision" not in col_idx_map and "subdivision" in alt_prev_map:
                    col_idx_map["subdivision"] = alt_prev_map["subdivision"]
                    raw_headers[col_idx_map["subdivision"]] = COL_SUBDIV

            if required_keys - set(col_idx_map.keys()):
                continue

            df = raw_df.iloc[header_idx + 1 :].reset_index(drop=True).copy()

            s_div_raw = pd.Series(
                [
                    normalize_text(v)
                    for v in df.iloc[:, col_idx_map["division"]].tolist()
                ],
                index=df.index,
                dtype=str,
            )

            has_subdiv_col = "subdivision" in col_idx_map
            if has_subdiv_col:
                s_subdiv_raw = pd.Series(
                    [
                        normalize_text(v)
                        for v in df.iloc[:, col_idx_map["subdivision"]].tolist()
                    ],
                    index=df.index,
                    dtype=str,
                )
            else:
                s_subdiv_raw = pd.Series([""] * len(df), index=df.index, dtype=str)

            s_subdiv_mapped = s_subdiv_raw.map(map_subdivision_or_unspecified)

            s_so_raw = pd.Series(
                [normalize_text(v) for v in df.iloc[:, col_idx_map["so"]].tolist()],
                index=df.index,
                dtype=str,
            )
            s_aee_raw = pd.Series(
                [normalize_text(v) for v in df.iloc[:, col_idx_map["aee"]].tolist()],
                index=df.index,
                dtype=str,
            )
            s_status = pd.Series(
                [normalize_text(v) for v in df.iloc[:, col_idx_map["status"]].tolist()],
                index=df.index,
                dtype=str,
            )
            s_status_upper = s_status.str.upper().str.strip()
            s_target_raw = df.iloc[:, col_idx_map["target_date"]]
            s_actual_raw = df.iloc[:, col_idx_map["actual_date"]]

            target_date_list = parse_indian_date_list(s_target_raw)
            actual_empty = is_empty_date(s_actual_raw)

            is_status_p = s_status_upper == "P"
            is_status_ets = s_status_upper == "ETS"

            target_is_before_today = pd.Series(
                [
                    (d is not None) and (d < today_ist_date)
                    for d in target_date_list
                ],
                index=df.index,
                dtype=bool,
            )

            is_overdue = is_status_p & target_is_before_today & actual_empty

            valid_row_mask = s_div_raw.map(is_valid_division_text)

            s_so_mapped = s_so_raw.map(map_officer_or_unassigned)
            s_aee_mapped = s_aee_raw.map(map_officer_or_unassigned)

            sub = pd.DataFrame(
                {
                    "sheet_name": target,
                    "division_name": s_div_raw,
                    "subdivision_name": s_subdiv_mapped,
                    "so_name": s_so_mapped,
                    "aee_name": s_aee_mapped,
                    "status_raw": s_status,
                    "target_date_raw": [
                        format_cell_for_display(v) for v in s_target_raw.tolist()
                    ],
                    "actual_date_raw": [
                        format_cell_for_display(v) for v in s_actual_raw.tolist()
                    ],
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

                for col in full_df.columns:
                    full_df[col] = [
                        format_cell_for_display(v) for v in full_df[col].tolist()
                    ]

                unnamed_cols = [
                    c for c in full_df.columns if c.startswith("Unnamed_Col_")
                ]
                for u_col in unnamed_cols:
                    if (full_df[u_col] == "").all():
                        full_df.drop(columns=[u_col], inplace=True)

                full_df.insert(0, "Worksheet / ಯೋಜನೆ", target)
                full_df["_division_name"] = s_div_raw.loc[overdue_mask].values
                full_df["_subdivision_name"] = s_subdiv_mapped.loc[overdue_mask].values
                full_df["_so_name"] = s_so_mapped.loc[overdue_mask].values
                full_df["_aee_name"] = s_aee_mapped.loc[overdue_mask].values

                full_overdue_rows.append(full_df.reset_index(drop=True))

            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Loaded",
                    "ಉಪವಿಭಾಗ": "Found" if has_subdiv_col else "Missing (Included)",
                    "Rows (ವಿಭಾಗ)": len(sub),
                    "Overdue ('P')": int(overdue_mask.sum()),
                    "ETS": int((valid_row_mask & is_status_ets).sum()),
                }
            )

        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "ಉಪವಿಭಾಗ": "Error",
                    "Rows (ವಿಭಾಗ)": 0,
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
    return master_df, overdue_full_df, diag_df, today_ist_date.strftime("%d-%m-%Y")


def save_shared_state(state: dict, data_tuple, source_label: str):
    now_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %I:%M %p")
    state["version"] = SCHEMA_VERSION
    state["data"] = data_tuple
    state["last_synced"] = now_str
    state["source_label"] = source_label
    try:
        with open(DISK_BACKUP_FILE, "wb") as f:
            pickle.dump(
                {
                    "version": SCHEMA_VERSION,
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

shared_state = get_shared_state_v9()

if (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
):
    with shared_state["lock"]:
        if (
            shared_state["data"] is None
            or shared_state.get("version") != SCHEMA_VERSION
        ):
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
# TOUCH-FRIENDLY TOP CONTROLS IN SPACED CARDS FOR CHIEF ENGINEER
# -----------------------------------------------------------------------------
unique_divisions = sorted(
    [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
)
division_options = [ALL_DIVISIONS_OPTION] + unique_divisions

# Card 1: ವಿಭಾಗ & Cascading ಉಪವಿಭಾಗ Dropdowns
with st.container(border=True):
    selected_division = st.selectbox(
        f"🏛️ {COL_DIV} ಆಯ್ಕೆಮಾಡಿ / Select {COL_DIV}:",
        options=division_options,
        index=0,
    )

    if selected_division != ALL_DIVISIONS_OPTION:
        div_filtered_df = master_df[
            master_df["division_name"] == selected_division
        ].copy()
    else:
        div_filtered_df = master_df.copy()

    raw_subdivs = [
        s
        for s in div_filtered_df["subdivision_name"].dropna().unique().tolist()
        if s != "" and s != UNSPECIFIED_SUBDIV_LABEL
    ]
    named_subdivs = sorted(raw_subdivs)
    has_empty_subdivs = (
        div_filtered_df["subdivision_name"] == UNSPECIFIED_SUBDIV_LABEL
    ).any()

    subdivision_options = [ALL_SUBDIVISIONS_OPTION] + named_subdivs
    if has_empty_subdivs:
        subdivision_options.append(UNSPECIFIED_SUBDIV_LABEL)

    selected_subdivision = st.selectbox(
        f"🏢 {COL_SUBDIV} ಆಯ್ಕೆಮಾಡಿ / Select {COL_SUBDIV}:",
        options=subdivision_options,
        index=0,
    )

# Card 2: Officer Level & Sort Order
with st.container(border=True):
    view_choice = st.radio(
        "👤 ವೀಕ್ಷಣೆ ಆಯ್ಕೆಮಾಡಿ / Select Officer Level:",
        options=[COL_SO, COL_AEE],
        horizontal=True,
    )

    c1, c2 = st.columns([1.15, 1])
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
                "All Officers (incl. Empty)",
                "Overdue > 0 Only",
                "Top 20 Officers",
                "Top 35 Officers",
            ],
            index=0,
        )

group_col = "so_name" if view_choice == COL_SO else "aee_name"
full_filter_col = "_so_name" if view_choice == COL_SO else "_aee_name"

# -----------------------------------------------------------------------------
# APPLY ವಿಭಾಗ & ಉಪವಿಭಾಗ FILTERS TO ALL VALUES
# -----------------------------------------------------------------------------
if selected_subdivision != ALL_SUBDIVISIONS_OPTION:
    subdiv_filtered_df = div_filtered_df[
        div_filtered_df["subdivision_name"] == selected_subdivision
    ].copy()
else:
    subdiv_filtered_df = div_filtered_df.copy()

valid_df = subdiv_filtered_df[subdiv_filtered_df[group_col] != ""].copy()

if valid_df.empty:
    st.warning("No works found for the selected filter.")
    st.stop()

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

# Pin Unassigned bucket at the bottom
summary_df["_is_unassigned"] = summary_df[view_choice] == UNASSIGNED_LABEL

if sort_choice == "Highest Overdue %":
    summary_df = summary_df.sort_values(
        by=["_is_unassigned", "overdue_pct", "overdue_works", "total_works"],
        ascending=[True, False, False, False],
    ).reset_index(drop=True)
else:
    summary_df = summary_df.sort_values(
        by=["_is_unassigned", "overdue_works", "overdue_pct", "total_works"],
        ascending=[True, False, False, False],
    ).reset_index(drop=True)

# -----------------------------------------------------------------------------
# 2x2 MOBILE PORTRAIT KPI GRID (NO ವಿಭಾಗ CARD, TIGHT VERTICAL SPACING)
# -----------------------------------------------------------------------------
total_assigned_all = int(summary_df["total_works"].sum())
total_overdue_all = int(summary_df["overdue_works"].sum())
total_ets_all = int(summary_df["ets_works"].sum())
named_officers_count = int((~summary_df["_is_unassigned"]).sum())
overall_overdue_pct = (
    round((total_overdue_all / total_assigned_all) * 100, 1)
    if total_assigned_all > 0
    else 0.0
)

div_badge = (
    "ಎಲ್ಲಾ ವಿಭಾಗಗಳು (ALL)"
    if selected_division == ALL_DIVISIONS_OPTION
    else selected_division
)
subdiv_badge = (
    "ಎಲ್ಲಾ ಉಪವಿಭಾಗಗಳು (ALL)"
    if selected_subdivision == ALL_SUBDIVISIONS_OPTION
    else selected_subdivision
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
            <div class="kpi-label">📋 Total Works</div>
            <div class="kpi-val">{total_assigned_all:,}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">👥 Officers</div>
            <div class="kpi-val">{named_officers_count:,}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# MOBILE PORTRAIT HORIZONTAL BAR CHART (SHRUNK >250, ALL VALUES 100% VISIBLE)
# -----------------------------------------------------------------------------
if filter_mode == "Overdue > 0 Only":
    chart_source_df = summary_df[
        (summary_df["overdue_works"] > 0) | summary_df["_is_unassigned"]
    ].copy()
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
    plot_df["y_short"] = plot_df[view_choice].map(lambda s: shorten_mobile_label(s, 14))
    plot_df["overdue_plus_ets"] = plot_df["overdue_works"] + plot_df["ets_works"]

    # Compress bars > 250 so normal bars stretch wide across the phone screen
    plot_df["vis_total"] = plot_df["total_works"].map(compress_bar_length)
    plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].map(compress_bar_length)
    plot_df["vis_overdue"] = plot_df["overdue_works"].map(compress_bar_length)

    if sort_choice == "Highest Overdue %":
        plot_df["bar_label"] = plot_df.apply(
            lambda r: (
                f" <b>{r['overdue_pct']:.1f}%</b> "
                f"({int(r['overdue_works'])} OD | "
                f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
                f"{int(r['total_works'])})"
            ),
            axis=1,
        )
    else:
        plot_df["bar_label"] = plot_df.apply(
            lambda r: (
                f" <b>{int(r['overdue_works'])} OD</b> | "
                f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
                f"{int(r['total_works'])} ({r['overdue_pct']:.1f}%)"
            ),
            axis=1,
        )

    fig = go.Figure()

    # 1. Grey Bar: Total Works (visible tail = Rest of Works)
    fig.add_trace(
        go.Bar(
            y=plot_df[view_choice],
            x=plot_df["vis_total"],
            name="Rest of Works",
            orientation="h",
            marker=dict(
                color="rgba(148, 163, 184, 0.28)",
                line=dict(color="#64748B", width=1),
            ),
            text=plot_df["bar_label"],
            textposition="outside",
            textfont=dict(size=10.5),
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
                f"🏛️ {COL_DIV}: {div_badge}<br>"
                f"🏢 {COL_SUBDIV}: {subdiv_badge}<br>"
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
            x=plot_df["vis_overdue_plus_ets"],
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
            x=plot_df["vis_overdue"],
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

    # Ample right-side range so full text labels on 500+ bars never clip on portrait phones
    max_x = max(float(plot_df["vis_total"].max()), 1.0)
    x_range_max = max_x * 1.65
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
            fixedrange=True,  # Prevents accidental zoom while thumb-scrolling
            showticklabels=False,
        ),
        yaxis=dict(
            title="",
            automargin=False,
            tickmode="array",
            tickvals=plot_df[view_choice],
            ticktext=plot_df["y_short"],
            tickfont=dict(size=11),
            fixedrange=True,  # Prevents accidental pan while thumb-scrolling
        ),
        margin=dict(l=112, r=18, t=8, b=20),
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False, "scrollZoom": False},
    )

st.divider()

# -----------------------------------------------------------------------------
# OFFICER DRILL-DOWN (FILTERED BY ವಿಭಾಗ & ಉಪವಿಭಾಗ + DOWNLOAD)
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

        if selected_division != ALL_DIVISIONS_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_division_name"] == selected_division
            ].copy()

        if selected_subdivision != ALL_SUBDIVISIONS_OPTION:
            officer_overdue_df = officer_overdue_df[
                officer_overdue_df["_subdivision_name"] == selected_subdivision
            ].copy()

        internal_cols = [
            c
            for c in ["_division_name", "_subdivision_name", "_so_name", "_aee_name"]
            if c in officer_overdue_df.columns
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
            st.success(f"🎉 **{selected_officer}** has **0** overdue works in this selection!")
        else:
            sheets_involved = officer_overdue_df["Worksheet / ಯೋಜನೆ"].unique().tolist()
            st.caption(
                f"**{COL_DIV}:** `{div_badge}` &nbsp;|&nbsp; "
                f"**{COL_SUBDIV}:** `{subdiv_badge}`  \n"
                f"**Overdue Works:** `{len(officer_overdue_df)}` &nbsp;|&nbsp; "
                f"**Sheets:** `{', '.join(sheets_involved)}`"
            )

            safe_name = (
                re.sub(r"[^\w\-]+", "_", selected_officer).strip("_") or "Unassigned"
            )
            safe_div = (
                re.sub(r"[^\w\-]+", "_", div_badge).strip("_") or "All_Div"
            )
            safe_subdiv = (
                re.sub(r"[^\w\-]+", "_", subdiv_badge).strip("_") or "All_SubDiv"
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
                    file_name=f"Overdue_{safe_div}_{safe_subdiv}_{safe_name}_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl2:
                st.download_button(
                    label="📥 CSV (.csv)",
                    data=csv_bytes,
                    file_name=f"Overdue_{safe_div}_{safe_subdiv}_{safe_name}_{today_str}.csv",
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
