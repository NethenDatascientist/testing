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
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Overdue Works Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .block-container {
            max-width: 100% !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
            padding-top: 2.8rem !important;
            padding-bottom: 2rem !important;
        }
        [data-testid="stPlotlyChart"] {
            width: 100% !important;
        }
        /* Extra breathing room inside the top control cards */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 10px !important;
            padding: 0.35rem 0.5rem !important;
            background-color: rgba(148, 163, 184, 0.05);
        }
        div[role="radiogroup"] {
            gap: 0.55rem !important;
            padding-top: 0.25rem !important;
        }
        /* Tighten the vertical gap around dividers and below KPI metrics */
        hr {
            margin-top: 0.35rem !important;
            margin-bottom: 0.45rem !important;
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

# Shrink bars above 250 works so normal bars stay wide and full labels never clip
BAR_SHRINK_THRESHOLD = 250
BAR_SHRINK_MAX_EXTRA = 24.0

DEFAULT_SHEET_URL = st.secrets.get(
    "SHEET_URL",
    "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869",
)

SCHEMA_VERSION = "v9_upavibhaga_logic"
DISK_BACKUP_FILE = "/tmp/mi_overdue_shared_snapshot_v9.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK
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
    For values > 250, smoothly compress the excess into a tight 250..274 band
    so 250+ bars still look longer than other bars while leaving full room for text labels.
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
    """Convert all datetime objects to plain Python date objects across the DataFrame."""
    cleaned_rows = [
        [to_safe_python_date(cell) for cell in row]
        for row in df.to_numpy().tolist()
    ]
    return pd.DataFrame(cleaned_rows, dtype=object)


def read_sheet_safely(
    excel_file: pd.ExcelFile, workbook_bytes: bytes, sheet_name: str
) -> pd.DataFrame:
    """
    Read a worksheet safely. Falls back to direct openpyxl reading if pd.read_excel
    encounters out-of-bounds dates or mixed timezones.
    """
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
    """Collapse multiple spaces/newlines and strip text."""
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
    """Format raw Excel cell values cleanly for table display and export."""
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
    """Extract the Google Spreadsheet ID from the URL."""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("Invalid Google Sheet URL. Could not extract spreadsheet key.")
    return match.group(1)


def download_xlsx_bytes(sheet_url: str) -> bytes:
    """Download the Google Sheet as .xlsx with retry logic and browser headers."""
    key = extract_spreadsheet_key(sheet_url)
    urls = [
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx",
        f"https://docs.google.com/spreadsheets/d/{key}/export?format=xlsx&id={key}",
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/128.0.0.0 Safari/537.36"
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
    """
    Find required and optional 0-based column positions across sheets:
    - ವಿಭಾಗ (Division, excluding ಉಪವಿಭಾಗ / ಉಪ ವಿಭಾಗ)
    - ಉಪವಿಭಾಗ (Sub-Division, optional — missing in 2 sheets)
    - ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು
    - ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು
    - ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ
    - ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ
    - Status column containing 'CP/CF'
    """
    col_idx_map = {}

    # Pass 1: Exact match for ವಿಭಾಗ and ಉಪವಿಭಾಗ
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

    # Pass 2: Flexible substring matching
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
    """Standardize key column names across sheets and deduplicate repeated headers."""
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
    """Parse dates into a list of plain Python `datetime.date` objects (or `None`)."""
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
    """Check if the 'ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ' cell is empty or placeholder text."""
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
    """
    Row inclusion rule:
    Consider a row if and only if it has a valid text value in the 'ವಿಭಾಗ' column.
    """
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
    """
    Map blank, missing, or placeholder 'ಉಪವಿಭಾಗ' cells to UNSPECIFIED_SUBDIV_LABEL
    so rows with empty 'ಉಪವಿಭಾಗ' (or from the 2 sheets without 'ಉಪವಿಭಾಗ') are always
    considered and never dropped.
    """
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
    """Map blank or placeholder officer cells to UNASSIGNED_LABEL."""
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


def shorten_label(name: str, max_len: int = 28) -> str:
    """Keep Y-axis labels compact so the chart always starts on the far left."""
    if name == UNASSIGNED_LABEL:
        return "⚠️ ಹೆಸರು ನಮೂದಿಸಿಲ್ಲ (Empty)"
    return name if len(name) <= max_len else name[: max_len - 1] + "…"


def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Export a DataFrame to an in-memory .xlsx file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Overdue_Works")
    return output.getvalue()


def parse_xlsx_workbook(workbook_bytes: bytes):
    """
    Parse all 23 target worksheets.
    A row is counted if and only if the 'ವಿಭಾಗ' column contains a valid text value.
    'ಉಪವಿಭಾಗ' is optional (gracefully handled if the column is missing or cells are empty).
    """
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
            text=f"Processing worksheet {idx_sheet}/{total_targets}: {target}...",
        )

        norm_target = normalize_text(target).lower()
        actual_sheet_name = available_sheets.get(norm_target)

        if not actual_sheet_name:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Missing in Workbook",
                    "ಉಪವಿಭಾಗ Column": "N/A",
                    "Rows Loaded (Valid ವಿಭಾಗ)": 0,
                    "Overdue Rows ('P')": 0,
                    "ETS Rows": 0,
                }
            )
            continue

        try:
            raw_df = read_sheet_safely(excel_file, workbook_bytes, actual_sheet_name)
            if raw_df.empty:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": "Empty Sheet",
                        "ಉಪವಿಭಾಗ Column": "N/A",
                        "Rows Loaded (Valid ವಿಭಾಗ)": 0,
                        "Overdue Rows ('P')": 0,
                        "ETS Rows": 0,
                    }
                )
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
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": "Header Row Not Found",
                        "ಉಪವಿಭಾಗ Column": "N/A",
                        "Rows Loaded (Valid ವಿಭಾಗ)": 0,
                        "Overdue Rows ('P')": 0,
                        "ETS Rows": 0,
                    }
                )
                continue

            raw_headers = [
                normalize_text(col) for col in raw_df.iloc[header_idx].tolist()
            ]
            col_idx_map = find_matching_column_indices(raw_headers)

            # Note: 'subdivision' is intentionally NOT in required_keys because 2 sheets don't have it
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

            missing_keys = required_keys - set(col_idx_map.keys())
            if missing_keys:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": f"Missing Columns: {', '.join(sorted(missing_keys))}",
                        "ಉಪವಿಭಾಗ Column": "Found" if "subdivision" in col_idx_map else "Not in Sheet",
                        "Rows Loaded (Valid ವಿಭಾಗ)": 0,
                        "Overdue Rows ('P')": 0,
                        "ETS Rows": 0,
                    }
                )
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

            # Extract ಉಪವಿಭಾಗ if column exists in this sheet; otherwise treat as empty/unspecified
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

            # Row Inclusion Rule: Strictly based on valid text in 'ವಿಭಾಗ' (even if ಉಪವಿಭಾಗ is empty)
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
                    "ಉಪವಿಭಾಗ Column": "Found" if has_subdiv_col else "Not in Sheet (Included as Empty)",
                    "Rows Loaded (Valid ವಿಭಾಗ)": len(sub),
                    "Overdue Rows ('P')": int(overdue_mask.sum()),
                    "ETS Rows": int((valid_row_mask & is_status_ets).sum()),
                }
            )

        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "ಉಪವಿಭಾಗ Column": "Error",
                    "Rows Loaded (Valid ವಿಭಾಗ)": 0,
                    "Overdue Rows ('P')": 0,
                    "ETS Rows": 0,
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
    """Save parsed dataset in shared RAM and on local disk."""
    now_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %I:%M:%S %p")
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
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Dashboard Controls")

sheet_url_input = st.sidebar.text_input(
    "Google Spreadsheet Link",
    value=DEFAULT_SHEET_URL,
    help="Ensure the Google Sheet is set to 'Anyone with the link can view'.",
)

uploaded_xlsx = st.sidebar.file_uploader(
    "Or Upload Workbook (.xlsx) Directly",
    type=["xlsx"],
    help="Optional: Upload a downloaded .xlsx copy to update all screens immediately.",
)

sync_clicked = st.sidebar.button(
    "🔄 Sync Latest Data from Google Sheets",
    use_container_width=True,
    type="primary",
)

denominator_mode = st.sidebar.radio(
    "Total Works Assigned Count Basis:",
    options=[
        "All Works Assigned (All Statuses)",
        "Only Ongoing Works (Status = 'Progress')",
    ],
    index=0,
)

chart_style = st.sidebar.radio(
    "Bar Width Mode:",
    options=[
        "Overlay (Overdue + ETS + Rest of Works)",
        "Scale by Overdue Works (Full-Width Bars)",
        "Scale by Overdue Percentage (%)",
    ],
    index=0,
    help="Choose how the horizontal bars are scaled across the page.",
)

shrink_outliers = st.sidebar.checkbox(
    f"Shrink long bars above {BAR_SHRINK_THRESHOLD} works",
    value=True,
    help=f"Compresses bars above {BAR_SHRINK_THRESHOLD} works so all other bars look long and readable.",
)

hide_zero_overdue = st.sidebar.checkbox(
    "Show only officers with Overdue > 0",
    value=False,
)

show_unassigned_bar = st.sidebar.checkbox(
    "Include Empty / Unassigned Officer bar at the bottom",
    value=True,
    help="Groups all works where the officer name is blank into a single bar pinned at the bottom.",
)

# -----------------------------------------------------------------------------
# MAIN UI & THREAD-SAFE DATA LOADING
# -----------------------------------------------------------------------------
st.title("📊 Minor Irrigation — Overdue Works Dashboard")

shared_state = get_shared_state_v9()

if uploaded_xlsx is not None:
    file_id = f"{uploaded_xlsx.name}_{uploaded_xlsx.size}_{SCHEMA_VERSION}"
    if st.session_state.get("_last_uploaded_id") != file_id:
        with shared_state["lock"]:
            parsed = parse_xlsx_workbook(uploaded_xlsx.getvalue())
            save_shared_state(
                shared_state, parsed, f"Uploaded File ({uploaded_xlsx.name})"
            )
            st.session_state["_last_uploaded_id"] = file_id

elif (
    shared_state["data"] is None
    or shared_state.get("version") != SCHEMA_VERSION
    or sync_clicked
):
    with shared_state["lock"]:
        if (
            shared_state["data"] is None
            or shared_state.get("version") != SCHEMA_VERSION
            or sync_clicked
        ):
            with st.spinner(
                "Downloading workbook snapshot from Google Sheets (takes ~15-20s)..."
            ):
                try:
                    xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                    parsed = parse_xlsx_workbook(xlsx_bytes)
                    save_shared_state(
                        shared_state, parsed, "Google Sheets Live Snapshot"
                    )
                except Exception as e:
                    if shared_state["data"] is not None:
                        st.warning(
                            f"Google Sheets is busy right now ({e}). Showing last synced snapshot from {shared_state['last_synced']}."
                        )
                    else:
                        st.error(
                            f"Could not download from Google Sheets right now.\n\n"
                            f"**Tip:** Wait 10 seconds and click **Sync Latest Data**, or upload the `.xlsx` file in the sidebar.\n\n"
                            f"Details: {e}"
                        )
                        st.stop()

master_df, overdue_full_df, diag_df, today_str = shared_state["data"]

st.caption(
    f"⚡ **Data Source:** {shared_state['source_label']} &nbsp;|&nbsp; "
    f"**Last Synced:** {shared_state['last_synced']} IST &nbsp;|&nbsp; "
    f"*(Row Count Rule: Rows with valid text in **'{COL_DIV}'** column — empty **'{COL_SUBDIV}'** rows are also included)*"
)

if master_df.empty:
    st.warning("No matching data rows were found across the target worksheets.")
    st.dataframe(diag_df, use_container_width=True)
    st.stop()

st.markdown("<div style='margin-top: 0.9rem;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# TOP CONTROLS: 4 SPACED-OUT BORDERED CARDS (ವಿಭಾಗ, ಉಪವಿಭಾಗ, OFFICER VIEW, SORT)
# -----------------------------------------------------------------------------
unique_divisions = sorted(
    [d for d in master_df["division_name"].dropna().unique().tolist() if d != ""]
)
division_options = [ALL_DIVISIONS_OPTION] + unique_divisions

ctrl_col1, ctrl_col2, ctrl_col3, ctrl_col4 = st.columns(
    [1.15, 1.15, 1.25, 1.3], gap="medium"
)

with ctrl_col1:
    with st.container(border=True):
        selected_division = st.selectbox(
            f"🏛️ {COL_DIV} ಆಯ್ಕೆಮಾಡಿ / Select {COL_DIV}:",
            options=division_options,
            index=0,
        )

# Filter master_df by selected ವಿಭಾಗ first so ಉಪವಿಭಾಗ options cascade dynamically
if selected_division != ALL_DIVISIONS_OPTION:
    div_filtered_df = master_df[
        master_df["division_name"] == selected_division
    ].copy()
else:
    div_filtered_df = master_df.copy()

# Build dynamic ಉಪವಿಭಾಗ options for the selected ವಿಭಾಗ
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

with ctrl_col2:
    with st.container(border=True):
        selected_subdivision = st.selectbox(
            f"🏢 {COL_SUBDIV} ಆಯ್ಕೆಮಾಡಿ / Select {COL_SUBDIV}:",
            options=subdivision_options,
            index=0,
            help=(
                "Shows sub-divisions under the selected ವಿಭಾಗ. "
                "'ALL' includes all sub-divisions as well as rows where ಉಪವಿಭಾಗ is empty."
            ),
        )

with ctrl_col3:
    with st.container(border=True):
        view_choice = st.radio(
            "👤 ವೀಕ್ಷಣೆ ಆಯ್ಕೆಮಾಡಿ / Select Officer View:",
            options=[COL_SO, COL_AEE],
            horizontal=False,
        )

with ctrl_col4:
    with st.container(border=True):
        sort_choice = st.radio(
            "⬇️ ಕ್ರಮಬದ್ಧಗೊಳಿಸಿ / Sort Chart & Table By:",
            options=[
                "Highest Overdue Count at Top",
                "Highest Overdue Percentage (%) at Top",
            ],
            horizontal=False,
        )

st.markdown("<div style='margin-top: 0.6rem;'></div>", unsafe_allow_html=True)

st.sidebar.markdown(
    f"**Active {COL_DIV}:** `{selected_division}`\n\n"
    f"**Active {COL_SUBDIV}:** `{selected_subdivision}`"
)

group_col = "so_name" if view_choice == COL_SO else "aee_name"
full_filter_col = "_so_name" if view_choice == COL_SO else "_aee_name"

# -----------------------------------------------------------------------------
# APPLY ಉಪವಿಭಾಗ FILTER TO ALL VALUES
# -----------------------------------------------------------------------------
if selected_subdivision != ALL_SUBDIVISIONS_OPTION:
    subdiv_filtered_df = div_filtered_df[
        div_filtered_df["subdivision_name"] == selected_subdivision
    ].copy()
else:
    subdiv_filtered_df = div_filtered_df.copy()

valid_df = subdiv_filtered_df[subdiv_filtered_df[group_col] != ""].copy()

if not show_unassigned_bar:
    valid_df = valid_df[valid_df[group_col] != UNASSIGNED_LABEL].copy()

if denominator_mode.startswith("Only Ongoing Works"):
    valid_df = valid_df[valid_df["is_status_p"]].copy()

if valid_df.empty:
    st.warning(
        f"No works found for **{COL_DIV}: {selected_division}** | **{COL_SUBDIV}: {selected_subdivision}**."
    )
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

if sort_choice == "Highest Overdue Percentage (%) at Top":
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
# KPI SUMMARY METRICS (5 COLUMNS, REDUCED BOTTOM GAP)
# -----------------------------------------------------------------------------
total_assigned_all = int(summary_df["total_works"].sum())
total_overdue_all = int(summary_df["overdue_works"].sum())
total_ets_all = int(summary_df["ets_works"].sum())
named_officers_count = int((~summary_df["_is_unassigned"]).sum())
unassigned_works_count = int(
    summary_df.loc[summary_df["_is_unassigned"], "total_works"].sum()
)

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
filter_scope_label = (
    f"{COL_DIV}: {div_badge} | {COL_SUBDIV}: {subdiv_badge}"
)

k1, k2, k3, k4, k5 = st.columns([1.15, 1.0, 1.05, 1.35, 1.0], gap="medium")
k1.metric("📅 Date (Today)", today_str)
k2.metric(
    "👥 Officers",
    f"{named_officers_count:,}",
    help=f"Officers in {view_choice} (plus {unassigned_works_count:,} unassigned works).",
)
k3.metric("📋 Total Works", f"{total_assigned_all:,}")
k4.metric(
    "🔴 Overdue Works ('P')",
    f"{total_overdue_all:,} ({overall_overdue_pct}%)",
)
k5.metric("🟢 'ETS' Works", f"{total_ets_all:,}")

st.divider()

# -----------------------------------------------------------------------------
# HORIZONTAL BAR CHART (ALL LABELS 100% VISIBLE ON SCREEN)
# -----------------------------------------------------------------------------
st.subheader(
    f"📈 Overdue & ETS Works by {view_choice} — [{filter_scope_label}] ({sort_choice})"
)

if hide_zero_overdue:
    chart_source_df = summary_df[
        (summary_df["overdue_works"] > 0) | summary_df["_is_unassigned"]
    ].copy()
else:
    chart_source_df = summary_df.copy()

if chart_source_df.empty:
    st.info("No officers match the current filter.")
else:
    plot_df = chart_source_df.iloc[::-1].copy()
    plot_df["y_short"] = plot_df[view_choice].map(lambda s: shorten_label(s, 28))
    plot_df["is_shrunk"] = shrink_outliers & (
        plot_df["total_works"] > BAR_SHRINK_THRESHOLD
    )

    def _build_bar_label(r) -> str:
        shrink_tag = (
            f" <span style='color:#FBBF24'>[>{BAR_SHRINK_THRESHOLD} Shrunk]</span>"
            if r["is_shrunk"]
            and chart_style == "Overlay (Overdue + ETS + Rest of Works)"
            else ""
        )
        if sort_choice == "Highest Overdue Percentage (%) at Top":
            return (
                f"  <b>{r['overdue_pct']:.1f}% Overdue</b> "
                f"({int(r['overdue_works'])} Overdue | "
                f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
                f"{int(r['total_works'])} Total){shrink_tag}"
            )
        return (
            f"  <b>{int(r['overdue_works'])} Overdue</b> | "
            f"<span style='color:#4ADE80'>{int(r['ets_works'])} ETS</span> / "
            f"{int(r['total_works'])} Total ({r['overdue_pct']:.1f}%){shrink_tag}"
        )

    plot_df["bar_label"] = plot_df.apply(_build_bar_label, axis=1)

    fig = go.Figure()

    if chart_style == "Scale by Overdue Works (Full-Width Bars)":
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_works"],
                name="Overdue Works ('P')",
                orientation="h",
                marker=dict(
                    color="#EF4444",
                    line=dict(color="#DC2626", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13),
                cliponaxis=False,
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "ongoing_p_works",
                        "ets_works",
                        "rest_works",
                        "overdue_pct",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>{COL_DIV}:</b> {div_badge}<br>"
                    f"<b>{COL_SUBDIV}:</b> {subdiv_badge}<br>"
                    "<b>Overdue Works ('P'):</b> %{x}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[3]}<br>"
                    "<b>Rest of Works:</b> %{customdata[4]}<br>"
                    "<b>Total Works Assigned:</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[5]:.1f}%<extra></extra>"
                ),
            )
        )
        max_x = max(int(plot_df["overdue_works"].max()), 1)
        x_range_max = max_x * 1.38
        x_axis_title = "Number of Overdue Works"

    elif chart_style == "Scale by Overdue Percentage (%)":
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_pct"],
                name="Overdue %",
                orientation="h",
                marker=dict(
                    color="#EF4444",
                    line=dict(color="#DC2626", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13),
                cliponaxis=False,
                customdata=plot_df[
                    [
                        view_choice,
                        "total_works",
                        "ongoing_p_works",
                        "overdue_works",
                        "ets_works",
                        "rest_works",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>{COL_DIV}:</b> {div_badge}<br>"
                    f"<b>{COL_SUBDIV}:</b> {subdiv_badge}<br>"
                    "<b>Overdue Percentage:</b> %{x:.1f}%<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[3]}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[4]}<br>"
                    "<b>Rest of Works:</b> %{customdata[5]}<br>"
                    "<b>Total Works Assigned:</b> %{customdata[1]}<extra></extra>"
                ),
            )
        )
        max_x = max(float(plot_df["overdue_pct"].max()), 1.0)
        x_range_max = max_x * 1.38
        x_axis_title = "Overdue Percentage (%)"

    else:
        plot_df["overdue_plus_ets"] = plot_df["overdue_works"] + plot_df["ets_works"]

        if shrink_outliers:
            plot_df["vis_total"] = plot_df["total_works"].map(compress_bar_length)
            plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].map(
                compress_bar_length
            )
            plot_df["vis_overdue"] = plot_df["overdue_works"].map(compress_bar_length)
        else:
            plot_df["vis_total"] = plot_df["total_works"].astype(float)
            plot_df["vis_overdue_plus_ets"] = plot_df["overdue_plus_ets"].astype(float)
            plot_df["vis_overdue"] = plot_df["overdue_works"].astype(float)

        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["vis_total"],
                name="Rest of Works (Other Statuses)",
                orientation="h",
                marker=dict(
                    color="rgba(148, 163, 184, 0.30)",
                    line=dict(color="#64748B", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13),
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
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>{COL_DIV}:</b> {div_badge}<br>"
                    f"<b>{COL_SUBDIV}:</b> {subdiv_badge}<br>"
                    "<b>Total Works Assigned (Actual):</b> %{customdata[1]}<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[2]}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[3]}<br>"
                    "<b>Rest of Works:</b> %{customdata[4]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[5]:.1f}%<extra></extra>"
                ),
            )
        )

        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["vis_overdue_plus_ets"],
                name="ETS Works ('ETS')",
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
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>{COL_DIV}:</b> {div_badge}<br>"
                    f"<b>{COL_SUBDIV}:</b> {subdiv_badge}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[3]}<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[2]}<br>"
                    "<b>Rest of Works:</b> %{customdata[4]}<br>"
                    "<b>Total Works Assigned (Actual):</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[5]:.1f}%<extra></extra>"
                ),
            )
        )

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
                        "ongoing_p_works",
                        "overdue_pct",
                        "ets_works",
                        "rest_works",
                        "overdue_works",
                    ]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>{COL_DIV}:</b> {div_badge}<br>"
                    f"<b>{COL_SUBDIV}:</b> {subdiv_badge}<br>"
                    "<b>Overdue Works ('P'):</b> %{customdata[6]}<br>"
                    "<b>ETS Works ('ETS'):</b> %{customdata[4]}<br>"
                    "<b>Rest of Works:</b> %{customdata[5]}<br>"
                    "<b>Total Works Assigned (Actual):</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[3]:.1f}%<extra></extra>"
                ),
            )
        )

        # Ensure full text labels on the longest bar never clip at the right edge
        max_x = max(float(plot_df["vis_total"].max()), 1.0)
        x_range_max = max_x * 1.38
        x_axis_title = "Number of Works"

    chart_height = max(500, len(plot_df) * 36 + 90)

    fig.update_layout(
        barmode="overlay",
        height=chart_height,
        bargap=0.22,
        xaxis=dict(
            title=x_axis_title,
            range=[0, x_range_max],
            gridcolor="rgba(200, 200, 200, 0.2)",
            zeroline=True,
        ),
        yaxis=dict(
            title="",
            automargin=False,
            tickmode="array",
            tickvals=plot_df[view_choice],
            ticktext=plot_df["y_short"],
            tickfont=dict(size=13),
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.01,
            xanchor="left",
            x=0,
            traceorder="reversed",
        ),
        margin=dict(l=225, r=50, t=35, b=40),
    )

    st.plotly_chart(fig, use_container_width=True)

# -----------------------------------------------------------------------------
# DETAILED SUMMARY TABLE (FILTERED BY ವಿಭಾಗ & ಉಪವಿಭಾಗ)
# -----------------------------------------------------------------------------
st.subheader(f"📋 Officer-Wise Summary Table — [{filter_scope_label}]")

display_table = summary_df[
    [
        view_choice,
        "total_works",
        "overdue_works",
        "ets_works",
        "rest_works",
        "ongoing_p_works",
        "overdue_pct",
    ]
].rename(
    columns={
        "total_works": "Total Works Assigned",
        "overdue_works": "Overdue Works ('P')",
        "ets_works": "ETS Works ('ETS')",
        "rest_works": "Rest of Works",
        "ongoing_p_works": "Ongoing Works ('P')",
        "overdue_pct": "Overdue %",
    }
)
display_table.index = display_table.index + 1

st.dataframe(
    display_table.style.format({"Overdue %": "{:.1f}%"}),
    use_container_width=True,
)

st.divider()

# -----------------------------------------------------------------------------
# BOTTOM SECTION: OFFICER DROPDOWN & COMBINED OVERDUE WORKS TABLE + DOWNLOAD
# -----------------------------------------------------------------------------
st.subheader(
    f"📂 ಅಧಿಕಾರಿಗಳ ವಾರು ಅವಧಿ ಮೀರಿದ ಕಾಮಗಾರಿಗಳ ವಿವರ / Combined Overdue Works by {view_choice} — [{filter_scope_label}]"
)

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
    f"Select {view_choice} (Filtered by {filter_scope_label}) to view and download their Overdue Works:",
    options=officer_options,
    index=0 if officer_options else None,
    format_func=lambda name: (
        f"{name}  —  ({officer_stats_map[name][0]} Overdue | "
        f"{officer_stats_map[name][1]} ETS / "
        f"{officer_stats_map[name][2]} Total | {officer_stats_map[name][3]:.1f}%)"
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
            st.success(
                f"🎉 **{selected_officer}** has **0** overdue works in **{filter_scope_label}**!"
            )
        else:
            sheets_involved = officer_overdue_df["Worksheet / ಯೋಜನೆ"].unique().tolist()
            st.markdown(
                f"**{COL_DIV}:** `{div_badge}` &nbsp;|&nbsp; "
                f"**{COL_SUBDIV}:** `{subdiv_badge}` &nbsp;|&nbsp; "
                f"**Selected {view_choice}:** `{selected_officer}` &nbsp;|&nbsp; "
                f"**Total Overdue Works:** `{len(officer_overdue_df)}` &nbsp;|&nbsp; "
                f"**Worksheets:** `{', '.join(sheets_involved)}`"
            )

            safe_name = (
                re.sub(r"[^\w\-]+", "_", selected_officer).strip("_") or "Unassigned"
            )
            safe_div = (
                re.sub(r"[^\w\-]+", "_", div_badge).strip("_") or "All_Divisions"
            )
            safe_subdiv = (
                re.sub(r"[^\w\-]+", "_", subdiv_badge).strip("_") or "All_SubDivisions"
            )
            excel_bytes = convert_df_to_excel_bytes(officer_overdue_df.reset_index())
            csv_bytes = (
                officer_overdue_df.reset_index().to_csv(index=False).encode("utf-8-sig")
            )

            dl_col1, dl_col2, _ = st.columns([1.5, 1.5, 5])
            with dl_col1:
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Overdue_Works_{safe_div}_{safe_subdiv}_{safe_name}_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl_col2:
                st.download_button(
                    label="📥 Download CSV (.csv)",
                    data=csv_bytes,
                    file_name=f"Overdue_Works_{safe_div}_{safe_subdiv}_{safe_name}_{today_str}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            st.dataframe(
                officer_overdue_df,
                use_container_width=True,
                height=min(600, max(220, (len(officer_overdue_df) + 1) * 42)),
            )

with st.expander("🔍 Inspect Worksheet Loading Status (23 Target Sheets)"):
    st.dataframe(diag_df, use_container_width=True)
