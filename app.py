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
            padding-left: 1.2rem !important;
            padding-right: 1.2rem !important;
            padding-top: 2.8rem !important;
            padding-bottom: 2rem !important;
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

# Standard known CP/CF status codes (also merged dynamically with any new terms in the sheet)
STANDARD_STATUS_CODES = ["P", "CP", "CF", "TBS", "TP", "ETS", "ETH", "OTH"]

DEFAULT_SHEET_URL = st.secrets.get(
    "SHEET_URL",
    "https://docs.google.com/spreadsheets/d/13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869",
)

DISK_BACKUP_FILE = "/tmp/mi_overdue_shared_snapshot_v2.pkl"


# -----------------------------------------------------------------------------
# SHARED SERVER MEMORY & LOCK (Prevents 410 Gone when multiple screens connect)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_shared_state() -> dict:
    """
    Shared across ALL browser sessions/screens connected to this Streamlit app.
    Ensures only one download happens at a time and all screens share the data.
    """
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
    """Collapse multiple spaces/newlines and strip text."""
    if val is None or pd.isna(val):
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def format_cell_for_display(val) -> str:
    """Format raw Excel cell values cleanly for table display and export."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%d-%m-%Y")
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
    """Find required 0-based column positions across sheets."""
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
    """Standardize key column names across sheets and deduplicate repeated headers."""
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
    """Safely parse dates (handling datetime objects and DD-MM-YYYY / DD.MM.YYYY strings)."""
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
    """Check if the 'ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ' cell is empty or placeholder text."""
    def _check_empty(val):
        if val is None or pd.isna(val):
            return True
        s = normalize_text(val)
        return s.upper() in {"", "-", "--", "NIL", "NA", "N/A", "NONE", "NAT"}

    return series.map(_check_empty)


def is_valid_officer_name(name: str) -> bool:
    """Filter out empty cells, serial numbers, or summary rows like 'ಒಟ್ಟು' / 'Total'."""
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


def is_valid_status_token(token: str) -> bool:
    """Filter out summary/header noise when building the status dropdown list."""
    if not token or len(token) > 25:
        return False
    if "CP/CF" in token or "ಒಟ್ಟು" in token or "TOTAL" in token:
        return False
    if token.isdigit():
        return False
    return True


def shorten_label(name: str, max_len: int = 26) -> str:
    """Keep Y-axis labels compact so the chart always starts on the far left."""
    return name if len(name) <= max_len else name[: max_len - 1] + "…"


def convert_df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Export a DataFrame to an in-memory .xlsx file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Works_Report")
    return output.getvalue()


def parse_xlsx_workbook(workbook_bytes: bytes):
    """
    Parse the 23 target worksheets and store all valid rows so any CP/CF status
    filter ('ALL', 'ETH', 'P, CP', etc.) can be applied instantaneously.
    """
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
    full_work_rows = []
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
                    "Rows Loaded": 0,
                    "Date-Overdue Rows (All Statuses)": 0,
                }
            )
            continue

        try:
            raw_df = pd.read_excel(excel_file, sheet_name=actual_sheet_name, header=None)
            if raw_df.empty:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": "Empty Sheet",
                        "Rows Loaded": 0,
                        "Date-Overdue Rows (All Statuses)": 0,
                    }
                )
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
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": "Header Row Not Found",
                        "Rows Loaded": 0,
                        "Date-Overdue Rows (All Statuses)": 0,
                    }
                )
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

            missing_keys = required_keys - set(col_idx_map.keys())
            if missing_keys:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": f"Missing Columns: {', '.join(sorted(missing_keys))}",
                        "Rows Loaded": 0,
                        "Date-Overdue Rows (All Statuses)": 0,
                    }
                )
                continue

            df = raw_df.iloc[header_idx + 1 :].reset_index(drop=True).copy()

            s_so = df.iloc[:, col_idx_map["so"]].map(normalize_text)
            s_aee = df.iloc[:, col_idx_map["aee"]].map(normalize_text)
            s_status_raw = df.iloc[:, col_idx_map["status"]].map(normalize_text)
            s_status_norm = s_status_raw.str.upper().str.strip()
            s_target_raw = df.iloc[:, col_idx_map["target_date"]]
            s_actual_raw = df.iloc[:, col_idx_map["actual_date"]]

            target_dt = parse_indian_date(s_target_raw)
            actual_empty = is_empty_date(s_actual_raw)

            # Date-based overdue condition:
            # ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ < Today AND ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ is empty
            is_date_overdue = (
                target_dt.notna()
                & (target_dt < today_ist)
                & actual_empty
            )

            valid_row_mask = (s_status_norm != "") | (s_so != "") | (s_aee != "")

            sub = pd.DataFrame(
                {
                    "sheet_name": target,
                    "so_name": s_so,
                    "aee_name": s_aee,
                    "status_raw": s_status_raw,
                    "status_norm": s_status_norm,
                    "target_date_raw": s_target_raw.map(format_cell_for_display),
                    "actual_date_raw": s_actual_raw.map(format_cell_for_display),
                    "target_dt": target_dt,
                    "actual_empty": actual_empty,
                    "is_date_overdue": is_date_overdue,
                }
            )
            sub = sub[valid_row_mask].reset_index(drop=True)
            compiled_rows.append(sub)

            # Store full columns for all valid rows so any selected CP/CF status can be viewed/downloaded
            if valid_row_mask.any():
                unique_headers = build_clean_unique_headers(raw_headers, col_idx_map)
                full_df = df.loc[valid_row_mask].copy()
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
                full_df["_so_name"] = s_so.loc[valid_row_mask].values
                full_df["_aee_name"] = s_aee.loc[valid_row_mask].values
                full_df["_status_norm"] = s_status_norm.loc[valid_row_mask].values
                full_df["_is_date_overdue"] = is_date_overdue.loc[valid_row_mask].values

                full_work_rows.append(full_df.reset_index(drop=True))

            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Loaded",
                    "Rows Loaded": len(sub),
                    "Date-Overdue Rows (All Statuses)": int(
                        (valid_row_mask & is_date_overdue).sum()
                    ),
                }
            )

        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "Rows Loaded": 0,
                    "Date-Overdue Rows (All Statuses)": 0,
                }
            )

    progress.empty()

    master_df = (
        pd.concat(compiled_rows, ignore_index=True)
        if compiled_rows
        else pd.DataFrame()
    )
    all_full_df = (
        pd.concat(full_work_rows, ignore_index=True).fillna("")
        if full_work_rows
        else pd.DataFrame()
    )
    diag_df = pd.DataFrame(sheet_diagnostics)
    return master_df, all_full_df, diag_df, today_ist.strftime("%d-%m-%Y")


def save_shared_state(state: dict, data_tuple, source_label: str):
    """Save parsed dataset in shared RAM and on local disk."""
    now_str = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%m-%Y %I:%M:%S %p")
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

chart_style = st.sidebar.radio(
    "Bar Width Mode:",
    options=[
        "Overlay (Overdue inside Total Works)",
        "Scale by Overdue Works (Full-Width Bars)",
        "Scale by Overdue Percentage (%)",
    ],
    index=0,
    help="Choose how the horizontal bars are scaled across the page.",
)

hide_zero_overdue = st.sidebar.checkbox(
    "Show only officers with Overdue > 0",
    value=False,
)

# -----------------------------------------------------------------------------
# MAIN UI & THREAD-SAFE DATA LOADING
# -----------------------------------------------------------------------------
st.title("📊 Minor Irrigation — Overdue Works Dashboard")

shared_state = get_shared_state()

if uploaded_xlsx is not None:
    file_id = f"{uploaded_xlsx.name}_{uploaded_xlsx.size}"
    if st.session_state.get("_last_uploaded_id") != file_id:
        with shared_state["lock"]:
            parsed = parse_xlsx_workbook(uploaded_xlsx.getvalue())
            save_shared_state(shared_state, parsed, f"Uploaded File ({uploaded_xlsx.name})")
            st.session_state["_last_uploaded_id"] = file_id

elif shared_state["data"] is None or sync_clicked:
    with shared_state["lock"]:
        if shared_state["data"] is None or sync_clicked:
            with st.spinner("Downloading workbook snapshot from Google Sheets (takes ~15-20s)..."):
                try:
                    xlsx_bytes = download_xlsx_bytes(sheet_url_input)
                    parsed = parse_xlsx_workbook(xlsx_bytes)
                    save_shared_state(shared_state, parsed, "Google Sheets Live Snapshot")
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

master_df, all_full_df, diag_df, today_str = shared_state["data"]

st.caption(
    f"⚡ **Data Source:** {shared_state['source_label']} &nbsp;|&nbsp; "
    f"**Last Synced:** {shared_state['last_synced']} IST &nbsp;|&nbsp; "
    f"*(Cached in shared memory — TV & secondary screens load instantaneously)*"
)

if master_df.empty:
    st.warning("No matching data rows were found across the target worksheets.")
    st.dataframe(diag_df, use_container_width=True)
    st.stop()

# -----------------------------------------------------------------------------
# BUILD DYNAMIC CP/CF STATUS OPTIONS & SMART MULTI-SELECT CALLBACK
# -----------------------------------------------------------------------------
detected_statuses = [
    s
    for s in master_df["status_norm"].dropna().unique().tolist()
    if is_valid_status_token(s)
]

ordered_statuses = []
for code in STANDARD_STATUS_CODES + sorted(detected_statuses):
    if code not in ordered_statuses:
        ordered_statuses.append(code)

status_dropdown_options = ["ALL"] + ordered_statuses

if "cp_cf_selected" not in st.session_state:
    st.session_state["cp_cf_selected"] = ["P"]
if "cp_cf_prev" not in st.session_state:
    st.session_state["cp_cf_prev"] = list(st.session_state["cp_cf_selected"])


def on_cp_cf_change():
    """
    Smart multi-select handler:
    - If 'ALL' was active and user clicks 'ETH' -> automatically unselects 'ALL' and selects 'ETH'.
    - If user selects multiple terms like 'P', 'CP' -> keeps both active.
    - If user clicks 'ALL' while specific terms are active -> resets to ['ALL'].
    - If user clears everything -> defaults back to ['ALL'].
    """
    current = st.session_state.get("cp_cf_selected", [])
    prev = st.session_state.get("cp_cf_prev", ["ALL"])

    if not current:
        st.session_state["cp_cf_selected"] = ["ALL"]
    elif "ALL" in current and len(current) > 1:
        if "ALL" not in prev:
            st.session_state["cp_cf_selected"] = ["ALL"]
        else:
            st.session_state["cp_cf_selected"] = [x for x in current if x != "ALL"]

    st.session_state["cp_cf_prev"] = list(st.session_state["cp_cf_selected"])


# -----------------------------------------------------------------------------
# TOP CONTROLS: OFFICER VIEW, CP/CF STATUS DROPDOWN, & SORT ORDER
# -----------------------------------------------------------------------------
ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1.3, 1.6, 1.4])

with ctrl_col1:
    view_choice = st.radio(
        "ವೀಕ್ಷಣೆ ಆಯ್ಕೆಮಾಡಿ / Select Officer View:",
        options=[COL_SO, COL_AEE],
        horizontal=True,
    )

with ctrl_col2:
    selected_statuses = st.multiselect(
        "CP/CF ಸ್ಥಿತಿ ಆಯ್ಕೆಮಾಡಿ / Filter by CP/CF Status (Select 'ALL', single e.g. 'ETH', or multiple e.g. 'P, CP'):",
        options=status_dropdown_options,
        key="cp_cf_selected",
        on_change=on_cp_cf_change,
        help="Select 'ALL' for all works, or pick one ('ETH', 'P') or multiple ('P', 'CP') statuses.",
    )

with ctrl_col3:
    sort_choice = st.radio(
        "ಕ್ರಮಬದ್ಧಗೊಳಿಸಿ / Sort Chart & Table By:",
        options=[
            "Highest Overdue Count at Top",
            "Highest Overdue Percentage (%) at Top",
        ],
        horizontal=True,
    )

# Also show the active status filter in the sidebar for convenience
st.sidebar.markdown(
    f"**Active CP/CF Filter:** `{', '.join(selected_statuses)}`"
)

is_all_statuses = ("ALL" in selected_statuses) or (len(selected_statuses) == 0)
status_badge = "ALL Statuses" if is_all_statuses else ", ".join(selected_statuses)

group_col = "so_name" if view_choice == COL_SO else "aee_name"
full_filter_col = "_so_name" if view_choice == COL_SO else "_aee_name"

# -----------------------------------------------------------------------------
# APPLY CP/CF STATUS FILTER TO ALL VALUES (TOTAL WORKS, OVERDUE WORKS, %)
# -----------------------------------------------------------------------------
valid_df = master_df[master_df[group_col].map(is_valid_officer_name)].copy()

if not is_all_statuses:
    filtered_df = valid_df[valid_df["status_norm"].isin(selected_statuses)].copy()
else:
    filtered_df = valid_df.copy()

# Overdue works within the selected CP/CF status(es)
filtered_df["is_overdue"] = filtered_df["is_date_overdue"]

if filtered_df.empty:
    st.warning(
        f"No works found matching CP/CF status: **{status_badge}**. Try selecting another status above."
    )
    st.stop()

summary_df = (
    filtered_df.groupby(group_col, as_index=False)
    .agg(
        total_works=("is_overdue", "count"),
        overdue_works=("is_overdue", "sum"),
    )
    .rename(columns={group_col: view_choice})
)

summary_df["overdue_pct"] = (
    (summary_df["overdue_works"] / summary_df["total_works"]) * 100
).round(1).fillna(0.0)

if sort_choice == "Highest Overdue Percentage (%) at Top":
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
# KPI SUMMARY METRICS
# -----------------------------------------------------------------------------
total_assigned_all = int(summary_df["total_works"].sum())
total_overdue_all = int(summary_df["overdue_works"].sum())
overall_overdue_pct = (
    round((total_overdue_all / total_assigned_all) * 100, 1)
    if total_assigned_all > 0
    else 0.0
)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("📅 Comparison Date (Today)", today_str)
k2.metric("🏷️ CP/CF Status Filter", status_badge)
k3.metric(f"👥 Total Officers ({view_choice})", len(summary_df))
k4.metric(f"📋 Total Works ({status_badge})", f"{total_assigned_all:,}")
k5.metric(
    f"⚠️ Overdue Works ({status_badge})",
    f"{total_overdue_all:,} ({overall_overdue_pct}%)",
)

st.divider()

# -----------------------------------------------------------------------------
# HORIZONTAL BAR CHART (FULL WIDTH & LEFT ALIGNED)
# -----------------------------------------------------------------------------
st.subheader(
    f"📈 Overdue Works by {view_choice} — [CP/CF: {status_badge}] ({sort_choice})"
)

chart_source_df = (
    summary_df[summary_df["overdue_works"] > 0].copy()
    if hide_zero_overdue
    else summary_df.copy()
)

if chart_source_df.empty:
    st.info(f"No officers have overdue works > 0 for CP/CF status: {status_badge}.")
else:
    plot_df = chart_source_df.iloc[::-1].copy()
    plot_df["y_short"] = plot_df[view_choice].map(lambda s: shorten_label(s, 28))

    if sort_choice == "Highest Overdue Percentage (%) at Top":
        plot_df["bar_label"] = plot_df.apply(
            lambda r: (
                f"  <b>{r['overdue_pct']:.1f}% Overdue</b> "
                f"({int(r['overdue_works'])} Overdue / {int(r['total_works'])} Total)"
            ),
            axis=1,
        )
    else:
        plot_df["bar_label"] = plot_df.apply(
            lambda r: (
                f"  <b>{int(r['overdue_works'])} Overdue</b> / "
                f"{int(r['total_works'])} Total ({r['overdue_pct']:.1f}%)"
            ),
            axis=1,
        )

    fig = go.Figure()

    if chart_style == "Scale by Overdue Works (Full-Width Bars)":
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_works"],
                name=f"Overdue Works ({status_badge})",
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
                    [view_choice, "total_works", "overdue_pct"]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>CP/CF Filter:</b> {status_badge}<br>"
                    "<b>Overdue Works:</b> %{x}<br>"
                    "<b>Total Works (Selected Status):</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[2]:.1f}%<extra></extra>"
                ),
            )
        )
        max_x = max(int(plot_df["overdue_works"].max()), 1)
        x_range_max = max_x * 1.30
        x_axis_title = f"Number of Overdue Works ({status_badge})"

    elif chart_style == "Scale by Overdue Percentage (%)":
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_pct"],
                name=f"Overdue % ({status_badge})",
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
                    [view_choice, "total_works", "overdue_works"]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>CP/CF Filter:</b> {status_badge}<br>"
                    "<b>Overdue Percentage:</b> %{x:.1f}%<br>"
                    "<b>Overdue Works:</b> %{customdata[2]}<br>"
                    "<b>Total Works (Selected Status):</b> %{customdata[1]}<extra></extra>"
                ),
            )
        )
        max_x = max(float(plot_df["overdue_pct"].max()), 1.0)
        x_range_max = max_x * 1.30
        x_axis_title = f"Overdue Percentage (%) — [{status_badge}]"

    else:
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["total_works"],
                name=f"Total Works ({status_badge})",
                orientation="h",
                marker=dict(
                    color="rgba(148, 163, 184, 0.30)",
                    line=dict(color="#64748B", width=1),
                ),
                text=plot_df["bar_label"],
                textposition="outside",
                textfont=dict(size=13),
                cliponaxis=False,
                customdata=plot_df[[view_choice]],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>CP/CF Filter:</b> {status_badge}<br>"
                    "<b>Total Works:</b> %{x}<extra></extra>"
                ),
            )
        )
        fig.add_trace(
            go.Bar(
                y=plot_df[view_choice],
                x=plot_df["overdue_works"],
                name=f"Overdue Works ({status_badge})",
                orientation="h",
                marker=dict(color="#EF4444"),
                customdata=plot_df[
                    [view_choice, "total_works", "overdue_pct"]
                ],
                hovertemplate=(
                    f"<b>{view_choice}:</b> %{{customdata[0]}}<br>"
                    f"<b>CP/CF Filter:</b> {status_badge}<br>"
                    "<b>Overdue Works:</b> %{x}<br>"
                    "<b>Total Works (Selected Status):</b> %{customdata[1]}<br>"
                    "<b>Overdue Percentage:</b> %{customdata[2]:.1f}%<extra></extra>"
                ),
            )
        )
        max_x = max(int(plot_df["total_works"].max()), 1)
        x_range_max = max_x * 1.32
        x_axis_title = f"Number of Works ({status_badge})"

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
        ),
        margin=dict(l=210, r=40, t=40, b=40),
    )

    st.plotly_chart(fig, use_container_width=True)

# -----------------------------------------------------------------------------
# DETAILED SUMMARY TABLE
# -----------------------------------------------------------------------------
st.subheader(f"📋 Officer-Wise Summary Table — [CP/CF: {status_badge}]")

display_table = summary_df.copy()
display_table.insert(1, "CP/CF Filter", status_badge)
display_table = display_table.rename(
    columns={
        "total_works": f"Total Works ({status_badge})",
        "overdue_works": f"Overdue Works ({status_badge})",
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
# BOTTOM SECTION: OFFICER DROPDOWN & COMBINED WORKS TABLE + DOWNLOAD
# -----------------------------------------------------------------------------
st.subheader(
    f"📂 ಅಧಿಕಾರಿಗಳ ವಾರು ಕಾಮಗಾರಿಗಳ ವಿವರ / Combined Works by {view_choice} — [CP/CF: {status_badge}]"
)

officer_options = summary_df[view_choice].tolist()

officer_stats_map = {
    row[view_choice]: (
        int(row["overdue_works"]),
        int(row["total_works"]),
        float(row["overdue_pct"]),
    )
    for _, row in summary_df.iterrows()
}

b_col1, b_col2 = st.columns([2.2, 1.3])

with b_col1:
    selected_officer = st.selectbox(
        f"Select {view_choice} (Filtered by CP/CF: {status_badge}):",
        options=officer_options,
        index=0 if officer_options else None,
        format_func=lambda name: (
            f"{name}  —  ({officer_stats_map[name][0]} Overdue / "
            f"{officer_stats_map[name][1]} Total in [{status_badge}] | {officer_stats_map[name][2]:.1f}%)"
        ),
    )

with b_col2:
    table_row_mode = st.radio(
        "Rows to Display & Download:",
        options=[
            "Overdue Works Only",
            "All Works (for Selected Status)",
        ],
        horizontal=True,
    )

if selected_officer:
    if all_full_df.empty:
        st.info("No work rows found across the 23 worksheets.")
    else:
        officer_df = all_full_df[
            all_full_df[full_filter_col] == selected_officer
        ].copy()

        if not is_all_statuses:
            officer_df = officer_df[
                officer_df["_status_norm"].isin(selected_statuses)
            ].copy()

        if table_row_mode == "Overdue Works Only":
            officer_df = officer_df[officer_df["_is_date_overdue"]].copy()

        internal_cols = [
            c
            for c in ["_so_name", "_aee_name", "_status_norm", "_is_date_overdue"]
            if c in officer_df.columns
        ]
        officer_df.drop(columns=internal_cols, inplace=True)

        extra_unnamed = [
            c
            for c in officer_df.columns
            if c.startswith("Unnamed_Col_") and officer_df[c].eq("").all()
        ]
        if extra_unnamed:
            officer_df.drop(columns=extra_unnamed, inplace=True)

        officer_df = officer_df.reset_index(drop=True)
        officer_df.index = officer_df.index + 1
        officer_df.index.name = "Sl. No."

        if officer_df.empty:
            st.success(
                f"🎉 **{selected_officer}** has **0** {table_row_mode.lower()} under CP/CF status **{status_badge}**!"
            )
        else:
            sheets_involved = officer_df["Worksheet / ಯೋಜನೆ"].unique().tolist()
            st.markdown(
                f"**Selected {view_choice}:** `{selected_officer}` &nbsp;|&nbsp; "
                f"**CP/CF Filter:** `{status_badge}` &nbsp;|&nbsp; "
                f"**Rows Shown ({table_row_mode}):** `{len(officer_df)}` &nbsp;|&nbsp; "
                f"**Worksheets:** `{', '.join(sheets_involved)}`"
            )

            safe_name = re.sub(r"[^\w\-]+", "_", selected_officer).strip("_") or "Officer"
            safe_status = re.sub(r"[^\w\-]+", "_", status_badge).strip("_") or "ALL"
            excel_bytes = convert_df_to_excel_bytes(officer_df.reset_index())
            csv_bytes = officer_df.reset_index().to_csv(index=False).encode("utf-8-sig")

            dl_col1, dl_col2, _ = st.columns([1.5, 1.5, 5])
            with dl_col1:
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Works_{safe_name}_{safe_status}_{today_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl_col2:
                st.download_button(
                    label="📥 Download CSV (.csv)",
                    data=csv_bytes,
                    file_name=f"Works_{safe_name}_{safe_status}_{today_str}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            st.dataframe(
                officer_df,
                use_container_width=True,
                height=min(600, max(220, (len(officer_df) + 1) * 42)),
            )

with st.expander("🔍 Inspect Worksheet Loading Status (23 Target Sheets)"):
    st.dataframe(diag_df, use_container_width=True)
