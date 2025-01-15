import streamlit as st  # type: ignore
import gspread  # type: ignore
from oauth2client.service_account import ServiceAccountCredentials  # type: ignore
from datetime import datetime
from collections import defaultdict
import difflib
import re
import pandas as pd  # type: ignore
import logging
import urllib.parse
import json
import os
# ------------------------------------------------------------------------------
# Setup Logging
# ------------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ------------------------------------------------------------------------------
# User Authentication Setup
# ------------------------------------------------------------------------------
# Define a simple user database
# In a production environment, consider using a secure method to handle user credentials
USER_DB_PATH = "users.json"

def load_user_db(path: str):
    """
    Load the user database from a JSON file.
    """
    if not os.path.exists(path):
        logger.error(f"User database file '{path}' not found.")
        st.error(f"User database file '{path}' not found.")
        return {}
    try:
        with open(path, "r") as f:
            user_db = json.load(f)
        logger.info("User database loaded successfully.")
        return user_db
    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON from '{path}': {e}")
        st.error(f"Error decoding JSON from '{path}': {e}")
        return {}
    except Exception as e:
        logger.error(f"Unexpected error loading user database: {e}")
        st.error(f"Unexpected error loading user database: {e}")
        return {}

USER_DB = load_user_db(USER_DB_PATH)


# ------------------------------------------------------------------------------
# Initialize Session State for Authentication
# ------------------------------------------------------------------------------
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'username' not in st.session_state:
    st.session_state.username = ""
if 'user_companies' not in st.session_state:
    st.session_state.user_companies = []

# ------------------------------------------------------------------------------
# Authentication Interface
# ------------------------------------------------------------------------------
def login():
    st.title("🔒 Training & Parade Management - Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in USER_DB and USER_DB[username]["password"] == password:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.session_state.user_companies = USER_DB[username]["companies"]
            logger.info(f"User '{username}' authenticated successfully.")
            st.success(f"Welcome, {username}!")
            # Rerun to display the main app
            st.rerun()
        else:
            st.error("Invalid username or password.")
            logger.warning(f"Failed login attempt for username '{username}'.")

def logout():
    st.sidebar.button("Logout", on_click=lambda: logout_callback())

def logout_callback():
    st.session_state.authenticated = False
    st.session_state.username = ""
    st.session_state.user_companies = []
    st.success("You have been logged out.")
    logger.info("User logged out.")
    st.experimental_rerun()

# Show login if not authenticated
if not st.session_state.authenticated:
    login()
    st.stop()

# ------------------------------------------------------------------------------
# If authenticated, display the main app
# ------------------------------------------------------------------------------

# ------------------------------------------------------------------------------
# 1) Must be the very first Streamlit command:
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Training & Parade Management", layout="centered")

# ------------------------------------------------------------------------------
# 2) Google Sheets Setup (Cached)
# ------------------------------------------------------------------------------
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPES)

# List of available companies and their corresponding spreadsheet names
COMPANY_SPREADSHEETS = {
    "Alpha": "Alpha",
    "Bravo": "Bravo",
    "Charlie": "Charlie",
    "Support": "Support",
    "MSC": "MSC"  # Added MSC
}

@st.cache_resource
def get_sheets(selected_company: str):
    """
    Open the spreadsheet based on the selected company and return references to worksheets.
    This is cached to avoid re-opening on each script run unless the company changes.
    """
    spreadsheet_name = COMPANY_SPREADSHEETS.get(selected_company)
    if not spreadsheet_name:
        logger.error(f"Spreadsheet for company '{selected_company}' not found.")
        st.error(f"Spreadsheet for company '{selected_company}' not found.")
        return None
    try:
        gc = gspread.authorize(creds)
        sh = gc.open(spreadsheet_name)
        return {
            "nominal": sh.worksheet("Nominal_Roll"),
            "parade": sh.worksheet("Parade_State"),
            "conducts": sh.worksheet("Conducts")
        }
    except Exception as e:
        logger.error(f"Error accessing spreadsheet '{spreadsheet_name}': {e}")
        st.error(f"Error accessing spreadsheet '{spreadsheet_name}': {e}")
        return None

# ------------------------------------------------------------------------------
# 3) Helper Functions + Caching
# ------------------------------------------------------------------------------

def ensure_str(value) -> str:
    """
    Convert any value to a string and strip leading/trailing whitespaces.
    """
    if value is None:
        return ""
    return str(value).strip()

def is_valid_4d(four_d: str) -> str:
    """
    Validate and format the 4D_Number.
    If the '4D' prefix is missing, add it.
    Returns the formatted 4D_Number if valid, else returns an empty string.
    """
    four_d = ensure_str(four_d).upper()
    if not four_d.startswith('4D'):
        four_d = f'4D{four_d}'
    
    if re.match(r'^4D\d+$', four_d):
        return four_d
    else:
        logger.error(f"Invalid 4D_Number format: {four_d}")
        return ""

def ensure_date_str(date_value) -> str:
    """
    Ensure that the date is a string in DDMMYYYY format with leading zeros.
    If the input is an integer or float, convert it to a string with leading zeros.
    If it's a string, pad with leading zeros if necessary.
    """
    if isinstance(date_value, int):
        return f"{date_value:08d}"
    elif isinstance(date_value, float):
        return f"{int(date_value):08d}"
    elif isinstance(date_value, str):
        # Remove any non-digit characters and pad with leading zeros
        cleaned = re.sub(r'\D', '', date_value)
        return cleaned.zfill(8)
    else:
        return ""

def normalize_name(name: str) -> str:
    """Normalize by uppercase + removing spaces and special characters."""
    return re.sub(r'\W+', '', name.upper())

def get_nominal_records(selected_company: str, _sheet_nominal):
    """
    Returns all rows from Nominal_Roll as a list of dicts.
    Handles case-insensitive and whitespace-trimmed headers.
    """
    records = _sheet_nominal.get_all_records()
    if not records:
        logger.warning(f"No records found in Nominal_Roll for company '{selected_company}'.")
        return []

    # Normalize keys: strip spaces and convert to lower case
    normalized_records = []
    for row in records:
        normalized_row = {k.strip().lower(): v for k, v in row.items()}
        normalized_row['rank'] = ensure_str(normalized_row.get('rank', ''))
        normalized_row['name'] = ensure_str(normalized_row.get('name', ''))
        normalized_row['4d_number'] = is_valid_4d(normalized_row.get('4d_number', ''))
        normalized_row['platoon'] = ensure_str(normalized_row.get('platoon', ''))
        normalized_row['number of leaves left'] = ensure_str(normalized_row.get('number of leaves left', '14'))
        normalized_row['dates taken'] = ensure_str(normalized_row.get('dates taken', ''))
        normalized_records.append(normalized_row)
    
    # Remove any records with invalid 4D_Number
    records = [row for row in normalized_records if row['4d_number']]
    return records

def get_parade_records(selected_company: str, _sheet_parade):
    """
    Returns all rows from Parade_State as a list of dicts, including row numbers.
    Only includes statuses where End_Date is today or in the future.
    Handles case-insensitive and whitespace-trimmed headers.
    """
    today = datetime.today().date()
    all_values = _sheet_parade.get_all_values()  # includes header row at index 0
    if not all_values or len(all_values) < 2:
        logger.warning(f"No records found in Parade_State for company '{selected_company}'.")
        return []
    
    header = [h.strip().lower() for h in all_values[0]]
    records = []
    for idx, row in enumerate(all_values[1:], start=2):  # Start at row 2 in Google Sheets
        if len(row) < len(header):
            logger.warning(f"Skipping malformed row {idx} in Parade_State.")
            continue  # skip malformed row
        record = dict(zip(header, row))
        # Ensure all relevant fields are strings and properly formatted
        record['4d_number'] = is_valid_4d(record.get('4d_number', ''))
        record['platoon'] = ensure_str(record.get('platoon', ''))
        record['start_date_ddmmyyyy'] = ensure_date_str(record.get('start_date_ddmmyyyy', ''))
        record['end_date_ddmmyyyy'] = ensure_date_str(record.get('end_date_ddmmyyyy', ''))
        record['status'] = ensure_str(record.get('status', ''))
        # Parse End_Date to filter out expired statuses
        try:
            end_dt = datetime.strptime(record['end_date_ddmmyyyy'], "%d%m%Y").date()
            if end_dt >= today:
                # Include only active or future statuses
                record['_row_num'] = idx  # Track row number for updating
                records.append(record)
        except ValueError:
            logger.warning(f"Invalid date format in Parade_State for {record.get('4d_number', '')}: {record.get('end_date_ddmmyyyy', '')}")
            continue
    return records

def get_conduct_records(selected_company: str, _sheet_conducts):
    """
    Returns all rows from Conducts as a list of dicts.
    Handles case-insensitive and whitespace-trimmed headers.
    """
    records = _sheet_conducts.get_all_records()
    if not records:
        logger.warning(f"No records found in Conducts for company '{selected_company}'.")
        return []

    # Normalize keys: strip spaces and convert to lower case
    normalized_records = []
    for row in records:
        normalized_row = {k.strip().lower(): v for k, v in row.items()}
        normalized_row['date'] = ensure_date_str(normalized_row.get('date', ''))
        normalized_row['conduct_name'] = ensure_str(normalized_row.get('conduct_name', ''))
        normalized_row['p/t plt1'] = ensure_str(normalized_row.get('p/t plt1', '0/0'))
        normalized_row['p/t plt2'] = ensure_str(normalized_row.get('p/t plt2', '0/0'))
        normalized_row['p/t plt3'] = ensure_str(normalized_row.get('p/t plt3', '0/0'))
        normalized_row['p/t plt4'] = ensure_str(normalized_row.get('p/t plt4', '0/0'))
        normalized_row['p/t alpha'] = ensure_str(normalized_row.get('p/t alpha', '0/0'))
        normalized_row['outliers'] = ensure_str(normalized_row.get('outliers', ''))
        normalized_row['pointers'] = ensure_str(normalized_row.get('pointers', ''))
        normalized_row['submitted_by'] = ensure_str(normalized_row.get('submitted_by', ''))
        normalized_records.append(normalized_row)
    
    return normalized_records

def get_company_strength(platoon: str, records_nominal):
    """
    Count how many rows in Nominal_Roll belong to that platoon.
    """
    return sum(
        1 for row in records_nominal
        if normalize_name(row.get('platoon', '')) == normalize_name(platoon)
    )

def get_company_personnel(platoon: str, records_nominal, records_parade):
    """
    Returns a list of dicts for 'Update Parade' with existing parade statuses first,
    followed by all nominal rows without statuses.
    Each parade status is a separate entry.
    Includes '_row_num' to track the row in Parade_State for status entries.
    """
    parade_map = defaultdict(list)
    for row in records_parade:
        four_d = row.get('4d_number', '').strip().upper()
        parade_map[four_d].append(row)
    
    data_with_status = []
    data_nominal = []
    
    for row in records_nominal:
        p = row.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        four_d = row.get('4d_number', '').strip().upper()
        name = row.get('name', '')
        rank = row.get('rank', '')  # Retrieve Rank
        # Retrieve all parade statuses for the person
        person_parades = parade_map.get(four_d, [])
        for parade in person_parades:
            data_with_status.append({
                'Rank': rank,  # Include Rank
                'Name': name,
                '4D_Number': four_d,
                'Status': parade.get('status', ''),
                'Start_Date': parade.get('start_date_ddmmyyyy', ''),
                'End_Date': parade.get('end_date_ddmmyyyy', ''),
                'Number_of_Leaves_Left': row.get('number of leaves left', 14),
                'Dates_Taken': row.get('dates taken', ''),
                '_row_num': parade.get('_row_num')  # Track row number for updating
            })
        # Add the nominal entry without status
        data_nominal.append({
            'Rank': rank,  # Include Rank
            'Name': name,
            '4D_Number': four_d,
            'Status': '',
            'Start_Date': '',
            'End_Date': '',
            'Number_of_Leaves_Left': row.get('number of leaves left', 14),
            'Dates_Taken': row.get('dates taken', ''),
            '_row_num': None  # No row number for nominal entries without status
        })
    
    # Combine both lists: statuses first, then nominal rows
    combined_data = data_with_status + data_nominal
    return combined_data

def find_name_by_4d(four_d: str, records_nominal) -> str:
    """
    Optional helper: If you want to look up person's Name from Nominal_Roll
    given a 4D_Number.
    """
    four_d = ensure_str(four_d).upper()
    for row in records_nominal:
        if ensure_str(row.get("4d_number", "")).upper() == four_d:
            return ensure_str(row.get("name", ""))
    return ""

def build_onstatus_table(platoon: str, date_obj: datetime, records_nominal, records_parade):
    """
    Return a list of dicts for everyone on status for that date + platoon.
    If multiple statuses exist for the same person, prioritize based on a hierarchy.
    For example: 'Leave' > 'Fever' > 'MC'
    """
    status_priority = {'leave': 3, 'fever': 2, 'mc': 1}  # Define priority
    out = {}
    parade_map = defaultdict(list)
    for row in records_parade:
        four_d = row.get('4d_number', '').strip().upper()
        parade_map[four_d].append(row)
    
    for row in records_nominal:
        p = row.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        four_d = row.get('4d_number', '').strip().upper()
        name = row.get('name', '')
        rank = row.get('rank', '')  # Retrieve Rank
        # Check if person has an active parade status on the given date
        active_status = False
        status_desc = ""
        for parade in parade_map.get(four_d, []):
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', '01012000'), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', '01012000'), "%d%m%Y").date()
                if start_dt <= date_obj.date() <= end_dt:
                    status = ensure_str(parade.get('status', '')).lower()
                    if status in status_priority:
                        if four_d in out:
                            existing_status = out[four_d]['StatusDesc'].lower()
                            if status_priority.get(status, 0) > status_priority.get(existing_status, 0):
                                out[four_d] = {
                                    "Rank": rank,  # Include Rank
                                    "Name": name,
                                    "4D_Number": four_d,
                                    "StatusDesc": ensure_str(parade.get('status', '')),
                                    "Is_Outlier": True
                                }
                        else:
                            out[four_d] = {
                                "Rank": rank,  # Include Rank
                                "Name": name,
                                "4D_Number": four_d,
                                "StatusDesc": ensure_str(parade.get('status', '')),
                                "Is_Outlier": True
                            }
            except ValueError:
                logger.warning(f"Invalid date format for {four_d}: {parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}")
                continue
    logger.info(f"Built on-status table with {len(out)} entries for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
    return list(out.values())

def build_conduct_table(platoon: str, date_obj: datetime, records_nominal, records_parade):
    """
    Return a list of dicts for all personnel in the platoon.
    'Is_Outlier' is True if the person has an active status on the given date, else False.
    Includes 'StatusDesc' for personnel on status.
    Also includes 'Rank'.
    """
    parade_map = defaultdict(list)
    for row in records_parade:
        four_d = row.get('4d_number', '').strip().upper()
        parade_map[four_d].append(row)
    
    data = []
    for person in records_nominal:
        p = person.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        four_d = person.get('4d_number', '').strip().upper()
        name = person.get('name', '')
        rank = person.get('rank', '')  # Retrieve Rank
        # Check if person has an active parade status on the given date
        active_status = False
        status_desc = ""
        for parade in parade_map.get(four_d, []):
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                if start_dt <= date_obj.date() <= end_dt:
                    active_status = True
                    status_desc = parade.get('status', '')
                    break
            except ValueError:
                logger.warning(f"Invalid date format for {four_d}: {parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}")
                continue
        data.append({
            'Rank': rank,  # Include Rank
            'Name': name,
            '4D_Number': four_d,
            'Is_Outlier': active_status,
            'StatusDesc': status_desc if active_status else ""
        })
    logger.info(f"Built conduct table with {len(data)} personnel for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
    return data

def calculate_leaves_used(dates_str: str) -> int:
    """
    Calculate the number of leaves used based on the dates string.
    Dates can be single dates like '05052025' or ranges like '05052025-07052025'.
    """
    if not dates_str:
        return 0

    leaves_used = 0
    # Split by comma to handle multiple entries
    date_entries = [entry.strip() for entry in dates_str.split(',') if entry.strip()]
    for entry in date_entries:
        if '-' in entry:
            # Date range
            start_str, end_str = entry.split('-')
            start_str = ensure_date_str(start_str)
            end_str = ensure_date_str(end_str)
            try:
                start_dt = datetime.strptime(start_str, "%d%m%Y")
                end_dt = datetime.strptime(end_str, "%d%m%Y")
                delta = (end_dt - start_dt).days + 1  # inclusive
                if delta > 0:
                    leaves_used += delta
            except ValueError:
                logger.warning(f"Invalid date range format: {entry}")
                continue  # Skip invalid date formats
        else:
            # Single date
            single_str = ensure_date_str(entry)
            try:
                datetime.strptime(single_str, "%d%m%Y")
                leaves_used += 1
            except ValueError:
                logger.warning(f"Invalid single date format: {entry}")
                continue  # Skip invalid date formats
    logger.info(f"Calculated leaves used: {leaves_used} from dates: {dates_str}")
    return leaves_used

def has_overlapping_status(four_d: str, new_start: datetime, new_end: datetime, records_parade):
    """
    Check if the new status dates overlap with existing statuses for the given 4D_Number.
    """
    four_d = is_valid_4d(four_d)
    if not four_d:
        return False  # Invalid 4D_Number, cannot have overlapping status
    
    for row in records_parade:
        if is_valid_4d(row.get("4d_number", "")) == four_d:
            start_date = row.get("start_date_ddmmyyyy", "")
            end_date = row.get("end_date_ddmmyyyy", "")
            
            try:
                existing_start = datetime.strptime(start_date, "%d%m%Y")
                existing_end = datetime.strptime(end_date, "%d%m%Y")
                # Check for overlap
                if (new_start <= existing_end) and (new_end >= existing_start):
                    return True
            except ValueError:
                logger.warning(f"Invalid date format in Parade_State for {four_d}: {start_date} - {end_date}")
                continue
    return False

# ------------------------------------------------------------------------------
# 4) Streamlit Layout
# ------------------------------------------------------------------------------

st.title("Training & Parade Management App")

# ------------------------------------------------------------------------------
# 5) Sidebar: Configuration and Logout
# ------------------------------------------------------------------------------

st.sidebar.header("Configuration")
logout()

# Company selection: Dropdown to select one of the accessible companies
selected_company = st.sidebar.selectbox(
    "Select Company",
    options=st.session_state.user_companies
)

# Load the selected company's sheets
worksheets = get_sheets(selected_company)
if not worksheets:
    st.error("Failed to load the selected company's spreadsheets. Please check the logs for more details.")
    st.stop()

# Assign worksheets
SHEET_NOMINAL = worksheets["nominal"]
SHEET_PARADE = worksheets["parade"]  # Correct variable name
SHEET_CONDUCTS = worksheets["conducts"]

# ------------------------------------------------------------------------------
# 6) Session State: We store data so it's not lost on each run
# ------------------------------------------------------------------------------

# Conduct Session State
if "conduct_date" not in st.session_state:
    st.session_state.conduct_date = ""
if "conduct_platoon" not in st.session_state:
    st.session_state.conduct_platoon = 1  # Initialize as integer
if "conduct_name" not in st.session_state:
    st.session_state.conduct_name = ""
if "conduct_table" not in st.session_state:
    st.session_state.conduct_table = []
if "conduct_pointers_observation" not in st.session_state:
    st.session_state.conduct_pointers_observation = ""
if "conduct_pointers_reflection" not in st.session_state:
    st.session_state.conduct_pointers_reflection = ""
if "conduct_pointers_recommendation" not in st.session_state:
    st.session_state.conduct_pointers_recommendation = ""
# Removed manual 'conduct_submitted_by' since it will be auto-assigned

# Parade Session State
if "parade_platoon" not in st.session_state:
    st.session_state.parade_platoon = 1  # Initialize as integer
if "parade_table" not in st.session_state:
    st.session_state.parade_table = []
# Removed manual 'parade_submitted_by' since it will be auto-assigned

# Conduct Update Session State
if "update_conduct_selected" not in st.session_state:
    st.session_state.update_conduct_selected = None
if "update_conduct_platoon" not in st.session_state:
    st.session_state.update_conduct_platoon = 1  # Initialize as integer (Platoon 1-4)
if "update_conduct_pointers_observation" not in st.session_state:
    st.session_state.update_conduct_pointers_observation = ""
if "update_conduct_pointers_reflection" not in st.session_state:
    st.session_state.update_conduct_pointers_reflection = ""
if "update_conduct_pointers_recommendation" not in st.session_state:
    st.session_state.update_conduct_pointers_recommendation = ""
if "update_conduct_table" not in st.session_state:
    st.session_state.update_conduct_table = []

# ------------------------------------------------------------------------------
# 7) Feature Selection
# ------------------------------------------------------------------------------

feature = st.sidebar.selectbox(
    "Select Feature",
    ["Add Conduct", "Update Conduct", "Update Parade", "Queries", "Overall View", "Generate WhatsApp Message"]
)

# ------------------------------------------------------------------------------
# 8) Feature A: Add Conduct (table-based On-Status approach)
# ------------------------------------------------------------------------------

if feature == "Add Conduct":
    st.header("Add Conduct - Table-Based On-Status")

    # (a) Basic Inputs
    st.session_state.conduct_date = st.text_input(
        "Date (DDMMYYYY)",
        value=st.session_state.conduct_date
    )
    st.session_state.conduct_platoon = st.selectbox(
        "Your Platoon",
        options=[1, 2, 3, 4],
        format_func=lambda x: str(x)
    )
    st.session_state.conduct_name = st.text_input(
        "Conduct Name (e.g. IPPT)",
        value=st.session_state.conduct_name
    )

    # Separate inputs for Pointers: Observation, Reflection, Recommendation
    st.subheader("Pointers (ORR, Observation, Reflection)")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state.conduct_pointers_observation = st.text_input(
            "Observation",
            value=st.session_state.conduct_pointers_observation
        )
    with col2:
        st.session_state.conduct_pointers_reflection = st.text_input(
            "Reflection",
            value=st.session_state.conduct_pointers_reflection
        )
    with col3:
        st.session_state.conduct_pointers_recommendation = st.text_input(
            "Recommendation",
            value=st.session_state.conduct_pointers_recommendation
        )

    # Removed manual 'Submitted By' input
    submitted_by = st.session_state.username  # Automatically assign the username

    # (b) "Load On-Status"
    if st.button("Load On-Status"):
        date_str = st.session_state.conduct_date.strip()
        platoon = str(st.session_state.conduct_platoon).strip()

        if not date_str or not platoon:
            st.error("Please enter both Date and Platoon.")
            st.stop()

        # Validate date format
        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format (use DDMMYYYY).")
            st.stop()

        # Fetch records without caching
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        # Build conduct table with all personnel, marking 'Is_Outlier' based on status
        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)

        # Store in session
        st.session_state.conduct_table = conduct_data
        st.success(f"Loaded {len(conduct_data)} personnel for Platoon {platoon} ({date_obj.strftime('%d%m%Y')}).")
        logger.info(f"Loaded conduct personnel for Platoon {platoon} on {date_obj.strftime('%d%m%Y')} in company '{selected_company}' by user '{submitted_by}'.")

    # (c) Data Editor (allow new rows) - ALWAYS show, so you can finalize even with zero outliers
    if st.session_state.conduct_table:
        st.write("Toggle 'Is_Outlier' if not participating, or add new rows for extra people.")
        edited_data = st.data_editor(
            st.session_state.conduct_table,
            num_rows="dynamic",
            use_container_width=True
        )
    else:
        edited_data = st.data_editor(
            [],
            num_rows="dynamic",
            use_container_width=True
        )

    # (d) Finalize Conduct
    if st.button("Finalize Conduct"):
        date_str = st.session_state.conduct_date.strip()
        platoon = str(st.session_state.conduct_platoon).strip()
        cname = st.session_state.conduct_name.strip()
        observation = st.session_state.conduct_pointers_observation.strip()
        reflection = st.session_state.conduct_pointers_reflection.strip()
        recommendation = st.session_state.conduct_pointers_recommendation.strip()
        # submitted_by is already assigned

        if not date_str or not platoon or not cname:
            st.error("Please fill all fields (Date, Platoon, Conduct Name) first.")
            st.stop()

        # Validate date
        try:
            datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format.")
            st.stop()

        # Combine Pointers in the specified format
        pointers = f"Observation :\n{observation}\nReflection :\n{reflection}\nRecommendation :\n{recommendation}"

        # Fetch records without caching
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        # We'll figure out who is outlier + who is new to Nominal_Roll
        existing_4ds = {row.get("4d_number", "").strip().upper() for row in records_nominal}

        new_people = []
        all_outliers = []

        for row in edited_data:
            four_d = is_valid_4d(row.get("4D_Number", ""))
            name_ = ensure_str(row.get("Name", ""))
            rank_ = ensure_str(row.get("Rank", ""))  # Retrieve Rank
            is_outlier = row.get("Is_Outlier", False)
            status_desc = ensure_str(row.get("StatusDesc", ""))  # Get StatusDesc

            # Validate 4D_Number format
            if not four_d:
                st.error(f"Invalid 4D_Number format: {row.get('4D_Number', '')}. Skipping.")
                logger.error(f"Invalid 4D_Number format: {row.get('4D_Number', '')}.")
                continue

            # If user left 4D blank, skip
            if four_d:
                # If not in Nominal_Roll => new person
                if four_d not in existing_4ds:
                    if not name_:
                        st.error(f"Name is required for new 4D_Number: {four_d}. Skipping.")
                        logger.error(f"Name missing for new 4D_Number: {four_d}.")
                        continue
                    if not rank_:
                        st.error(f"Rank is required for new 4D_Number: {four_d}. Skipping.")
                        logger.error(f"Rank missing for new 4D_Number: {four_d}.")
                        continue
                    new_people.append((rank_, name_, four_d, platoon))
                    logger.info(f"Adding new person: Rank={rank_}, Name={name_}, 4D_Number={four_d}, Platoon={platoon} in company '{selected_company}' by user '{submitted_by}'.")

                # If is_outlier, we'll add to outliers list with StatusDesc
                if is_outlier:
                    if status_desc:
                        all_outliers.append(f"{four_d} ({status_desc})")
                    else:
                        all_outliers.append(f"{four_d}")

        # Insert new people into Nominal_Roll
        for (rank, nm, fd, p_) in new_people:
            formatted_fd = ensure_date_str(fd)
            SHEET_NOMINAL.append_row([rank, nm, formatted_fd, p_, 14, ""])  # Initialize leaves
            logger.info(f"Added new person to Nominal_Roll: Rank={rank}, Name={nm}, 4D_Number={formatted_fd}, Platoon={p_} in company '{selected_company}' by user '{submitted_by}'.")

        # Now recalc total strength for all platoons to calculate P/T Alpha
        total_strength_platoons = {}
        for plt in [1, 2, 3, 4]:
            strength = get_company_strength(str(plt), records_nominal)
            total_strength_platoons[plt] = strength

        # Prepare P/T PLT1 to PLT4
        pt_plts = ['0/0', '0/0', '0/0', '0/0']

        # Calculate participating for the selected platoon
        participating = 0
        for row in edited_data:
            if not row.get('Is_Outlier', False):
                participating += 1

        # Set participation for the selected platoon
        pt_plts[int(platoon)-1] = f"{participating}/{total_strength_platoons[int(platoon)]}"

        # Calculate P/T Alpha as the sum of participations across all platoons / sum of total strengths
        x_total = 0
        for pt in pt_plts:
            x = int(pt.split('/')[0]) if '/' in pt and pt.split('/')[0].isdigit() else 0
            x_total += x
        y_total = sum(total_strength_platoons.values())
        pt_alpha = f"{x_total}/{y_total}"

        # Append row to Conducts with Submitted By and Pointers
        formatted_date_str = ensure_date_str(date_str)
        SHEET_CONDUCTS.append_row([
            formatted_date_str,
            cname,
            pt_plts[0],  # P/T PLT1
            pt_plts[1],  # P/T PLT2
            pt_plts[2],  # P/T PLT3
            pt_plts[3],  # P/T PLT4
            pt_alpha,     # P/T Alpha as "x/y"
            ", ".join(all_outliers) if all_outliers else "None",
            pointers,
            submitted_by  # Automatically assigned
        ])
        logger.info(f"Appended Conduct: {formatted_date_str}, {cname}, P/T PLT1: {pt_plts[0]}, P/T PLT2: {pt_plts[1]}, P/T PLT3: {pt_plts[2]}, P/T PLT4: {pt_plts[3]}, P/T Alpha: {pt_alpha}, Outliers: {', '.join(all_outliers) if all_outliers else 'None'}, Pointers: {pointers}, Submitted_By: {submitted_by} in company '{selected_company}'.")

        # Find the row number of the newly appended conduct
        try:
            conduct_cell = SHEET_CONDUCTS.find(cname, in_column=2)  # Conduct_Name is column B
            if conduct_cell:
                conduct_row = conduct_cell.row
            else:
                st.error("Failed to locate the newly added conduct in the sheet.")
                logger.error(f"Failed to locate the newly added conduct '{cname}' in the sheet.")
                st.stop()
        except Exception as e:
            st.error(f"Error locating Conduct in the sheet: {e}")
            logger.error(f"Exception while locating Conduct '{cname}': {e}")
            st.stop()

        # Update P/T Alpha in the sheet
        try:
            SHEET_CONDUCTS.update_cell(conduct_row, 7, pt_alpha)  # P/T Alpha is column 7
            logger.info(f"Updated P/T Alpha to {pt_alpha} for conduct '{cname}' in company '{selected_company}'.")
        except Exception as e:
            st.error(f"Error updating P/T Alpha: {e}")
            logger.error(f"Exception while updating P/T Alpha for conduct '{cname}': {e}")
            st.stop()

        st.success(
            f"Conduct Finalized!\n\n"
            f"Date: {formatted_date_str}\n"
            f"Conduct Name: {cname}\n"
            f"P/T PLT1: {pt_plts[0]}\n"
            f"P/T PLT2: {pt_plts[1]}\n"
            f"P/T PLT3: {pt_plts[2]}\n"
            f"P/T PLT4: {pt_plts[3]}\n"
            f"P/T Alpha: {pt_alpha}\n"
            f"Outliers: {', '.join(all_outliers) if all_outliers else 'None'}\n"
            f"Pointers:\n{pointers if pointers else 'None'}\n"
            f"Submitted By: {submitted_by}"
        )

        # Clear session state variables
        st.session_state.conduct_date = ""
        st.session_state.conduct_platoon = 1
        st.session_state.conduct_name = ""
        st.session_state.conduct_table = []
        st.session_state.conduct_pointers_observation = ""
        st.session_state.conduct_pointers_reflection = ""
        st.session_state.conduct_pointers_recommendation = ""
        # 'Submitted By' is auto-assigned, no need to clear session_state variables

        # **Clear Cached Data to Reflect Updates**
        # Removed caching, so no need to clear cache

# ------------------------------------------------------------------------------
# 9) Feature B: Update Conduct
# ------------------------------------------------------------------------------

elif feature == "Update Conduct":
    st.header("Update Conduct")

    # (a) Select Conduct to Update
    records_conducts = get_conduct_records(selected_company, SHEET_CONDUCTS)
    conduct_names = [f"{row['date']} - {row['conduct_name']}" for row in records_conducts]
    
    if not conduct_names:
        st.warning("No Conducts available to update.")
        st.stop()
    
    selected_conduct = st.selectbox(
        "Select Conduct to Update",
        options=conduct_names,
        key="update_conduct_selected"
    )

    if not selected_conduct:
        st.error("Please select a conduct to update.")
        st.stop()

    # Find the selected conduct record
    conduct_index = conduct_names.index(selected_conduct) if selected_conduct in conduct_names else -1
    if conduct_index == -1:
        st.error("Selected conduct not found.")
        st.stop()

    conduct_record = records_conducts[conduct_index]

    # (b) Select Platoon to Update
    st.subheader("Select Platoon to Update")
    selected_platoon = st.selectbox(
        "Select Platoon",
        options=[1, 2, 3, 4],
        format_func=lambda x: f"Platoon {x}",
        key="update_conduct_platoon_select"
    )

    # (c) Input for Pointers with guidance
    # Fetch existing pointers and split into Observation, Reflection, Recommendation
    existing_pointers = conduct_record.get('pointers', '')
    if existing_pointers:
        # Use regular expressions to extract the text after each label
        observation_match = re.search(r'Observation\s*:\s*(.*)', existing_pointers, re.IGNORECASE)
        reflection_match = re.search(r'Reflection\s*:\s*(.*)', existing_pointers, re.IGNORECASE)
        recommendation_match = re.search(r'Recommendation\s*:\s*(.*)', existing_pointers, re.IGNORECASE)

        observation_existing = observation_match.group(1).strip() if observation_match else ""
        reflection_existing = reflection_match.group(1).strip() if reflection_match else ""
        recommendation_existing = recommendation_match.group(1).strip() if recommendation_match else ""
    else:
        observation_existing = ""
        reflection_existing = ""
        recommendation_existing = ""

    st.subheader("Update Pointers (ORR, Observation, Reflection)")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state.update_conduct_pointers_observation = st.text_input(
            "Observation",
            value=observation_existing
        )
    with col2:
        st.session_state.update_conduct_pointers_reflection = st.text_input(
            "Reflection",
            value=reflection_existing
        )
    with col3:
        st.session_state.update_conduct_pointers_recommendation = st.text_input(
            "Recommendation",
            value=recommendation_existing
        )

    # (d) "Load On-Status for the selected Conduct and Platoon"
    if st.button("Load On-Status for Update"):
        platoon = str(selected_platoon).strip()
        date_str = conduct_record['date']
        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format in selected Conduct.")
            st.stop()

        # Fetch records without caching
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        # Build conduct table with all personnel, marking 'Is_Outlier' based on status
        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)

        # Store in session
        st.session_state.update_conduct_table = conduct_data
        st.success(f"Loaded {len(conduct_data)} personnel for Platoon {platoon} from Conduct '{selected_conduct}'.")
        logger.info(f"Loaded conduct personnel for Platoon {platoon} from Conduct '{selected_conduct}' in company '{selected_company}' by user '{st.session_state.username}'.")

    # (e) Data Editor for Conduct Update
    if "update_conduct_table" in st.session_state and st.session_state.update_conduct_table:
        st.subheader(f"Edit Conduct Data for Platoon {selected_platoon}")
        st.write("Toggle 'Is_Outlier' if not participating, or add new rows for extra people.")
        edited_data = st.data_editor(
            st.session_state.update_conduct_table,
            num_rows="dynamic",
            use_container_width=True
        )
    else:
        edited_data = None

    # (f) Finalize Conduct Update
    if st.button("Update Conduct Data") and edited_data is not None:
        rows_updated = 0
        platoon = str(selected_platoon).strip()
        pt_field = f"P/T PLT{platoon}"
        # Calculate new Participating and Total based on edited_data
        new_participating = sum([1 for row in edited_data if not row.get('Is_Outlier', False)])
        new_total = len(edited_data)
        new_outliers = []
        observation = st.session_state.update_conduct_pointers_observation.strip()
        reflection = st.session_state.update_conduct_pointers_reflection.strip()
        recommendation = st.session_state.update_conduct_pointers_recommendation.strip()

        # Combine Pointers in the specified format
        new_pointers = f"Observation :\n{observation}\nReflection :\n{reflection}\nRecommendation :\n{recommendation}"

        for row in edited_data:
            four_d = is_valid_4d(row.get("4D_Number", ""))
            status_desc = ensure_str(row.get("StatusDesc", ""))
            if row.get("Is_Outlier", False):
                if status_desc:
                    new_outliers.append(f"{four_d} ({status_desc})")
                else:
                    new_outliers.append(f"{four_d}")

        # Fetch existing outliers and append new ones
        existing_outliers = ensure_str(conduct_record.get('outliers', ''))
        if existing_outliers.lower() != 'none' and existing_outliers:
            # Remove any existing outliers from the selected platoon to prevent duplication
            # Assuming outliers are in the format "4DXXXX (Reason)"
            pattern = re.compile(r'4D\d{3,4}\s*\([^)]*\)')
            existing_outliers_list = pattern.findall(existing_outliers)
            # Append new outliers
            updated_outliers = ", ".join(existing_outliers_list + new_outliers)
        else:
            updated_outliers = ", ".join(new_outliers) if new_outliers else "None"

        # Update P/T PLTx
        new_pt_value = f"{new_participating}/{new_total}"  # Removed leading single quote

        # Find the row number in the sheet
        try:
            conduct_name = selected_conduct.split(" - ")[1]
            cell = SHEET_CONDUCTS.find(conduct_name, in_column=2)  # Conduct_Name is column B
            if not cell:
                st.error("Conduct not found in the sheet.")
                logger.error(f"Conduct '{selected_conduct}' not found in the sheet.")
                st.stop()
            row_number = cell.row
        except Exception as e:
            st.error(f"Error locating Conduct in the sheet: {e}")
            logger.error(f"Exception while locating Conduct '{selected_conduct}': {e}")
            st.stop()

        try:
            SHEET_CONDUCTS.update_cell(row_number, 3 + int(platoon) - 1, new_pt_value)  # P/T PLT1 is column 3
            logger.info(f"Updated {pt_field} to {new_pt_value} for conduct '{selected_conduct}' in company '{selected_company}' by user '{st.session_state.username}'.")
        except Exception as e:
            st.error(f"Error updating {pt_field}: {e}")
            logger.error(f"Exception while updating {pt_field}: {e}")
            st.stop()

        # Update Outliers
        try:
            SHEET_CONDUCTS.update_cell(row_number, 8, updated_outliers if updated_outliers else "None")  # Outliers is column 8
            logger.info(f"Updated Outliers to '{updated_outliers}' for conduct '{selected_conduct}' in company '{selected_company}' by user '{st.session_state.username}'.")
        except Exception as e:
            st.error(f"Error updating Outliers: {e}")
            logger.error(f"Exception while updating Outliers: {e}")
            st.stop()

        # Update Pointers
        if new_pointers:
            try:
                SHEET_CONDUCTS.update_cell(row_number, 9, new_pointers)  # Pointers is column 9
                logger.info(f"Updated Pointers to '{new_pointers}' for conduct '{selected_conduct}' in company '{selected_company}' by user '{st.session_state.username}'.")
            except Exception as e:
                st.error(f"Error updating Pointers: {e}")
                logger.error(f"Exception while updating Pointers for conduct '{selected_conduct}': {e}")
                st.stop()

        # Calculate P/T Alpha as the sum of P/T PLT1 to P/T PLT4
        try:
            pt1 = SHEET_CONDUCTS.cell(row_number, 3).value  # P/T PLT1
            pt2 = SHEET_CONDUCTS.cell(row_number, 4).value  # P/T PLT2
            pt3 = SHEET_CONDUCTS.cell(row_number, 5).value  # P/T PLT3
            pt4 = SHEET_CONDUCTS.cell(row_number, 6).value  # P/T PLT4

            # Extract participating numbers
            pt1_part = int(pt1.split('/')[0]) if '/' in pt1 and pt1.split('/')[0].isdigit() else 0
            pt2_part = int(pt2.split('/')[0]) if '/' in pt2 and pt2.split('/')[0].isdigit() else 0
            pt3_part = int(pt3.split('/')[0]) if '/' in pt3 and pt3.split('/')[0].isdigit() else 0
            pt4_part = int(pt4.split('/')[0]) if '/' in pt4 and pt4.split('/')[0].isdigit() else 0

            x_total = pt1_part + pt2_part + pt3_part + pt4_part
            y_total = sum([int(p.split('/')[1]) if '/' in p and p.split('/')[1].isdigit() else 0 for p in [pt1, pt2, pt3, pt4]])

            pt_alpha = f"{x_total}/{y_total}"

            # Update P/T Alpha in the sheet (column 7)
            SHEET_CONDUCTS.update_cell(row_number, 7, pt_alpha)
            logger.info(f"Updated P/T Alpha to {pt_alpha} for conduct '{selected_conduct}' in company '{selected_company}' by user '{st.session_state.username}'.")
        except Exception as e:
            st.error(f"Error calculating/updating P/T Alpha: {e}")
            logger.error(f"Exception while calculating/updating P/T Alpha for conduct '{selected_conduct}': {e}")
            st.stop()

        st.success(f"Conduct '{selected_conduct}' updated successfully.")
        logger.info(f"Conduct '{selected_conduct}' updated successfully in company '{selected_company}' by user '{st.session_state.username}'.")

        # **Reset session_state variables**
        st.session_state.update_conduct_pointers_observation = ""
        st.session_state.update_conduct_pointers_reflection = ""
        st.session_state.update_conduct_pointers_recommendation = ""

        # **Clear Cached Data to Reflect Updates**
        # Removed caching, so no need to clear cache

# ------------------------------------------------------------------------------
# 10) Feature C: Update Parade
# ------------------------------------------------------------------------------

elif feature == "Update Parade":
    st.header("Update Parade State")

    # (a) Input for platoon
    # **Modification: Added "Coy HQ" to the platoon options below**
    st.session_state.parade_platoon = st.selectbox(
        "Platoon for Parade Update:",
        options=[1, 2, 3, 4, "Coy HQ"],  # Added "Coy HQ" here
        format_func=lambda x: str(x)
    )

    # Removed manual 'Submitted By' input
    submitted_by = st.session_state.username  # Automatically assign the username

    if st.button("Load Personnel"):
        platoon = str(st.session_state.parade_platoon).strip()
        if not platoon:
            st.error("Please select a valid platoon.")
            st.stop()

        # Fetch records without caching
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        data = get_company_personnel(platoon, records_nominal, records_parade)
        st.session_state.parade_table = data
        st.info(f"Loaded {len(data)} personnel for Platoon {platoon} in company '{selected_company}'.")
        logger.info(f"Loaded personnel for Platoon {platoon} in company '{selected_company}' by user '{submitted_by}'.")

        # ------------------------------------------------------------------------------
        # Display Current Parade Statuses for the Platoon
        # ------------------------------------------------------------------------------
        current_statuses = [
            row for row in records_parade
            if normalize_name(row.get('platoon', '')) == normalize_name(platoon)
        ]

        if current_statuses:
            st.subheader("Current Parade Status")
            # Format the current statuses for better readability
            formatted_statuses = []
            for status in current_statuses:
                formatted_statuses.append({
                    "4D_Number": status.get("4d_number", ""),
                    "Name": find_name_by_4d(status.get("4d_number", ""), records_nominal),
                    "Status": status.get("status", ""),
                    "Start_Date": status.get("start_date_ddmmyyyy", ""),
                    "End_Date": status.get("end_date_ddmmyyyy", "")
                })
            # Display as a table
            #st.table(formatted_statuses)
            logger.info(f"Displayed current parade statuses for platoon {platoon} in company '{selected_company}' by user '{submitted_by}'.")

    # (b) Show data editor if we have data
    if st.session_state.parade_table:
        st.subheader("Edit Parade Data, Then Click 'Update'")
        st.write("Fill in 'Status', 'Start_Date (DDMMYYYY)', 'End_Date (DDMMYYYY)'")
        edited_data = st.data_editor(
            st.session_state.parade_table,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True  # Hide the default index
        )
    else:
        edited_data = None

    # (c) Finalize Parade Update
    if st.button("Update Parade State") and edited_data is not None:
        rows_updated = 0
        platoon = str(st.session_state.parade_platoon).strip()
        # submitted_by is already assigned

        # Fetch records without caching
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        for idx, row in enumerate(edited_data):
            four_d = is_valid_4d(row.get("4D_Number", ""))
            status_val = ensure_str(row.get("Status", "")).strip()
            start_val = ensure_str(row.get("Start_Date", "")).strip()
            end_val = ensure_str(row.get("End_Date", "")).strip()

            # Retrieve the corresponding row number from the original parade_table
            parade_entry = st.session_state.parade_table[idx]
            row_num = parade_entry.get('_row_num')  # Get row number for updating

            # Validate 4D_Number format
            if not four_d:
                st.error(f"Invalid 4D_Number format: {row.get('4D_Number', '')}. Skipping.")
                logger.error(f"Invalid 4D_Number format: {row.get('4D_Number', '')}.")
                continue

            # Determine if the row was changed by comparing with the original
            original_entry = st.session_state.parade_table[idx]
            is_changed = False
            if original_entry['_row_num'] is not None:
                # Compare relevant fields
                if (
                    row.get('Status', '') != original_entry.get('Status', '') or
                    row.get('Start_Date', '') != original_entry.get('Start_Date', '') or
                    row.get('End_Date', '') != original_entry.get('End_Date', '')
                ):
                    is_changed = True

            if not status_val and row_num:
                # If status fields are cleared and row_num exists, consider deleting the status
                try:
                    SHEET_PARADE.delete_rows(row_num)
                    logger.info(f"Deleted Parade_State row {row_num} for {four_d} in company '{selected_company}' by user '{submitted_by}'.")
                    rows_updated += 1
                    continue
                except Exception as e:
                    st.error(f"Error deleting row for {four_d}: {e}. Skipping.")
                    logger.error(f"Exception while deleting row for {four_d}: {e}.")
                    continue

            if not status_val or not start_val or not end_val:
                # Skip entries with missing required fields
                logger.error(f"Missing fields for {four_d} in company '{selected_company}'.")
                continue

            # Ensure dates are properly formatted
            formatted_start_val = ensure_date_str(start_val)
            formatted_end_val = ensure_date_str(end_val)

            # Validate date formats
            try:
                start_dt = datetime.strptime(formatted_start_val, "%d%m%Y")
                end_dt = datetime.strptime(formatted_end_val, "%d%m%Y")
                if end_dt < start_dt:
                    st.error(f"End date is before start date for {four_d}, skipping.")
                    logger.error(f"End date before start date for {four_d} in company '{selected_company}'.")
                    continue
            except ValueError:
                st.error(f"Invalid date(s) for {four_d}, skipping.")
                logger.error(f"Invalid date format for {four_d}: Start={formatted_start_val}, End={formatted_end_val} in company '{selected_company}'.")
                continue

            if status_val.lower() == "leave":
                # Calculate number of leaves used
                if formatted_start_val != formatted_end_val:
                    dates_str = f"{formatted_start_val}-{formatted_end_val}"
                else:
                    dates_str = formatted_start_val
                leaves_used = calculate_leaves_used(dates_str)
                if leaves_used <= 0:
                    st.error(f"Invalid leave duration for {four_d}, skipping.")
                    logger.error(f"Invalid leave duration for {four_d}: {dates_str} in company '{selected_company}'.")
                    continue

                # Check for overlapping statuses
                if has_overlapping_status(four_d, start_dt, end_dt, records_parade):
                    logger.error(f"Leave dates overlap for {four_d}: {dates_str} in company '{selected_company}'.")
                    continue

                # Fetch current leaves and dates taken
                try:
                    nominal_record = SHEET_NOMINAL.find(four_d, in_column=3)  # Assuming 4D_Number is column C
                    if nominal_record:
                        current_leaves_left = SHEET_NOMINAL.cell(nominal_record.row, 5).value  # Number of Leaves Left is column E
                        try:
                            current_leaves_left = int(current_leaves_left)
                        except ValueError:
                            current_leaves_left = 14  # Default if invalid
                            logger.warning(f"Invalid 'Number of Leaves Left' for {four_d}. Resetting to 14 in company '{selected_company}'.")

                        if leaves_used > current_leaves_left:
                            st.error(f"{four_d} does not have enough leaves left. Available: {current_leaves_left}, Requested: {leaves_used}. Skipping.")
                            logger.error(f"{four_d} insufficient leaves. Available: {current_leaves_left}, Requested: {leaves_used} in company '{selected_company}'.")
                            continue

                        # Update leaves left
                        new_leaves_left = current_leaves_left - leaves_used
                        SHEET_NOMINAL.update_cell(nominal_record.row, 5, new_leaves_left)
                        logger.info(f"Updated 'Number of Leaves Left' for {four_d}: {new_leaves_left} in company '{selected_company}' by user '{submitted_by}'.")

                        # Update Dates Taken
                        existing_dates = SHEET_NOMINAL.cell(nominal_record.row, 6).value  # Dates Taken is column F
                        new_dates_entry = dates_str
                        if existing_dates:
                            updated_dates = existing_dates + f",{new_dates_entry}"
                        else:
                            updated_dates = new_dates_entry
                        SHEET_NOMINAL.update_cell(nominal_record.row, 6, updated_dates)
                        logger.info(f"Updated 'Dates Taken' for {four_d}: {updated_dates} in company '{selected_company}' by user '{submitted_by}'.")
                    else:
                        st.error(f"{four_d} not found in Nominal_Roll. Skipping.")
                        logger.error(f"{four_d} not found in Nominal_Roll in company '{selected_company}'.")
                        continue
                except Exception as e:
                    st.error(f"Error updating leaves for {four_d}: {e}. Skipping.")
                    logger.error(f"Exception while updating leaves for {four_d}: {e} in company '{selected_company}'.")
                    continue

            # Update the existing Parade_State row instead of appending
            if row_num:
                # Find the column numbers based on header
                header = SHEET_PARADE.row_values(1)
                header = [h.strip().lower() for h in header]
                try:
                    status_col = header.index("status") + 1
                    start_date_col = header.index("start_date_ddmmyyyy") + 1
                    end_date_col = header.index("end_date_ddmmyyyy") + 1
                    submitted_by_col = header.index("submitted_by") + 1 if "submitted_by" in header else None
                except ValueError as ve:
                    st.error(f"Required column missing in Parade_State: {ve}.")
                    logger.error(f"Required column missing in Parade_State: {ve} in company '{selected_company}'.")
                    continue

                SHEET_PARADE.update_cell(row_num, status_col, status_val)  # Corrected SHEET_PARDE to SHEET_PARADE
                SHEET_PARADE.update_cell(row_num, start_date_col, formatted_start_val)  # Corrected SHEET_PARDE to SHEET_PARADE
                SHEET_PARADE.update_cell(row_num, end_date_col, formatted_end_val)  # Corrected SHEET_PARDE to SHEET_PARADE

                # Update 'Submitted_By' only if the row was changed
                if submitted_by_col and is_changed:
                    SHEET_PARADE.update_cell(row_num, submitted_by_col, submitted_by)

                rows_updated += 1
            else:
                # If no existing row, append as a new entry
                SHEET_PARADE.append_row([platoon, four_d, status_val, formatted_start_val, formatted_end_val, submitted_by])  # Corrected SHEET_PARDE to SHEET_PARADE
                logger.info(f"Appended Parade_State for {four_d}: Status={status_val}, Start={formatted_start_val}, End={formatted_end_val}, Submitted_By={submitted_by} in company '{selected_company}' by user '{submitted_by}'.")
                rows_updated += 1

        st.success(f"Parade State updated.")
        logger.info(f"Parade State updated for {rows_updated} row(s) for platoon {platoon} in company '{selected_company}' by user '{submitted_by}'.")

        # **Reset session_state variables**
        st.session_state.parade_platoon = 1
        st.session_state.parade_table = []
        # 'Submitted By' is auto-assigned, no need to clear

        # **Clear Cached Data to Reflect Updates**
        # Removed caching, so no need to clear cache

# ------------------------------------------------------------------------------
# 11) Feature D: Queries (Combined Query Person & Query Outliers)
# ------------------------------------------------------------------------------

elif feature == "Queries":
    st.header("Queries")

    # Create tabs for "Query Person" and "Query Outliers"
    query_tabs = st.tabs(["Query Person", "Query Outliers"])

    # ---------------------------
    # Tab 1: Query Person
    # ---------------------------
    with query_tabs[0]:
        st.subheader("Query All Statuses for a Person")

        four_d_input = st.text_input("Enter the 4D Number (e.g. 4D001)", key="query_person_4d")
        if st.button("Get Statuses", key="btn_query_person"):
            four_d_input_clean = is_valid_4d(four_d_input)
            if not four_d_input_clean:
                st.error(f"Invalid 4D_Number format: {four_d_input}. It should start with '4D' followed by digits.")
                logger.error(f"Invalid 4D_Number format: {four_d_input}.")
                st.stop()

            # Fetch records without caching
            records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
            records_parade = get_parade_records(selected_company, SHEET_PARADE)

            parade_data = records_parade
            # Filter rows for this 4D
            person_rows = [
                row for row in parade_data
                if row.get("4d_number", "").strip().upper() == four_d_input_clean
            ]

            if not person_rows:
                st.warning(f"No Parade_State records found for {four_d_input_clean}")
                logger.info(f"No Parade_State records found for {four_d_input_clean} in company '{selected_company}'.")
            else:
                # Sort by start date
                def parse_ddmmyyyy(d):
                    try:
                        return datetime.strptime(str(d), "%d%m%Y")
                    except ValueError:
                        return datetime.min

                person_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("start_date_ddmmyyyy", "")))

                # Enhance display by adding Name
                enhanced_rows = []
                for row in person_rows:
                    enhanced_rows.append({
                        "4D_Number": row.get("4d_number", ""),
                        "Name": find_name_by_4d(row.get("4d_number", ""), records_nominal),
                        "Status": row.get("status", ""),
                        "Start_Date": row.get("start_date_ddmmyyyy", ""),
                        "End_Date": row.get("end_date_ddmmyyyy", "")
                    })

                st.subheader(f"Statuses for {four_d_input_clean}")
                # Show as a table
                st.table(enhanced_rows)
                logger.info(f"Displayed statuses for {four_d_input_clean} in company '{selected_company}'.")

    # ---------------------------
    # Tab 2: Query Outliers
    # ---------------------------
    with query_tabs[1]:
        st.subheader("Query Outliers for a Specific Platoon & Conduct")

        # Input fields
        platoon_q = st.selectbox("Platoon", options=[1, 2, 3, 4], key="query_outliers_platoon")
        cond_q = st.text_input("Conduct Name", key="query_outliers_conduct")

        if st.button("Get Outliers", key="btn_query_outliers"):
            platoon_query = str(st.session_state.query_outliers_platoon).strip()
            conduct_query = ensure_str(cond_q)

            if not platoon_query or not conduct_query:
                st.error("Please enter both Platoon and Conduct Name.")
                st.stop()

            conduct_norm = normalize_name(conduct_query)

            # Fetch records without caching
            conducts_data = get_conduct_records(selected_company, SHEET_CONDUCTS)

            # Filter records matching both platoon and conduct name
            matched_records = [
                row for row in conducts_data
                if normalize_name(row.get('conduct_name', '')) == conduct_norm and
                   row.get(f'p/t plt{platoon_query}', '').split('/')[0].isdigit() and
                   int(row.get(f'p/t plt{platoon_query}', '').split('/')[0]) > 0
            ]

            if not matched_records:
                # Attempt fuzzy matching if no exact match found
                conduct_pairs = [
                    (normalize_name(row.get('conduct_name', '')))
                    for row in conducts_data
                    if row.get('conduct_name', '').strip()
                ]
                closest_matches = difflib.get_close_matches(conduct_norm, conduct_pairs, n=1, cutoff=0.6)
                if not closest_matches:
                    st.error("❌ **No similar conduct name found.**\n\nPlease check your input and try again.")
                    logger.error(f"No similar conduct name found for: {conduct_norm} in company '{selected_company}'.")
                    st.stop()
                matched_norm = closest_matches[0]
                # Retrieve the original names
                matched_records = [
                    row for row in conducts_data
                    if normalize_name(row.get('conduct_name', '')) == matched_norm
                ]
                if not matched_records:
                    st.error("❌ **No data found for the matched conduct.**")
                    logger.error(f"No data found for the matched conduct: {matched_norm} in company '{selected_company}'.")
                    st.stop()

            # Collect outliers from matched records for the specific platoon
            all_outliers = []
            for row in matched_records:
                p_t_field = f"P/T PLT{platoon_query}"
                p_t_value = row.get(p_t_field, '0/0')
                participating = 0
                try:
                    participating = int(p_t_value.split('/')[0]) if '/' in p_t_value and p_t_value.split('/')[0].isdigit() else 0
                except:
                    participating = 0
                if participating > 0:
                    outliers_value = row.get('outliers', '')
                    if isinstance(outliers_value, (int, float)):
                        outliers_str = str(outliers_value)
                    elif isinstance(outliers_value, str):
                        outliers_str = outliers_value
                    else:
                        outliers_str = ''

                    if outliers_str.lower() != 'none' and outliers_str.strip():
                        # Assuming outliers are specific to platoon, you might need to adjust based on actual data structure
                        outliers = [o.strip() for o in outliers_str.split(',') if o.strip()]
                        all_outliers.extend(outliers)

            if all_outliers:
                # Count frequency of each outlier
                outlier_freq = {}
                for o in all_outliers:
                    outlier_freq[o] = outlier_freq.get(o, 0) + 1
                # Sort outliers by frequency
                sorted_outliers = sorted(outlier_freq.items(), key=lambda x: x[1], reverse=True)
                # Prepare data for table
                outlier_table = [{"Outlier": o, "Frequency": c} for o, c in sorted_outliers]
                # Display as a table
                st.markdown(f"📈 **Outliers for '{conduct_query}' at Platoon {platoon_query} in company '{selected_company}':**")
                st.table(outlier_table)
                logger.info(f"Displayed outliers for '{conduct_query}' at Platoon {platoon_query} in company '{selected_company}' by user '{st.session_state.username}'.")
            else:
                st.info(f"✅ **No outliers recorded for '{conduct_query}' at Platoon {platoon_query}' in company '{selected_company}'.**")
                logger.info(f"No outliers recorded for '{conduct_query}' at Platoon {platoon_query}' in company '{selected_company}' by user '{st.session_state.username}'.")

# ------------------------------------------------------------------------------
# 12) Feature E: Overall View
# ------------------------------------------------------------------------------

elif feature == "Overall View":
    st.header("Overall View of All Conducts")

    # (a) Fetch all conducts
    conducts = get_conduct_records(selected_company, SHEET_CONDUCTS)
    if not conducts:
        st.info("No conducts available to display.")
    else:
        # (b) Convert to DataFrame
        df = pd.DataFrame(conducts)

        # (c) Convert 'date' to datetime objects for sorting
        def parse_date(date_str):
            try:
                return datetime.strptime(date_str, "%d%m%Y")
            except ValueError:
                return None

        df['Date'] = df['date'].apply(parse_date)

        # (d) Handle rows with invalid dates
        invalid_dates = df['Date'].isnull()
        if invalid_dates.any():
            st.warning(f"{invalid_dates.sum()} conduct(s) have invalid date formats and will appear at the bottom.")
            logger.warning(f"{invalid_dates.sum()} conduct(s) have invalid date formats in company '{selected_company}'.")

        # (e) Sort the DataFrame by 'Date' in descending order (latest first)
        df_sorted = df.sort_values(by='Date', ascending=False)

        # (f) Format the 'Date' column for better readability
        df_sorted['Date'] = df_sorted['Date'].dt.strftime("%d-%m-%Y")

        # (g) Select and rename columns for display
        display_columns = {
            'Date': 'Date',
            'conduct_name': 'Conduct Name',
            'p/t plt1': 'P/T PLT1',
            'p/t plt2': 'P/T PLT2',
            'p/t plt3': 'P/T PLT3',
            'p/t plt4': 'P/T PLT4',
            'p/t alpha': 'P/T Alpha',
            'outliers': 'Outliers',
            'pointers': 'Pointers',
            'submitted_by': 'Submitted By'
        }

        df_display = df_sorted.rename(columns=display_columns)[list(display_columns.values())]

        # ------------------------------------------------------------------------------
        # **Added: Filtering and Sorting Options**
        # ------------------------------------------------------------------------------

        st.subheader("Filter and Sort Conducts")

        # (h) Filtering Inputs
        with st.expander("🔍 Filter Conducts"):
            search_term = st.text_input("Search by Conduct Name or Date (DDMMYYYY):", value="", help="Enter a keyword to filter conducts by name or date.")
            sort_field = st.selectbox(
                "Sort By",
                options=["Date", "Conduct Name"],
                index=0
            )
            sort_order = st.radio(
                "Sort Order",
                options=["Ascending", "Descending"],
                index=1  # Default to Descending
            )

            # Apply filtering
            if search_term:
                search_term_upper = search_term.upper()
                df_display = df_display[
                    df_display['Conduct Name'].str.upper().str.contains(search_term_upper) |
                    df_display['Date'].str.contains(search_term)
                ]

            # Apply sorting
            ascending = True if sort_order == "Ascending" else False
            if sort_field == "Date":
                # Convert 'Date' back to datetime for accurate sorting
                df_display['Date_Sort'] = pd.to_datetime(df_display['Date'], format="%d-%m-%Y", errors='coerce')
                df_display = df_display.sort_values(by='Date_Sort', ascending=ascending)
                df_display = df_display.drop(columns=['Date_Sort'])
            elif sort_field == "Conduct Name":
                df_display = df_display.sort_values(by='Conduct Name', ascending=ascending)

        # ------------------------------------------------------------------------------
        # (h) Display the DataFrame with sorting enabled
        # ------------------------------------------------------------------------------

        st.subheader("All Conducts")
        st.dataframe(df_display, use_container_width=True)

        # ------------------------------------------------------------------------------
        # **Added: Individuals' Missed Conducts**
        # ------------------------------------------------------------------------------

        st.subheader("Individuals' Missed Conducts")

        # Create a dictionary to hold missed conducts per individual
        missed_conducts_dict = defaultdict(list)

        for conduct in conducts:
            conduct_name = ensure_str(conduct.get('conduct_name', ''))
            outliers_str = ensure_str(conduct.get('outliers', ''))
            if outliers_str.lower() == 'none' or not outliers_str.strip():
                continue  # No outliers to process

            # Split outliers by comma
            outliers = [o.strip() for o in outliers_str.split(',') if o.strip()]
            for outlier in outliers:
                # Extract 4D_Number (with or without reason)
                match = re.match(r'(4D\d{3,4})(?:\s*\(.*\))?', outlier, re.IGNORECASE)
                if match:
                    four_d = match.group(1).upper()
                    missed_conducts_dict[four_d].append(conduct_name)

        # Create a DataFrame from the dictionary
        missed_conducts_data = []
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        four_d_to_name = {row['4d_number']: row['name'] for row in records_nominal}

        for four_d, conducts_missed in missed_conducts_dict.items():
            name = four_d_to_name.get(four_d, "Unknown")
            missed_conducts_data.append({
                "4D_Number": four_d,
                "Name": name,
                "Missed Conducts Count": len(conducts_missed),
                "Missed Conducts": ", ".join(conducts_missed)
            })

        if missed_conducts_data:
            df_missed = pd.DataFrame(missed_conducts_data)
            # Sort from most to least missed
            df_missed = df_missed.sort_values(by="Missed Conducts Count", ascending=False)

            # Reset index for styling
            df_missed = df_missed.reset_index(drop=True)

            # Apply styling to bold top 3
            def highlight_top3(row):
                if row.name < 3:
                    return ['font-weight: bold'] * len(row)
                else:
                    return [''] * len(row)

            styled_df = df_missed.style.apply(highlight_top3, axis=1)

            st.subheader("Missed Conducts by Individuals (Most to Least)")
            st.dataframe(styled_df, use_container_width=True)
        else:
            st.info("✅ **No individuals have missed any conducts.**")
            logger.info(f"No missed conducts recorded in company '{selected_company}' by user '{st.session_state.username}'.")

        # ------------------------------------------------------------------------------
        # **End of Added Section**
        # ------------------------------------------------------------------------------

        logger.info(f"Displayed overall view of all conducts in company '{selected_company}' by user '{st.session_state.username}'.")

# ------------------------------------------------------------------------------
# 13) Feature F: Generate WhatsApp Message
# ------------------------------------------------------------------------------

elif feature == "Generate WhatsApp Message":

    # (a) Fetch parade records and nominal records
    records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
    records_parade = get_parade_records(selected_company, SHEET_PARADE)

    # **Fix Applied: Filter parade records to include only those applicable to today**
    today_date = datetime.today().date()
    filtered_parade = []
    for parade in records_parade:
        start_date_str = parade.get('start_date_ddmmyyyy', '')
        end_date_str = parade.get('end_date_ddmmyyyy', '')
        try:
            start_dt = datetime.strptime(start_date_str, "%d%m%Y").date()
            end_dt = datetime.strptime(end_date_str, "%d%m%Y").date()
            if start_dt <= today_date <= end_dt:
                filtered_parade.append(parade)
        except ValueError:
            logger.warning(f"Invalid date format for {parade.get('4d_number', '')}: {start_date_str} - {end_date_str}")
            continue
    records_parade = filtered_parade

    # (b) Calculate TOTAL STR
    total_strength = len(records_nominal)

    # (c) Calculate CURRENT STR (Total - number of MCs)
    # Modification: CURRENT STR = TOTAL STR - number of MCs
    mc_count = len([person for person in records_parade if person.get('status', '').strip().lower() == 'mc'])
    current_strength = total_strength - mc_count

    # (d) Categorize personnel
    mc_list = []
    statuses_list = []
    others_list = []

    for parade in records_parade:
        status = ensure_str(parade.get('status', '')).lower()
        four_d = parade.get('4d_number', '').strip().upper()
        rank = ""
        name = ""
        # Find the person's name and rank from nominal records
        for nominal in records_nominal:
            if nominal['4d_number'] == four_d:
                rank = nominal['rank']
                name = nominal['name']
                break
        if status == 'mc':
            # Calculate number of days
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                delta_days = (end_dt - start_dt).days + 1
                delta_str = f"{delta_days}D"
                if delta_days == 1:
                    details = f"{delta_str} MC ({start_dt.strftime('%d%m%y')})"
                else:
                    details = f"{delta_str} MC ({start_dt.strftime('%d%m%y')}-{end_dt.strftime('%d%m%y')})"
            except ValueError:
                details = f"MC (Invalid Dates)"
                logger.warning(f"Invalid dates for MC {four_d}: {parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}")
            mc_list.append({
                "Rank": rank,
                "Name": name,
                "Details": details
            })
        elif status in ['rib', 'ld'] or status.startswith('ex'):
            # Assuming 'rib' and 'ld' fall under STATUSES
            # Also, any status starting with 'ex' falls under STATUSES
            # You can adjust the conditions based on actual status values
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                delta_days = (end_dt - start_dt).days + 1
                delta_str = f"{delta_days}D"
                if delta_days == 1:
                    details = f"{delta_str} {parade.get('status', '').upper()} ({start_dt.strftime('%d%m%y')})"
                else:
                    details = f"{delta_str} {parade.get('status', '').upper()} ({start_dt.strftime('%d%m%y')}-{end_dt.strftime('%d%m%y')})"
            except ValueError:
                details = f"{parade.get('status', '').upper()} (Invalid Dates)"
                logger.warning(f"Invalid dates for Status {four_d}: {parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}")
            statuses_list.append({
                "Rank": rank,
                "Name": name,
                "Details": details
            })
        else:
            # For any other status, calculate days similarly
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                delta_days = (end_dt - start_dt).days + 1
                delta_str = f"{delta_days}D"
                if delta_days == 1:
                    details = f"{delta_str} {parade.get('status', '').upper()} ({start_dt.strftime('%d%m%y')})"
                else:
                    details = f"{delta_str} {parade.get('status', '').upper()} ({start_dt.strftime('%d%m%y')}-{end_dt.strftime('%d%m%y')})"
            except ValueError:
                details = f"{parade.get('status', '').upper()} (Invalid Dates)"
                logger.warning(f"Invalid dates for Status {four_d}: {parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}")
            others_list.append({
                "Rank": rank,
                "Name": name,
                "Details": details
            })

    # (f) Build the message
    today_str = datetime.today().strftime("%d%m%y")  # DDMMYY
    platoon = "1"  # Assuming platoon 1; adjust as necessary
    company_name = selected_company  # You can customize this if needed

    message_lines = []
    message_lines.append(f"{today_str} {platoon} SIR ({company_name.upper()}) PARADE STRENGTH\n")

    message_lines.append(f"TOTAL STR: {total_strength}")
    message_lines.append(f"CURRENT STR: {current_strength}\n")

    # MC Section
    message_lines.append(f"MC: {len(mc_list)}")
    for idx, person in enumerate(mc_list, start=1):
        message_lines.append(f"{idx}. {person['Rank']}/{person['Name']}")
        message_lines.append(f"* {person['Details']}")
    message_lines.append("")  # Empty line

    # STATUSES Section
    # Count how many STATUSES have "OOT" and "EX STAY IN"
    oot_count = sum(1 for s in statuses_list if re.search(r'\bOOT\b', s['Details'], re.IGNORECASE))
    ex_stay_in_count = sum(1 for s in statuses_list if re.search(r'\bEX STAY IN\b', s['Details'], re.IGNORECASE))
    total_statuses = len(statuses_list)
    message_lines.append(f"STATUSES: {total_statuses} [*{oot_count} OOTS ({ex_stay_in_count} EX STAY IN)]")
    for idx, person in enumerate(statuses_list, start=1):
        message_lines.append(f"{idx}. {person['Rank']}/{person['Name']}")
        message_lines.append(f"* {person['Details']}")
    message_lines.append("")  # Empty line

    # OTHERS Section
    message_lines.append(f"OTHERS: {len(others_list):02d}")
    for idx, person in enumerate(others_list, start=1):
        message_lines.append(f"{idx}. {person['Rank']}/{person['Name']}")
        message_lines.append(f"* {person['Details']}")

    # Combine all lines into a single message
    whatsapp_message = "\n".join(message_lines)

    st.code(whatsapp_message, language='text')

# ------------------------------------------------------------------------------
# 14) End of Features
# ------------------------------------------------------------------------------

# The code ends here.
