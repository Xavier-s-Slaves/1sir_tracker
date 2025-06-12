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
from typing import List, Dict, Optional
from zoneinfo import ZoneInfo  # type: ignore
from datetime import timedelta
# ------------------------------------------------------------------------------
# Setup Logging
# ------------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
TIMEZONE = ZoneInfo('Asia/Singapore')  # Replace with your desired timezone
# ------------------------------------------------------------------------------
# User Authentication Setup
# ------------------------------------------------------------------------------
# Define a simple user database
# In a production environment, consider using a secure method to handle user credentials
USER_DB_PATH = "users.json"
LEGEND_STATUS_PREFIXES = {
        "ol": "[OL]",   # Overseas Leave
        "ll": "[LL]",   # Local Leave
        "ml": "[ML]",   # Medical Leave
        "mc": "[MC]",   # Medical Course
        "ao": "[AO]",   # Attached Out
        "oil": "[OIL]", # Off in Lieu
        "ma": "[MA]",   # Medical Appointment
        "so": "[SO]",   # Stay Out
        "cl": "[CL]",   # Compassionate Leave
        "i/a": "[I/A]",   # Interview/Appt
        "awol": "[AWOL]",   # AWOL
        "hl": "[HL]",   # Hospitalisation Leave
        "others": "[Others]",   # Hospitalisation Leave
    }
def parse_existing_outliers(existing_outliers_str):
    """
    Splits on commas (top-level), extracts parentheses as 'status_desc',
    and treats the rest as the name (with optional leading '4Dxxx' stripped).
    """

    # If the string is just "none", return an empty dict.
    if existing_outliers_str.strip().lower() == "none":
        return {}

    def split_outliers(s):
        """
        Splits the string on commas that are NOT inside any parentheses (including nested).
        E.g. "ABC (1,2), DEF" => ["ABC (1,2)", "DEF"].
        """
        parts = []
        current = []
        depth = 0
        for char in s:
            if char == '(':
                depth += 1
            elif char == ')':
                depth -= 1

            # A comma at depth 0 means a new entry.
            if char == ',' and depth == 0:
                parts.append(''.join(current).strip())
                current = []
            else:
                current.append(char)

        # Add the final piece
        if current:
            parts.append(''.join(current).strip())

        return parts

    def extract_top_level_parentheses(chunk):
        """
        Extracts *all* top-level parenthetical groups from a string.
        Returns: (text_without_parentheses, combined_status_string)

        Example:
          "ABC (STUFF (INSIDE)) (ANOTHER)" -> ("ABC", "STUFF (INSIDE), ANOTHER")
        """
        status_parts = []
        result_chars = []
        depth = 0
        start_idx = None

        i = 0
        while i < len(chunk):
            c = chunk[i]
            if c == '(':
                if depth == 0:
                    # start of a top-level group
                    start_idx = i
                depth += 1
            elif c == ')':
                depth -= 1
                if depth == 0 and start_idx is not None:
                    group_text = chunk[start_idx+1 : i]
                    status_parts.append(group_text.strip())
                    start_idx = None
            i += 1

        # Build the "outside" text, ignoring the contents of top-level parentheses
        depth = 0
        i = 0
        while i < len(chunk):
            c = chunk[i]
            if c == '(':
                depth += 1
            if depth == 0:
                result_chars.append(c)
            if c == ')':
                depth -= 1
            i += 1

        remainder = ''.join(result_chars).strip()
        combined_status = ', '.join(status_parts)
        return remainder, combined_status

    parts = split_outliers(existing_outliers_str)
    outliers_dict = {}

    for part in parts:
        # 1) Extract parentheses => statuses
        remainder, status_desc = extract_top_level_parentheses(part)

        # 2) Optional: Strip out a leading "4Dxxxx" if present. Examples:
        #    4D1106 NG YONG ZHENG => remainder_of_name = "NG YONG ZHENG"
        #    We'll do a simple check: if it starts with "4D", then skip it plus any letters/digits up to a space.
        remainder = remainder.strip()
        # Use a regex like:  ^4D[0-9A-Za-z]+\s+(.*)
        # If it matches, we drop that "4Dxxxx" portion from the name.
        match_4d = re.match(r'^4D[0-9A-Za-z]+\s+(.*)$', remainder, flags=re.IGNORECASE)
        if match_4d:
            name_str = match_4d.group(1).strip()
        else:
            name_str = remainder  # no leading "4D..." found

        # Convert to lowercase key
        key = name_str.lower()

        outliers_dict[key] = {
            "original": name_str,
            "status_desc": status_desc.strip()
        }

    return outliers_dict
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
    st.rerun()

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
    "Wolf": "Wolf",
    "Wolf1": "Wolf1",
    "Bravo": "Bravo",
    "Viper(bmt)": "Viper(bmt)",
    "Viper": "Viper",
    "MSC": "MSC",
    "Pegasus": "Pegasus",
    "Scorpion": "Scorpion",
    "WOSpec": "WOSpec"  # Added MSC
}


def extract_attendance_data(edited_data):
    """
    Extracts attendance data from the edited conduct data.
    Returns a list of tuples containing (name, rank, is_present).
    """
    attendance_data = []
    for row in edited_data:
        name = row.get("Name", "").strip()
        rank = row.get("Rank", "").strip()
        is_present = not row.get("Is_Outlier", False)
        attendance_data.append((name, rank, is_present))
    return attendance_data
def parse_4d_number(num_str: str):
    """
    Parses a 4D number string.
    Assuming the string is like '1101' where:
    - First digit: platoon
    - Second digit: section
    - Last two digits: roll number
    """
    if len(num_str) < 4:
        return None, None, None
    platoon = num_str[2]
    section = num_str[3]
    roll = num_str[4:]
    return platoon, section, roll

def analyze_attendance(everything_data: list,
                       nominal_data: list,
                       conduct_header: str):
    """
    Same signature as before, but all denominators now exclude
    individuals for whom the conduct column does not apply
    (based on their 'bmt_ptp' field).
    """
    # ── setup ───────────────────────────────────────────────────────────────────
    nominal_mapping = {r['name'].strip(): r for r in nominal_data}
    headers = everything_data[0]
    if conduct_header not in headers:
        raise ValueError(f"Conduct column '{conduct_header}' not found.")
    conduct_idx = headers.index(conduct_header)
    conduct_type = classify_conduct(conduct_header)        # NEW

    attendance_mapping = {row[2].strip(): row for row in everything_data[1:]}

    overall_total = overall_present = 0
    platoon_summary, section_summary, individual_details = {}, {}, {}

    # ── main loop ───────────────────────────────────────────────────────────────
    for rec in nominal_data:
        name = rec['name'].strip()
        status = rec.get('bmt_ptp', 'combined').lower()     # NEW

        applies = (
            conduct_type == 'combined'
            or status == 'combined'
            or status == conduct_type
        )                                                  # NEW

        # If this conduct is irrelevant to the soldier, skip all counters
        # but still write "N/A" so the UI can display something.
        if not applies:
            individual_details[name] = {
                'platoon': '', 'section': '', 'roll': '',
                'attendance': 'N/A'
            }
            continue                                       # NEW

        overall_total += 1                                 # CHANGED
        row = attendance_mapping.get(name)
        value = row[conduct_idx].strip() if row and len(row) > conduct_idx else ""
        is_present = value.lower() == "yes"
        if is_present:
            overall_present += 1

        platoon, section, roll = parse_4d_number(rec.get('4d_number', ''))
        if platoon and section:
            platoon_summary.setdefault(platoon, {'total': 0, 'present': 0})
            platoon_summary[platoon]['total'] += 1
            platoon_summary[platoon]['present'] += is_present

            key = (platoon, section)
            section_summary.setdefault(key, {'total': 0, 'present': 0})
            section_summary[key]['total'] += 1
            section_summary[key]['present'] += is_present

        individual_details[name] = {
            'platoon': platoon, 'section': section,
            'roll': roll, 'attendance': value or "Absent"
        }

    overall_pct = (overall_present / overall_total * 100) if overall_total else 0

    # ── column‑by‑column conduct summary ────────────────────────────────────────
    conduct_summary = {}
    for idx in range(3, len(headers)):
        col_header = headers[idx]
        ctype = classify_conduct(col_header)               # NEW
        total = present = 0

        for rec in nominal_data:
            status = rec.get('bmt_ptp', 'combined').lower()
            applies = (
                ctype == 'combined'
                or status == 'combined'
                or status == ctype
            )
            if not applies:
                continue

            total += 1
            row = attendance_mapping.get(rec['name'].strip())
            val = row[idx].strip() if row and len(row) > idx else ""
            if val.lower() == "yes":
                present += 1

        pct = (present / total * 100) if total else 0
        conduct_summary[col_header] = {
            'present': present, 'total': total, 'percentage': pct
        }

    # ── return payload ──────────────────────────────────────────────────────────
    return {
        'overall': {
            'total': overall_total,
            'present': overall_present,
            'percentage': overall_pct
        },
        'platoon_summary': platoon_summary,
        'section_summary': section_summary,
        'individual_details': individual_details,
        'conduct_summary': conduct_summary
    }


def add_conduct_column_everything(sheet_everything, conduct_date: str, conduct_name: str, attendance_data: List[tuple]):
    """
    Adds a new column to the 'Everything' sheet with the conduct details and updates attendance.
    
    Parameters:
    - sheet_everything: gspread Worksheet object for 'Everything' sheet
    - conduct_date (str): Date of the conduct in DDMMYYYY format
    - conduct_name (str): Name of the conduct
    - attendance_data: List of tuples containing (name, rank, is_present)
    """
    # Define the new column header
    new_col_header = f"{conduct_date}, {conduct_name}"
    
    try:
        # Get all data from Everything sheet
        all_data = sheet_everything.get_all_values()
        if not all_data:
            raise ValueError("No data found in Everything sheet")
        
        # Get current number of columns and add new header
        new_col_index = len(all_data[0]) + 1
        sheet_everything.update_cell(1, new_col_index, new_col_header)
        
        # Create a mapping of names to their attendance
        attendance_map = {name: is_present for name, rank, is_present in attendance_data}
        
        # Prepare batch updates
        updates = []
        for row_idx, row in enumerate(all_data[1:], start=2):  # Start from row 2
            name = row[2].strip()  # Assuming Name is in second column
            # Check if this person was in the conduct
            if name in attendance_map:
                value = "Yes" if attendance_map[name] else "No"
            else:
                value = "No"  # Default to No if person wasn't in the conduct
            
            cell = gspread.utils.rowcol_to_a1(row_idx, new_col_index)
            updates.append({
                'range': cell,
                'values': [[value]]
            })
        
        # Batch update the sheet
        if updates:
            sheet_everything.batch_update(updates)
            
    except Exception as e:
        logger.error(f"Error updating Everything sheet: {str(e)}")
        st.error(f"Error updating Everything sheet: {str(e)}")
        return

def update_conduct_column_everything(sheet_everything, conduct_date: str, conduct_name: str, attendance_data: List[tuple]):
    """
    Updates an existing conduct column in the 'Everything' sheet with updated attendance.
    
    Parameters:
    - sheet_everything: gspread Worksheet object for 'Everything' sheet
    - conduct_date (str): Date of the conduct in DDMMYYYY format
    - conduct_name (str): Name of the conduct
    - attendance_data: List of tuples containing (name, rank, is_present)
    """
    target_col_header = f"{conduct_date}, {conduct_name}"
    
    try:
        # Get all data from Everything sheet
        all_data = sheet_everything.get_all_values()
        if not all_data:
            raise ValueError("No data found in Everything sheet")
            
        # Find the column index for the conduct
        headers = all_data[0]
        try:
            conduct_col_index = headers.index(target_col_header) + 1  # 1-based index for gspread
        except ValueError:
            logger.error(f"Conduct column '{target_col_header}' not found in Everything sheet")
            #st.error(f"Conduct column '{target_col_header}' not found in Everything sheet")
            return

        # Create a mapping of names to their attendance
        attendance_map = {name: is_present for name, rank, is_present in attendance_data}
        
        # Prepare updates
        updates = []
        for row_idx, row in enumerate(all_data[1:], start=2):  # Start from 2 to skip header
            name = row[2].strip()  # Assuming Name is in second column
            if name in attendance_map:
                value = "Yes" if attendance_map[name] else "No"
                cell = gspread.utils.rowcol_to_a1(row_idx, conduct_col_index)
                updates.append({
                    'range': cell,
                    'values': [[value]]
                })
        
        # Batch update the sheet
        if updates:
            sheet_everything.batch_update(updates)
            
    except Exception as e:
        logger.error(f"Error updating Everything sheet: {str(e)}")
        st.error(f"Error updating Everything sheet: {str(e)}")
        return
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
            "conducts": sh.worksheet("Conducts"),
            "everything": sh.worksheet("Everything"),
        }
    except Exception as e:
        logger.error(f"Error accessing spreadsheet '{spreadsheet_name}': {e}")
        st.error(f"Error accessing spreadsheet '{spreadsheet_name}': {e}")
        return None
def is_leave_accounted(existing_dates, new_dates_str):
    """
    Checks if the new_dates_str is already present in existing_dates.
    
    Parameters:
        existing_dates (str): Existing "Dates Taken" string, e.g., "15012025-20012025,21012025-22012025"
        new_dates_str (str): New leave dates to check, e.g., "15012025-20012025"
    
    Returns:
        bool: True if the leave already exists, False otherwise.
    """
    if not existing_dates:
        return False
    
    # Split existing dates by comma to get individual leave periods
    existing_leave_periods = [period.strip() for period in existing_dates.split(',')]
    
    # Normalize for comparison
    normalized_existing = set(existing_leave_periods)
    normalized_new = new_dates_str.strip()
    
    return normalized_new in normalized_existing


# ------------------------------------------------------------------------------
# 3) Helper Functions + Caching
# ------------------------------------------------------------------------------

def generate_company_message(selected_company: str, nominal_records: List[Dict], parade_records: List[Dict], target_date: Optional[datetime] = None) -> str:
    """
    Generate a company-specific message in the specified format.

    Parameters:
    - selected_company: The company name.
    - nominal_records: List of nominal records from Nominal_Roll.
    - parade_records: List of parade records from Parade_State.

    Returns:
    - A formatted string message.
    """
    # Get current date and time
    today = target_date if target_date else datetime.now(TIMEZONE)
    t = datetime.now(TIMEZONE)
    date_str = today.strftime("%d%m%y, %A")
    # Determine parade state based on the time: if after 4pm, mark as "LAST PARADE STATE"
    parade_state = "LAST PARADE STATE" if t.hour >= 16 else "FIRST PARADE STATE"

    # Filter nominal records for the selected company
    company_nominal_records = [
        record for record in nominal_records if record['company'] == selected_company
    ]

    # Create a mapping from Name to Rank for quick lookup (case-insensitive)
    name_to_rank = {
        record['name'].strip().lower(): record['rank']
        for record in company_nominal_records
        if record['name']
    }

    # Extract all platoons from nominal records for the selected company
    all_platoons = set(record.get('platoon', 'Coy HQ') for record in company_nominal_records)

    # Initialize a dictionary to hold parade records active today, organized by platoon
    active_parade_by_platoon = defaultdict(list)

    # Process parade records to find those active today and organize them by platoon
    for parade in parade_records:
        if parade.get('company', '') != selected_company:
            continue

        platoon = parade.get('platoon', 'Coy HQ')  # Default to 'Coy HQ' if not specified

        start_str = parade.get('start_date_ddmmyyyy', '')
        end_str = parade.get('end_date_ddmmyyyy', '')
        try:
            start_dt = datetime.strptime(start_str, "%d%m%Y").date()
            end_dt = datetime.strptime(end_str, "%d%m%Y").date()
            if start_dt <= today.date() <= end_dt:
                active_parade_by_platoon[platoon].append(parade)
        except ValueError:
            logger.warning(
                f"Invalid date format for {parade.get('name', '')}: {start_str} - {end_str} in company '{selected_company}'"
            )
            continue

    # Initialize counters for overall nominal and absent strengths
    total_nominal = len(company_nominal_records)
    total_absent = 0

    # Initialize storage for platoon-wise details
    platoon_details = []

    # Sort platoons so that Coy HQ appears first
    sorted_platoons = sorted(all_platoons, key=lambda x: (0, x) if x.lower() == 'coy hq' else (1, x))
    for platoon in sorted_platoons:
        records = active_parade_by_platoon.get(platoon, [])

        # Determine platoon label
        if platoon.lower() == 'coy hq':
            platoon_label = "Coy HQ"
        else:
            platoon_label = f"Platoon {platoon}"

        # Total nominal strength for this platoon
        platoon_nominal = len([
            record for record in company_nominal_records
            if record.get('platoon', 'Coy HQ') == platoon
        ])

        # Initialize lists for conformant absentees split into commander and REC,
        # plus non-conformant parade records (to be shown under "Pl Statuses")
        commander_absentees = []
        rec_absentees = []
        non_conformant_absentees = []

        for parade in records:
            name = parade.get('name', '')
            name_key = name.strip().lower()
            status = parade.get('status', '').upper()
            d = parade.get('4d_number', '')
            start_str = parade.get('start_date_ddmmyyyy', '')
            end_str = parade.get('end_date_ddmmyyyy', '')
            try:
                start_dt = datetime.strptime(start_str, "%d%m%Y").date()
                end_dt = datetime.strptime(end_str, "%d%m%Y").date()
                if start_dt == end_dt:
                    details = f"{start_dt.strftime('%d%m%y')}"
                else:
                    details = f"{start_dt.strftime('%d%m%y')} - {end_dt.strftime('%d%m%y')}"
            except ValueError:
                details = "Invalid Dates"
                logger.warning(
                    f"Invalid dates for {name}: {start_str} - {end_str} in company '{selected_company}'"
                )
            # Look up the nominal rank; default to "N/A" if not found
            rank = name_to_rank.get(name_key, "N/A")
            status_prefix = status.lower().split()[0]
            if status_prefix in LEGEND_STATUS_PREFIXES:
                # Split conformant absentees by whether their rank indicates a REC
                if "rec" in rank.lower():
                    rec_absentees.append({
                        'rank': rank,
                        '4d': d,
                        'name': name,
                        'status': status,
                        'details': details
                    })
                else:
                    commander_absentees.append({
                        'rank': rank,
                        '4d': d,
                        'name': name,
                        'status': status,
                        'details': details
                    })
            else:
                non_conformant_absentees.append({
                    'rank': rank,
                    '4d': d,
                    'name': name,
                    'status': status,
                    'details': details
                })

        # Total absent strength only counts conformant absentees
        commander_group = defaultdict(list)
        for absentee in commander_absentees:
            key = (absentee['4d'].strip(), absentee['rank'].strip(), absentee['name'].strip())
            commander_group[key].append(f"{absentee['status']} {absentee['details']}")

        rec_group = defaultdict(list)
        for absentee in rec_absentees:
            key = (absentee['4d'].strip(), absentee['rank'].strip(), absentee['name'].strip())
            rec_group[key].append(f"{absentee['status']} {absentee['details']}")

        if platoon.lower() != 'coy hq':
            platoon_absent = len(commander_group) + len(rec_group)
        else:
            # For Coy HQ, combine both groups
            combined_group = defaultdict(list)
            for absentee in (commander_absentees + rec_absentees):
                key = (absentee['4d'].strip(), absentee['rank'].strip(), absentee['name'].strip())
                combined_group[key].append(f"{absentee['status']} {absentee['details']}")
            platoon_absent = len(combined_group)
        total_absent += platoon_absent

        # For platoons (other than Coy HQ), calculate nominal breakdown based on rank
        if platoon.lower() != 'coy hq':
            platoon_nominal_records = [
                r for r in company_nominal_records
                if r.get('platoon', 'Coy HQ') == platoon
            ]
            commander_nominal = sum(
                1 for r in platoon_nominal_records if "rec" not in r.get('rank', '').lower()
            )
            rec_nominal = sum(
                1 for r in platoon_nominal_records if "rec" in r.get('rank', '').lower()
            )
        else:
            commander_nominal = None
            rec_nominal = None

        platoon_details.append({
            'label': platoon_label,
            'nominal': platoon_nominal,
            'unique_absent': platoon_absent,  # use the grouped count here
            'present': platoon_nominal - platoon_absent,
            'commander_group': commander_group,
            'rec_group': rec_group,
            'non_conformant': non_conformant_absentees,
            'commander_nominal': commander_nominal,
            'rec_nominal': rec_nominal
        })

    # Calculate overall present strength
    total_present = total_nominal - total_absent

    # Start building the message header
    message_lines = []
    message_lines.append(f"*🐆 {selected_company.upper()} COY*")
    message_lines.append(f"*🗒️ {parade_state}*")
    message_lines.append(f"*🗓️ {date_str}*\n")
    message_lines.append(f"Coy Present Strength: {total_present:02d}/{total_nominal:02d}")
    message_lines.append(f"Coy Absent Strength: {total_absent:02d}/{total_nominal:02d}\n")

    # Build platoon-specific sections
    for detail in platoon_details:
        message_lines.append(f"_*{detail['label']}*_")
        message_lines.append(f"Pl Present Strength: {detail['present']:02d}/{detail['nominal']:02d}")
        message_lines.append(f"Pl Absent Strength: {detail['unique_absent']:02d}/{detail['nominal']:02d}")

        # For platoons other than Coy HQ, show commander/REC breakdown
        if detail['label'] != "Coy HQ":
            message_lines.append(
                f"Commander Absent Strength: {len(detail['commander_group']):02d}/{detail['commander_nominal']:02d}"
            )
            for (d, rank, name), details_list in detail['commander_group'].items():
                details_str = ", ".join(details_list)
                if d:
                    message_lines.append(f"> {d} {rank} {name} ({details_str})")
                else:
                    message_lines.append(f"> {rank} {name} ({details_str})")

            message_lines.append(
                f"REC Absent Strength: {len(detail['rec_group']):02d}/{detail['rec_nominal']:02d}"
            )
            for (d, rank, name), details_list in detail['rec_group'].items():
                details_str = ", ".join(details_list)
                if d:
                    message_lines.append(f"> {d} {rank} {name} ({details_str})")
                else:
                    message_lines.append(f"> {rank} {name} ({details_str})")
        else:
            # For Coy HQ, combine commander and REC into one list
            combined_group = defaultdict(list)
            for key, details_list in detail['commander_group'].items():
                combined_group[key].extend(details_list)
            for key, details_list in detail['rec_group'].items():
                combined_group[key].extend(details_list)
            for (d, rank, name), details_list in combined_group.items():
                details_str = ", ".join(details_list)
                if d:
                    message_lines.append(f"> {d} {rank} {name} ({details_str})")
                else:
                    message_lines.append(f"> {rank} {name} ({details_str})")

        # Add non-conformant parade statuses if any exist
        status_group = defaultdict(list)
        if detail['non_conformant']:
            for person in detail['non_conformant']:
                rank = person['rank']
                name = person['name']
                d = person['4d']
                status_code = person['status']
                details_str = person['details']
                key = (rank, name, d)
                status_entry = f"{status_code} {details_str}"
                status_group[key].append(status_entry)
            pl_status_count = len(status_group)
            message_lines.append(f"\nPl Statuses: {pl_status_count:02d}/{detail['nominal']:02d}")
            for (rank, name, d), details_list in status_group.items():
                if rank and name and d:
                    line_prefix = f"> {d} {rank} {name}"
                else:
                    line_prefix = f"> {rank} {name}"
                consolidated_details = ", ".join(details_list)
                message_lines.append(f"{line_prefix} ({consolidated_details})")

        message_lines.append("")  # Blank line for separation

    final_message = "\n".join(message_lines)
    return final_message



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
        # We log an error if it "looks" invalid, but we won't remove it from nominal if blank
        if four_d != '4D':  # i.e., truly invalid, not just empty
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
        cleaned = re.sub(r'\D', '', date_value)
        return cleaned.zfill(8)
    else:
        return ""

def normalize_name(name: str) -> str:
    """Normalize by uppercase + removing spaces and special characters."""
    return re.sub(r'\W+', '', name.upper())
def classify_conduct(header: str) -> str:
    """
    Returns 'bmt', 'ptp', or 'combined' for a conduct column name.
    Anything not explicitly BMT or PTP is treated as combined.
    """
    h = header.upper()
    if 'BMT' in h and 'PTP' not in h:
        return 'bmt'
    if 'PTP' in h and 'BMT' not in h:
        return 'ptp'
    return 'combined'
def get_nominal_records(selected_company: str, _sheet_nominal):
    """
    Returns all rows from Nominal_Roll as a list of dicts.
    Handles case-insensitive and whitespace-trimmed headers.no 
    Includes the 'company' field in each record.
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
        normalized_row['company'] = selected_company  # Add company information
        normalized_row['bmt_ptp'] = ensure_str(
            normalized_row.get('bmt_ptp', 'combined')
        ).lower()
        normalized_row['ration_type'] = ensure_str(
            normalized_row.get('ration_type', '')
        ).lower()
        normalized_records.append(normalized_row)
    
    return normalized_records

#
# Parade uses 'Name' (not 4D) for matching
#
def get_parade_records(selected_company: str, _sheet_parade):
    """
    Returns all rows from Parade_State as a list of dicts, including row numbers.
    Only includes statuses where End_Date is today or in the future.
    Uses 'name' to identify the individual (instead of '4d_number').
    Includes the 'company' field in each record.
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
            continue

        record = dict(zip(header, row))
        
        # Use Name
        record['name'] = ensure_str(record.get('name', ''))
        record['platoon'] = ensure_str(record.get('platoon', ''))
        record['4d_number'] = ensure_str(record.get('4d_number', ''))  # We'll keep it for any leaves logic
        record['start_date_ddmmyyyy'] = ensure_date_str(record.get('start_date_ddmmyyyy', ''))
        record['end_date_ddmmyyyy'] = ensure_date_str(record.get('end_date_ddmmyyyy', ''))
        record['status'] = ensure_str(record.get('status', ''))
        record['company'] = selected_company  # Add company information

        try:
            ed = datetime.strptime(record['end_date_ddmmyyyy'], "%d%m%Y").date()
            if ed >= today:
                record['_row_num'] = idx
                records.append(record)
        except ValueError:
            logger.warning(
                f"Invalid date format in Parade_State for {record.get('name', '')}: "
                f"{record.get('end_date_ddmmyyyy', '')}"
            )
            continue

    return records
def get_allparade_records(selected_company: str, _sheet_parade):
    """
    Returns all rows from Parade_State as a list of dicts, including row numbers.
    Only includes statuses where End_Date is today or in the future.
    Uses 'name' to identify the individual (instead of '4d_number').
    Includes the 'company' field in each record.
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
            continue

        record = dict(zip(header, row))
        
        # Use Name
        record['name'] = ensure_str(record.get('name', ''))
        record['platoon'] = ensure_str(record.get('platoon', ''))
        record['4d_number'] = ensure_str(record.get('4d_number', ''))  # We'll keep it for any leaves logic
        record['start_date_ddmmyyyy'] = ensure_date_str(record.get('start_date_ddmmyyyy', ''))
        record['end_date_ddmmyyyy'] = ensure_date_str(record.get('end_date_ddmmyyyy', ''))
        record['status'] = ensure_str(record.get('status', ''))
        record['company'] = selected_company  # Add company information

        try:
            ed = datetime.strptime(record['end_date_ddmmyyyy'], "%d%m%Y").date()
            if ed:
                record['_row_num'] = idx
                records.append(record)
        except ValueError:
            logger.warning(
                f"Invalid date format in Parade_State for {record.get('name', '')}: "
                f"{record.get('end_date_ddmmyyyy', '')}"
            )
            continue

    return records

def get_conduct_records(selected_company: str, _sheet_conducts):
    """
    Returns all rows from Conducts as a list of dicts.
    """
    records = _sheet_conducts.get_all_records()
    if not records:
        logger.warning(f"No records found in Conducts for company '{selected_company}'.")
        return []

    normalized_records = []
    for row in records:
        normalized_row = {k.strip().lower(): v for k, v in row.items()}
        normalized_row['date'] = ensure_date_str(normalized_row.get('date', ''))
        normalized_row['conduct_name'] = ensure_str(normalized_row.get('conduct_name', ''))
        normalized_row['p/t plt1'] = ensure_str(normalized_row.get('p/t plt1', '0/0'))
        normalized_row['p/t plt2'] = ensure_str(normalized_row.get('p/t plt2', '0/0'))
        normalized_row['p/t plt3'] = ensure_str(normalized_row.get('p/t plt3', '0/0'))
        normalized_row['p/t plt4'] = ensure_str(normalized_row.get('p/t plt4', '0/0'))
        normalized_row['p/t total'] = ensure_str(normalized_row.get('p/t total', '0/0'))
        normalized_row['plt1 outliers'] = ensure_str(normalized_row.get('plt1 outliers', ''))
        normalized_row['plt2 outliers'] = ensure_str(normalized_row.get('plt2 outliers', ''))
        normalized_row['plt3 outliers'] = ensure_str(normalized_row.get('plt3 outliers', ''))
        normalized_row['plt4 outliers'] = ensure_str(normalized_row.get('plt4 outliers', ''))
        normalized_row['coy hq outliers'] = ensure_str(normalized_row.get('coy hq outliers', ''))
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

#
# Map parade statuses by 'name' instead of '4d_number'
#
def get_company_personnel(platoon: str, records_nominal, records_parade):
    """
    Returns a list of dicts for 'Update Parade' with existing parade statuses first,
    followed by all nominal rows without statuses. 
    Matches by 'name' (uppercase) instead of '4d_number'.
    """
    from collections import defaultdict
    parade_map = defaultdict(list)
    for row in records_parade:
        person_name = row.get('name', '').strip().upper()
        parade_map[person_name].append(row)
    
    data_with_status = []
    data_nominal = []
    
    for row in records_nominal:
        p = row.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue

        rank = row.get('rank', '')
        original_name = row.get('name', '')
        four_d = row.get('4d_number', '')

        name_key = original_name.strip().upper()

        # Retrieve all parade statuses for the person (by name)
        person_parades = parade_map.get(name_key, [])
        for parade in person_parades:
            data_with_status.append({
                'Rank': rank,
                'Name': original_name,
                '4D_Number': four_d,
                'Status': parade.get('status', ''),
                'Start_Date': parade.get('start_date_ddmmyyyy', ''),
                'End_Date': parade.get('end_date_ddmmyyyy', ''),
                'Number_of_Leaves_Left': row.get('number of leaves left', 14),
                'Dates_Taken': row.get('dates taken', ''),
                '_row_num': parade.get('_row_num')
            })

        # Add the nominal entry without status
        data_nominal.append({
            'Rank': rank,
            'Name': original_name,
            '4D_Number': four_d,
            'Status': '',
            'Start_Date': '',
            'End_Date': '',
            'Number_of_Leaves_Left': row.get('number of leaves left', 14),
            'Dates_Taken': row.get('dates taken', ''),
            '_row_num': None
        })
    
    combined_data = data_with_status + data_nominal
    return combined_data

def find_name_by_4d(four_d: str, records_nominal) -> str:
    """
    If you want to look up person's Name from Nominal_Roll given a 4D_Number.
    """
    four_d = ensure_str(four_d).upper()
    for row in records_nominal:
        if ensure_str(row.get("4d_number", "")).upper() == four_d:
            return ensure_str(row.get("name", ""))
    return ""

def build_onstatus_table(platoon: str, date_obj: datetime, records_nominal, records_parade):
    """
    Return a list of dicts for everyone on status for that date + platoon.
    If multiple statuses exist for the same person, prioritize based on a small hierarchy.
    """
    status_priority = {'leave': 3, 'fever': 2, 'mc': 1}
    out = {}
    parade_map = defaultdict(list)
    for row in records_parade:
        person_name = row.get('name', '').strip().upper()
        parade_map[person_name].append(row)
    
    for row in records_nominal:
        p = row.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        name = row.get('name', '')
        rank = row.get('rank', '')
        four_d = row.get('4d_number', '')
        name_key = name.strip().upper()

        for parade in parade_map.get(name_key, []):
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', '01012000'), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', '01012000'), "%d%m%Y").date()
                if start_dt <= date_obj.date() <= end_dt:
                    status = ensure_str(parade.get('status', '')).lower()
                    if status in status_priority:
                        if name_key in out:
                            existing_status = out[name_key]['StatusDesc'].lower()
                            if status_priority.get(status, 0) > status_priority.get(existing_status, 0):
                                out[name_key] = {
                                    "Rank": rank,
                                    "Name": name,
                                    "4D_Number": four_d,
                                    "StatusDesc": ensure_str(parade.get('status', '')),
                                    "Is_Outlier": True
                                }
                        else:
                            out[name_key] = {
                                "Rank": rank,
                                "Name": name,
                                "4D_Number": four_d,
                                "StatusDesc": ensure_str(parade.get('status', '')),
                                "Is_Outlier": True
                            }
            except ValueError:
                logger.warning(
                    f"Invalid date format for {name_key}: "
                    f"{parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}"
                )
                continue
    logger.info(f"Built on-status table with {len(out)} entries for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
    return list(out.values())

def build_conduct_table(platoon: str, date_obj: datetime, records_nominal, records_parade):
    """
    Return a list of dicts for all personnel in the platoon.
    'Is_Outlier' is True if the person has an active status on the given date.
    """
    parade_map = defaultdict(list)
    for row in records_parade:
        person_name = row.get('name', '').strip().upper()
        parade_map[person_name].append(row)
    
    data = []
    for person in records_nominal:
        p = person.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        name = person.get('name', '')
        rank = person.get('rank', '')
        four_d = person.get('4d_number', '')
        name_key = name.strip().upper()

        active_statuses = []  # List to hold all active statuses for the person
        

        for parade in parade_map.get(name_key, []):
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                if start_dt <= date_obj.date() <= end_dt:
                    status = parade.get('status', '').strip().upper()
                    if status:  # Ensure status is not empty
                        active_statuses.append(status)
            except ValueError:
                logger.warning(
                    f"Invalid date format for {name_key}: "
                    f"{parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}"
                )
                continue
        is_outlier = len(active_statuses) > 0
        status_desc = ", ".join(active_statuses) if is_outlier else ""
        data.append({
            'Rank': rank,
            'Name': name,
            '4D_Number': four_d,
            'Is_Outlier': is_outlier,
            'StatusDesc': status_desc
        })
    logger.info(f"Built conduct table with {len(data)} personnel for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
    for person in data:
        if person.get("Rank", "").upper() == "REC":
            person["Personnel_Type"] = "Recruit"
        else:
            person["Personnel_Type"] = "Commander"
    return data
def build_fake_conduct_table(platoon: str, date_obj: datetime, records_nominal, records_parade):
    """
    Return a list of dicts for all personnel in the platoon.
    'Is_Outlier' is True if the person has an active status on the given date.
    """
    parade_map = defaultdict(list)
    for row in records_parade:
        person_name = row.get('name', '').strip().upper()
        parade_map[person_name].append(row)
    
    data = []
    for person in records_nominal:
        p = person.get('platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        name = person.get('name', '')
        rank = person.get('rank', '')
        four_d = person.get('4d_number', '')
        name_key = name.strip().upper()

        active_statuses = []  # List to hold all active statuses for the person
        

        for parade in parade_map.get(name_key, []):
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                if start_dt <= date_obj.date() <= end_dt:
                    status = parade.get('status', '').strip().upper()
                    if status:  # Ensure status is not empty
                        active_statuses.append(status)
            except ValueError:
                logger.warning(
                    f"Invalid date format for {name_key}: "
                    f"{parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}"
                )
                continue
        active_statuses = 0
        is_outlier = active_statuses > 0
        status_desc = ", ".join(active_statuses) if is_outlier else ""
        data.append({
            'Rank': rank,
            'Name': name,
            '4D_Number': four_d,
            'Is_Outlier': is_outlier,
            'StatusDesc': status_desc
        })
    logger.info(f"Built conduct table with {len(data)} personnel for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
    for person in data:
        if person.get("Rank", "").upper() == "REC":
            person["Personnel_Type"] = "Recruit"
        else:
            person["Personnel_Type"] = "Commander"
    return data
def calculate_leaves_used(dates_str: str) -> int:
    """
    Calculate the number of leaves used based on the dates string (which may include ranges).
    """
    if not dates_str:
        return 0

    leaves_used = 0
    date_entries = [entry.strip() for entry in dates_str.split(',') if entry.strip()]
    for entry in date_entries:
        if '-' in entry:
            start_str, end_str = entry.split('-')
            start_str = ensure_date_str(start_str)
            end_str = ensure_date_str(end_str)
            try:
                start_dt = datetime.strptime(start_str, "%d%m%Y")
                end_dt = datetime.strptime(end_str, "%d%m%Y")
                delta = (end_dt - start_dt).days + 1
                if delta > 0:
                    leaves_used += delta
            except ValueError:
                logger.warning(f"Invalid date range format: {entry}")
                continue
        else:
            single_str = ensure_date_str(entry)
            try:
                datetime.strptime(single_str, "%d%m%Y")
                leaves_used += 1
            except ValueError:
                logger.warning(f"Invalid single date format: {entry}")
                continue
    logger.info(f"Calculated leaves used: {leaves_used} from dates: {dates_str}")
    return leaves_used

def has_overlapping_status(four_d: str, new_start: datetime, new_end: datetime, records_parade):
    """
    Check if the new status dates overlap with existing statuses for the given 4D_Number (for leave).
    """
    four_d = is_valid_4d(four_d)
    if not four_d:
        return False
    
    for row in records_parade:
        if is_valid_4d(row.get("4d_number", "")) == four_d:
            start_date = row.get("start_date_ddmmyyyy", "")
            end_date = row.get("end_date_ddmmyyyy", "")
            
            try:
                existing_start = datetime.strptime(start_date, "%d%m%Y")
                existing_end = datetime.strptime(end_date, "%d%m%Y")
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

selected_company = st.sidebar.selectbox(
    "Select Company",
    options=st.session_state.user_companies
)

worksheets = get_sheets(selected_company)
if not worksheets:
    st.error("Failed to load the selected company's spreadsheets. Please check the logs for more details.")
    st.stop()

SHEET_NOMINAL = worksheets["nominal"]
SHEET_PARADE = worksheets["parade"]
SHEET_CONDUCTS = worksheets["conducts"]

# ------------------------------------------------------------------------------
# 6) Session State
# ------------------------------------------------------------------------------
if "conduct_date" not in st.session_state:
    st.session_state.conduct_date = ""
if "conduct_platoon" not in st.session_state:
    st.session_state.conduct_platoon = 1
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

if "parade_platoon" not in st.session_state:
    st.session_state.parade_platoon = 1
if "parade_table" not in st.session_state:
    st.session_state.parade_table = []

if "update_conduct_selected" not in st.session_state:
    st.session_state.update_conduct_selected = None
if "update_conduct_platoon" not in st.session_state:
    st.session_state.update_conduct_platoon = 1
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
    ["Add Conduct", "Update Conduct", "Update Parade", "Queries", "Overall View", "Generate WhatsApp Message", "Safety Sharing"]
)

def add_pointer():
    st.session_state.conduct_pointers.append(
        {"observation": "", "reflection": "", "recommendation": ""}
    )
    #Sst.rerun()
def add_update_pointer():
    st.session_state.update_conduct_pointers.append(
        {"observation": "", "reflection": "", "recommendation": ""}
    )
    #st.rerun()

# ------------------------------------------------------------------------------
# 8) Feature A: Add Conduct
# ------------------------------------------------------------------------------
if feature == "Add Conduct":
    st.header("Add Conduct - Table-Based On-Status")

    st.session_state.conduct_date = st.text_input(
        "Date (DDMMYYYY)",
        value=st.session_state.conduct_date
    )
    platoon_options = ["1", "2", "3", "4", "Coy HQ"]
    st.session_state.conduct_platoon = st.selectbox(
        "Your Platoon",
        options=platoon_options,
        index=platoon_options.index(st.session_state.conduct_platoon) if st.session_state.conduct_platoon in platoon_options else 0
    )
    conduct_options = [
        "",
        "2.4KM CONDITIONING RUN",
        "AQUA",
        "ARM DRILLS",
        "BALANCING, FLEXIBILITY, MOBILITY",
        "BCCT",
        "BIC",
        "BPK",
        "BSK",
        "BTP",
        "CBA&CPR-AED",
        "CADENCE RUN",
        "CAREHAB ADJUSTMENT SURVEY",
        "COMBAT CIRCUIT",
        "CONDUCTING BRIEF BTP",
        "COY FIRE DRILL",
        "CPT",
        "CQB",
        "DAC",
        "DISTANCE INTERVAL",
        "ELISS FAMILIARISATION",
        "ENDURANCE RUN",
        "ENDURANCE RUN TEMPO",
        "FARTLEK",
        "FIELDCAMP",
        "FOOT DRILLS",
        "GP REHEARSAL",
        "GYM ORIENTATION",
        "GYM TRAINING",
        "HAND GRENADE",
        "IFC",
        "IMT",
        "INFANTRY SMALL ARMS DEMONSTRATION",
        "INTRO TO HEARTRATE",
        "INTRO TO UO",
        "IPPT",
        "JS3",
        "JUDGEMENTAL SHOOT",
        "LEADERSHIP VALUES",
        "MO TALK",
        "METABOLIC CIRCUIT",
        "NATIONAL EDUCATION",
        "OO ENGAGEMENT",
        "ORIENTATION RUN",
        "PHYSICAL TRAINING LECTURE",
        "RAMADHAN BRIEF",
        "RESILIENCE LEARNING",
        "ROUTE MARCH(3KM)",
        "ROUTE MARCH (4KM)",
        "ROUTE MARCH (8KM)",
        "ROUTE MARCH (12KM)",
        "SAFE & INCLUSIVE WORKPLACE",
        "SAFRA TALK",
        "SIT TEST",
        "SOC",
        "SPEED AGILITY QUICKNESS",
        "SPORTS AND GAMES",
        "STRENGTH TRAINING",
        "TECHNICAL HANDLING",
        "WEAPON PRESENTATION PREPARATION"
    ]

    st.session_state.conduct_name = st.selectbox(
        "Conduct Name",
        options=conduct_options,
        index=conduct_options.index(st.session_state.conduct_name) if st.session_state.conduct_name in conduct_options else 0
    )

    suffix_options = ["COMBINED","BMT", "PTP"]
    if 'conduct_suffix' not in st.session_state:
        st.session_state.conduct_suffix = suffix_options[0]
    st.session_state.conduct_suffix = st.selectbox(
        "Phase",
        options=suffix_options,
        index=suffix_options.index(st.session_state.conduct_suffix) if st.session_state.conduct_suffix in suffix_options else 0
    )
    
    # Add a separate dropdown for MUT suffix
    mut_options = ["", "MUT"]
    if 'mut_suffix' not in st.session_state:
        st.session_state.mut_suffix = mut_options[0]
    st.session_state.mut_suffix = st.selectbox(
        "MUT (optional)",
        options=mut_options,
        index=mut_options.index(st.session_state.mut_suffix) if st.session_state.mut_suffix in mut_options else 0
    )

    if 'conduct_session' not in st.session_state:
        st.session_state.conduct_session = 1
    # Only show session number input if a conduct is selected
    if st.session_state.conduct_name:
        # Session Number Input using number_input
        st.session_state.conduct_session = st.number_input(
            "Session Number",
            min_value=1,
            step=1,
            value=int(st.session_state.conduct_session) if isinstance(st.session_state.conduct_session, int) else 1
        )

    # Display Final Conduct Name
    if st.session_state.conduct_name and st.session_state.conduct_session:
        # Combine the two suffixes with a space in between if both are present
        combined_suffix = st.session_state.conduct_suffix
        if st.session_state.mut_suffix:
            combined_suffix = f"{combined_suffix} {st.session_state.mut_suffix}"
            
        final_conduct_name = f"{st.session_state.conduct_name} {combined_suffix} {st.session_state.conduct_session}"
        st.write(f"**Final Conduct Name:** {final_conduct_name}")
    elif st.session_state.conduct_name:
        st.write(f"**Final Conduct Name:** {st.session_state.conduct_name}")

    if 'conduct_pointers' not in st.session_state:
        st.session_state.conduct_pointers = [
        {"observation": "", "reflection": "", "recommendation": ""}
    ] 

    st.subheader("Pointers (ORR, Observation, Reflection)")

    # Render input fields for each pointer in the session state
    for idx, pointer in enumerate(st.session_state.conduct_pointers):
        st.markdown(f"**Pointer {idx + 1}:**")
        col1, col2, col3 = st.columns(3)
        with col1:
            pointer["observation"] = st.text_input(
                f"Observation {idx + 1}",
                value=pointer["observation"],
                key=f"observation_{idx}"
            )
        with col2:
            pointer["reflection"] = st.text_input(
                f"Reflection {idx + 1}",
                value=pointer["reflection"],
                key=f"reflection_{idx}"
            )
        with col3:
            pointer["recommendation"] = st.text_input(
                f"Recommendation {idx + 1}",
                value=pointer["recommendation"],
                key=f"recommendation_{idx}"
            )
        st.markdown("---")  # Separator between pointers

    # Button to add a new pointer
    st.button("➕ Add Another Pointer", on_click=add_pointer)

    submitted_by = st.session_state.username

    if st.button("Load On-Status"):
        date_str = st.session_state.conduct_date.strip()
        platoon = str(st.session_state.conduct_platoon).strip()

        if not date_str or not platoon:
            st.error("Please enter both Date and Platoon.")
            st.stop()

        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format (use DDMMYYYY).")
            st.stop()

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_allparade_records(selected_company, SHEET_PARADE)

        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)

        st.session_state.conduct_table = conduct_data
        st.success(f"Loaded {len(conduct_data)} personnel for Platoon {platoon} ({date_obj.strftime('%d%m%Y')}).")
        logger.info(
            f"Loaded conduct personnel for Platoon {platoon} on {date_obj.strftime('%d%m%Y')} "
            f"in company '{selected_company}' by user '{submitted_by}'."
        )

    if st.session_state.conduct_table:
        st.write("Toggle 'Is_Outlier' if not participating, or add new rows for extra people.")
        sorted_conduct_table = sorted(st.session_state.conduct_table, 
                                 key=lambda x: "ZZZ" if x.get("Rank", "").upper() == "REC" else x.get("Rank", ""))
        edited_data = st.data_editor(
            st.session_state.conduct_table,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
        )
    else:
        edited_data = st.data_editor(
            [],
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
        )

    if st.button("Finalize Conduct"):
        date_str = st.session_state.conduct_date.strip()
        platoon = str(st.session_state.conduct_platoon).strip()
        cname = final_conduct_name.strip()
        observation = st.session_state.conduct_pointers_observation.strip()
        reflection = st.session_state.conduct_pointers_reflection.strip()
        recommendation = st.session_state.conduct_pointers_recommendation.strip()

        if not date_str or not platoon or not cname:
            st.error("Please fill all fields (Date, Platoon, Conduct Name) first.")
            st.stop()

        try:
            datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format.")
            st.stop()

        pointers_list = []
        for idx, pointer in enumerate(st.session_state.conduct_pointers, start=1):
            observation = pointer.get("observation", "").strip()
            reflection = pointer.get("reflection", "").strip()
            recommendation = pointer.get("recommendation", "").strip()

            pointer_str = ""
            if observation:
                pointer_str += f"Observation {idx}:\n{observation}\n"
            if reflection:
                pointer_str += f"Reflection {idx}:\n{reflection}\n"
            if recommendation:
                pointer_str += f"Recommendation {idx}:\n{recommendation}\n"
            pointers_list.append(pointer_str.strip())

        pointers = "\n\n".join(pointers_list)

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_allparade_records(selected_company, SHEET_PARADE)

        existing_4ds = {row.get("4d_number", "").strip().upper() for row in records_nominal}
        new_people = []
        all_outliers = []
        four_d_to_name = {row['4d_number']: row['name'] for row in records_nominal}

        for row in edited_data:
            four_d = is_valid_4d(row.get("4D_Number", ""))
            name_ = ensure_str(row.get("Name", ""))
            rank_ = ensure_str(row.get("Rank", ""))
            is_outlier = row.get("Is_Outlier", False)
            status_desc = ensure_str(row.get("StatusDesc", ""))

            # If both 4D and Name are missing, skip
            if not four_d and not name_:
                st.error(f"No valid Name/4D_Number provided. Skipping entry.")
                continue

            # If person is new (by Name), add to nominal if not found
            if name_ and all(n_.get("name", "").strip().upper() != name_.strip().upper() for n_ in records_nominal):
                if not rank_:
                    st.error(f"Rank is required for new Name '{name_}'. Skipping.")
                    logger.error(f"Rank missing for new Name: {name_}.")
                    continue
                new_people.append((rank_, name_, four_d, platoon))
                logger.info(
                    f"Adding new person: Rank={rank_}, Name={name_}, 4D_Number={four_d}, "
                    f"Platoon={platoon} in company '{selected_company}' by user '{submitted_by}'."
                )

            if is_outlier:
                if status_desc:
                    all_outliers.append(f"{four_d} {name_} ({status_desc})" if four_d else f"{name_} ({status_desc})")
                else:
                    all_outliers.append(f"{four_d} {name_}" if four_d else f"{name_}")

        for (rank, nm, fd, p_) in new_people:
            final_fd = fd if fd else ""
            SHEET_NOMINAL.append_row([rank, nm, final_fd, p_, 14, ""])  
            logger.info(
                f"Added new person to Nominal_Roll: Rank={rank}, Name={nm}, 4D_Number={final_fd}, "
                f"Platoon={p_} in company '{selected_company}' by user '{submitted_by}'."
            )

        total_strength_platoons = {}
        # Updated to include 'Coy HQ'
        for plt in platoon_options:
            strength = get_company_strength(plt, records_nominal)
            total_strength_platoons[plt] = strength
            print(total_strength_platoons[plt])

        # Initialize recruit and commander counts for each platoon
        rec_counts = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}
        cmd_counts = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}
        rec_totals = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}
        cmd_totals = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}

        # Calculate total rec and cmd for each platoon
        for person in records_nominal:
            plt = person.get("platoon", "")
            if plt in platoon_options:
                if person.get("rank", "").upper() == "REC":
                    rec_totals[plt] += 1
                else:
                    cmd_totals[plt] += 1

        # Count participating recruits and commanders
        for row in edited_data:
            if not row.get('Is_Outlier', False):
                plt = row.get('Platoon', platoon)
                if plt in platoon_options:
                    if row.get('Rank', '').upper() == 'REC':
                        rec_counts[plt] += 1
                    else:
                        cmd_counts[plt] += 1

        # Initialize pt_plts with detailed format for all platoons
        pt_plts = ['0/0\n0/0\n0/0', '0/0\n0/0\n0/0', '0/0\n0/0\n0/0', '0/0\n0/0\n0/0', '0/0\n0/0\n0/0']

        # Update the platoon that's participating in this conduct
        if platoon in platoon_options:
            if platoon != "Coy HQ":
                index = int(platoon) - 1  # Platoons 1-4 map to indices 0-3
            else:
                index = 4  # 'Coy HQ' maps to index 4
            
            rec_ratio = f"{rec_counts[platoon]}/{rec_totals[platoon]}"
            cmd_ratio = f"{cmd_counts[platoon]}/{cmd_totals[platoon]}"
            total_ratio = f"{rec_counts[platoon] + cmd_counts[platoon]}/{total_strength_platoons[platoon]}"
            
            pt_plts[index] = f"REC: {rec_ratio}\nCMD: {cmd_ratio}\nTOTAL: {total_ratio}"

        # Calculate total participants and total strength
        total_rec_part = sum(rec_counts.values())
        total_rec = sum(rec_totals.values())
        total_cmd_part = sum(cmd_counts.values())
        total_cmd = sum(cmd_totals.values())
        total_part = total_rec_part + total_cmd_part
        total_strength = sum(total_strength_platoons.values())

        # Format the totals
        pt_total = f"REC: {total_rec_part}/{total_rec}\nCMD: {total_cmd_part}/{total_cmd}\nTOTAL: {total_part}/{total_strength}"

        formatted_date_str = ensure_date_str(date_str)
        # Prepare outliers per platoon – order: PLT1, PLT2, PLT3, PLT4, Coy HQ
        outliers_list = ["None"] * 5
        if platoon in platoon_options:
            index = int(platoon) - 1 if platoon != "Coy HQ" else 4
            outliers_list[index] = ", ".join(all_outliers) if all_outliers else "None"

        SHEET_CONDUCTS.append_row([
            formatted_date_str,  # Column 1: Date
            cname,               # Column 2: Conduct_Name
            pt_plts[0],          # Column 3: P/T PLT1 (detailed format)
            pt_plts[1],          # Column 4: P/T PLT2 (detailed format)
            pt_plts[2],          # Column 5: P/T PLT3 (detailed format)
            pt_plts[3],          # Column 6: P/T PLT4 (detailed format)
            pt_plts[4],          # Column 7: P/T Coy HQ (detailed format)
            pt_total,            # Column 8: P/T Total (detailed format)
            outliers_list[0],    # Column 9: PLT1 Outliers
            outliers_list[1],    # Column 10: PLT2 Outliers
            outliers_list[2],    # Column 11: PLT3 Outliers
            outliers_list[3],    # Column 12: PLT4 Outliers
            outliers_list[4],    # Column 13: Coy HQ Outliers
            pointers,            # Column 14: Pointers
            submitted_by         # Column 15: Submitted_By
        ])

        logger.info(
            f"Appended Conduct: {formatted_date_str}, {cname}, "
            f"P/T PLT1: {pt_plts[0]}, P/T PLT2: {pt_plts[1]}, P/T PLT3: {pt_plts[2]}, "
            f"P/T PLT4: {pt_plts[3]}, P/T Coy HQ: {pt_plts[4]}, P/T Total: {pt_total}, Outliers: {', '.join(all_outliers) if all_outliers else 'None'}, "
            f"Pointers: {pointers}, Submitted_By: {submitted_by} in company '{selected_company}'."
        )

        SHEET_EVERYTHING = worksheets["everything"]
        attendance_data = extract_attendance_data(edited_data)
        add_conduct_column_everything(
            SHEET_EVERYTHING,
            formatted_date_str,
            cname,
            attendance_data
        )


        try:
            conduct_cell = SHEET_CONDUCTS.find(cname, in_column=2)
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

        try:
            SHEET_CONDUCTS.update_cell(conduct_row, 8, pt_total)
            logger.info(f"Updated P/T Total to {pt_total} for conduct '{cname}' in company '{selected_company}'.")
        except Exception as e:
            st.error(f"Error updating P/T Total: {e}")
            logger.error(f"Exception while updating P/T Total for conduct '{cname}': {e}")
            st.stop()

        st.success(
            f"Conduct Finalized!\n\n"
            f"Date: {formatted_date_str}\n"
            f"Conduct Name: {cname}\n"
            f"P/T PLT1: {pt_plts[0]}\n"
            f"P/T PLT2: {pt_plts[1]}\n"
            f"P/T PLT3: {pt_plts[2]}\n"
            f"P/T PLT4: {pt_plts[3]}\n"
            f"P/T Coy HQ: {pt_plts[4]}\n"
            f"P/T Total: {pt_total}\n"
            f"Outliers: {', '.join(all_outliers) if all_outliers else 'None'}\n"
            f"Pointers:\n{pointers if pointers else 'None'}\n"
            f"Submitted By: {submitted_by}"
        )

        st.session_state.conduct_date = ""
        st.session_state.conduct_platoon = platoon_options[0]
        st.session_state.conduct_name = conduct_options[0]
        st.session_state.conduct_table = []
        st.session_state.conduct_pointers = [
            {"observation": "", "reflection": "", "recommendation": ""}
        ]

# ------------------------------------------------------------------------------
# 9) Feature B: Update Conduct
# ------------------------------------------------------------------------------
elif feature == "Update Conduct":
    st.header("Update Conduct")

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

    # Extract date and name from the selected conduct
    try:
        selected_parts = selected_conduct.split(" - ")
        selected_date = selected_parts[0].strip()
        selected_name = selected_parts[1].strip()
        
        # Find the exact matching record by both date and name
        matching_records = [r for r in records_conducts 
                          if r.get('date') == selected_date and r.get('conduct_name') == selected_name]
        
        if not matching_records:
            st.error(f"No conduct found with date '{selected_date}' and name '{selected_name}'")
            logger.error(f"Conduct matching failed for '{selected_conduct}'")
            st.stop()
            
        # Use the first match (should be unique if date+name is unique)
        conduct_record = matching_records[0]
        
        # Log if multiple matches found (shouldn't happen)
        if len(matching_records) > 1:
            logger.warning(f"Multiple matching records found for '{selected_conduct}'. Using the first match.")
            
    except Exception as e:
        st.error(f"Error finding conduct record: {e}")
        logger.error(f"Exception while finding conduct record for '{selected_conduct}': {e}")
        st.stop()

    st.subheader("Select Platoon to Update")
    platoon_options = ["1", "2", "3", "4", "Coy HQ"]
    st.session_state.conduct_platoon = st.selectbox( 
        "Select Platoon",
        options=platoon_options,
        index=platoon_options.index(str(st.session_state.conduct_platoon)) if str(st.session_state.conduct_platoon) in platoon_options else 0,
        key="update_conduct_platoon_select"
    )

        # Initialize a session state variable to track the previous selection
    if 'update_conduct_selected_prev' not in st.session_state:
        st.session_state.update_conduct_selected_prev = None

    if 'update_platoon_selected_prev' not in st.session_state:
        st.session_state.update_platoon_selected_prev = None

    current_selected_conduct = selected_conduct
    current_selected_platoon = st.session_state.conduct_platoon
    if current_selected_platoon != st.session_state.update_platoon_selected_prev:
        # Update the previous selection
        st.session_state.update_platoon_selected_prev = current_selected_platoon
        
        # Re-initialize the pointers based on the newly selected conduct
        existing_pointers = conduct_record.get('pointers', '')
        st.session_state.update_conduct_pointers = []
        
        if existing_pointers:
            # Split pointers by double newlines assuming each pointer is separated by two newlines
            pointer_entries = existing_pointers.split('\n\n')
            for entry in pointer_entries:
                observation = ""
                reflection = ""
                recommendation = ""
                
                # Extract Observation, Reflection, Recommendation using regex
                obs_match = re.search(r'Observation\s*\d*:\s*([\s\S]*?)(?:\n|$)', entry, re.IGNORECASE)
                refl_match = re.search(r'Reflection\s*\d*:\s*([\s\S]*?)(?:\n|$)', entry, re.IGNORECASE)
                rec_match = re.search(r'Recommendation\s*\d*:\s*([\s\S]*?)(?:\n|$)', entry, re.IGNORECASE)
                
                if obs_match:
                    observation = obs_match.group(1).strip()
                if refl_match:
                    reflection = refl_match.group(1).strip()
                if rec_match:
                    recommendation = rec_match.group(1).strip()
                
                st.session_state.update_conduct_pointers.append({
                    "observation": observation,
                    "reflection": reflection,
                    "recommendation": recommendation
                })
        else:
            # Initialize with one empty pointer
            st.session_state.update_conduct_pointers = [
                {"observation": "", "reflection": "", "recommendation": ""}
            ]
# Check if the selected conduct has changed
    if current_selected_conduct != st.session_state.update_conduct_selected_prev:
        # Update the previous selection
        st.session_state.update_conduct_selected_prev = current_selected_conduct
        
        # Re-initialize the pointers based on the newly selected conduct
        existing_pointers = conduct_record.get('pointers', '')
        st.session_state.update_conduct_pointers = []
        
        if existing_pointers:
            # Split pointers by double newlines assuming each pointer is separated by two newlines
            pointer_entries = existing_pointers.split('\n\n')
            for entry in pointer_entries:
                observation = ""
                reflection = ""
                recommendation = ""
                
                # Extract Observation, Reflection, Recommendation using regex
                obs_match = re.search(r'Observation\s*\d*:\s*([\s\S]*?)(?:\n|$)', entry, re.IGNORECASE)
                refl_match = re.search(r'Reflection\s*\d*:\s*([\s\S]*?)(?:\n|$)', entry, re.IGNORECASE)
                rec_match = re.search(r'Recommendation\s*\d*:\s*([\s\S]*?)(?:\n|$)', entry, re.IGNORECASE)
                
                if obs_match:
                    observation = obs_match.group(1).strip()
                if refl_match:
                    reflection = refl_match.group(1).strip()
                if rec_match:
                    recommendation = rec_match.group(1).strip()
                
                st.session_state.update_conduct_pointers.append({
                    "observation": observation,
                    "reflection": reflection,
                    "recommendation": recommendation
                })
        else:
            # Initialize with one empty pointer
            st.session_state.update_conduct_pointers = [
                {"observation": "", "reflection": "", "recommendation": ""}
            ]
    st.subheader("Update Pointers (ORR, Observation, Reflection)")
    for idx, pointer in enumerate(st.session_state.update_conduct_pointers):
        print(idx, pointer)
        st.markdown(f"**Pointer {idx + 1}:**")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.session_state.update_conduct_pointers[idx]["observation"] = st.text_input(
                f"Observation {idx + 1}",
                value=pointer["observation"],
                key=f"update_observation_{idx}"
            )
        with col2:
            st.session_state.update_conduct_pointers[idx]["reflection"] = st.text_input(
                f"Reflection {idx + 1}",
                value=pointer["reflection"],
                key=f"update_reflection_{idx}"
            )
        with col3:
            st.session_state.update_conduct_pointers[idx]["recommendation"] = st.text_input(
                f"Recommendation {idx + 1}",
                value=pointer["recommendation"],
                key=f"update_recommendation_{idx}"
            )
        st.markdown("---")  # Separator between pointers

    # Button to add a new pointer
    st.button("➕ Add Another Pointer", on_click=add_update_pointer)
    if st.button("Load On-Status for Update"):
        platoon = str(st.session_state.conduct_platoon).strip()
        date_str = conduct_record['date']
        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format in selected Conduct.")
            st.stop()

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_allparade_records(selected_company, SHEET_PARADE)

        # Check if we should load from parade state instead of using existing data
        load_from_parade_state = False
        
        # Get the P/T and Outliers values for the selected platoon
        if platoon != "Coy HQ":
            pt_col = int(platoon) + 2  # Column index for P/T PLT1, PLT2, etc.
            outlier_col = int(platoon) + 8  # Column index for PLT1 Outliers, etc.
        else:
            pt_col = 7  # Column for P/T Coy HQ
            outlier_col = 13  # Column for Coy HQ Outliers
        
        try:
            row_num = None
            try:
                # Extract both date and conduct name from the selected_conduct
                conduct_parts = selected_conduct.split(" - ")
                conduct_date = conduct_parts[0].strip()  # Extract the date part
                conduct_name = conduct_parts[1].strip()  # Extract the name part
                
                # Find all instances of the conduct name in column 2
                matching_cells = []
                all_values = SHEET_CONDUCTS.get_all_values()
                
                # Start from row 2 (assuming row 1 is header)
                for row_idx, row_values in enumerate(all_values[1:], start=2):
                    if len(row_values) >= 2 and row_values[1] == conduct_name:
                        # Also check if the date matches (assuming date is in column 1)
                        if len(row_values) >= 1 and row_values[0] == conduct_date:
                            matching_cells.append(row_idx)
                
                if matching_cells:
                    # Use the first matching cell (should be only one if date+name combination is unique)
                    row_num = matching_cells[0]
                    
                    # Log if multiple matches were found (shouldn't happen if dates are unique)
                    if len(matching_cells) > 1:
                        logger.warning(f"Multiple matches found for conduct '{selected_conduct}'. Using the first match.")
            except Exception as e:
                logger.error(f"Error finding conduct row: {e}")
             
            if row_num:
                pt_value = SHEET_CONDUCTS.cell(row_num, pt_col).value
                outliers_value = SHEET_CONDUCTS.cell(row_num, outlier_col).value
                
                # Check if P/T is 0/0 and Outliers is empty or "None"
                if "0/0" in pt_value:
                    is_zero_pt = True
                # is_zero_pt = pt_value == "0/0"
                print(pt_value)
                print(outliers_value)
                is_empty_outliers = not outliers_value or outliers_value.strip().lower() == "none"
                
                if is_zero_pt and is_empty_outliers:
                    load_from_parade_state = True
                    logger.info(f"P/T is {pt_value} and Outliers is '{outliers_value}', loading from parade state.")
        except Exception as e:
            logger.error(f"Error checking P/T and Outliers: {e}")
            # Default to loading from conducts if there's an error
            load_from_parade_state = False

        # Now build the conduct table based on the decision
        if load_from_parade_state:
            # Load data from parade state
            conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)
            st.info("Loading personnel from parade state since no existing data found.")
        else:
            # Load using existing data
            conduct_data = build_fake_conduct_table(platoon, date_obj, records_nominal, records_parade)
            
            # Determine the outlier column for this platoon
            if platoon != "Coy HQ":
                outlier_key = f"plt{platoon} outliers"  # all lower
            else:
                outlier_key = "coy hq outliers"

            # Get the existing outlier string from the conduct_record
            # Normalize the keys in conduct_record to match exactly with outlier_key
            existing_outliers_str = ""
            for key, value in conduct_record.items():
                if key.lower().strip() == outlier_key.lower().strip():
                    existing_outliers_str = value
                    break
            existing_outliers = parse_existing_outliers(existing_outliers_str)
            
            # Helper to find rows in the conduct_data
            def find_in_table(data_list, identifier):
                """
                Returns the row (a dict) if the row's 4D_Number or Name matches `identifier`.
                Otherwise returns None.
                """
                for row in data_list:
                    if row.get("4D_Number", "").lower() == identifier.lower():
                        return row
                    if row.get("Name", "").lower() == identifier.lower():
                        return row
                return None

            # Merge existing outliers into the table
            for _, outlier_info in existing_outliers.items():
                identifier_original = outlier_info["original"]  # e.g. "4D123" or "John Doe"
                status_desc = outlier_info["status_desc"]       # e.g. "MC", "Excused", or ""

                # Check if they already exist in the table
                existing_row = find_in_table(conduct_data, identifier_original)

                if existing_row:
                    # Mark them outlier = True and update status_desc
                    existing_row["Is_Outlier"] = True
                    if status_desc:
                        existing_row["StatusDesc"] = status_desc


        st.session_state.update_conduct_table = conduct_data
        st.success(
            f"Loaded {len(conduct_data)} personnel for Platoon {platoon} from Conduct '{selected_conduct}'."
        )
        logger.info(
            f"Loaded conduct personnel for Platoon {platoon} from Conduct '{selected_conduct}' "
            f"in company '{selected_company}' by user '{st.session_state.username}'."
        )

    if "update_conduct_table" in st.session_state and st.session_state.update_conduct_table:
        #st.subheader(f"Edit Conduct Data for Platoon {st.session_state.conduct_platoon}")
        #st.write("Toggle 'Is_Outlier' if not participating, or add new rows for extra people.")
        sorted_conduct_table = sorted(st.session_state.update_conduct_table, 
                                 key=lambda x: "ZZZ" if x.get("Rank", "").upper() == "REC" else x.get("Rank", ""))
        st.write("In order to update, make sure correct platoon chosen and then press load on status for the table to reflect correct platoon. Hence, whenever changing platoon make sure to press load after that to reflect accordingly.")
        edited_data = st.data_editor(
            st.session_state.update_conduct_table,
            num_rows="fixed",
            hide_index=True,
        )
    else:
        edited_data = None

    if st.button("Update Conduct Data") and edited_data is not None:
        rows_updated = 0
        platoon = str(st.session_state.conduct_platoon).strip()
        pt_field = f"P/T PLT{platoon}"
        new_participating = sum([1 for row in edited_data if not row.get('Is_Outlier', False)])
        new_total = len(edited_data)
        new_outliers = []
        pointers_list = []

        SHEET_EVERYTHING = worksheets["everything"]
        # Extract the updated attendance data

        # Ensure the date is in DDMMYYYY format
        formatted_date_str = ensure_date_str(conduct_record['date'])

        attendance_data = extract_attendance_data(edited_data)
        update_conduct_column_everything(
            SHEET_EVERYTHING,
            formatted_date_str,
            conduct_record['conduct_name'],
            attendance_data
        )



        for idx, pointer in enumerate(st.session_state.update_conduct_pointers, start=1):
            observation = pointer.get("observation", "").strip()
            reflection = pointer.get("reflection", "").strip()
            recommendation = pointer.get("recommendation", "").strip()

            pointer_str = ""
            if observation:
                pointer_str += f"Observation {idx}:\n{observation}\n"
            if reflection:
                pointer_str += f"Reflection {idx}:\n{reflection}\n"
            if recommendation:
                pointer_str += f"Recommendation {idx}:\n{recommendation}\n"
            pointers_list.append(pointer_str.strip())

        new_pointers = "\n\n".join(pointers_list)

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_allparade_records(selected_company, SHEET_PARADE)

        def is_valid_4d(four_d_number):
            """
            Validates the 4D number format.
            Adjust the regex pattern as per your specific requirements.
            """
            pattern = re.compile(r'^4D\d{3,4}$', re.IGNORECASE)
            return four_d_number if pattern.match(four_d_number) else ''

        def ensure_str(value):
            """Ensures the value is a string."""
            return str(value) if value is not None else ''

        def parse_existing_outliers(existing_outliers):
            """
            Parses the existing_outliers string into a dictionary.
            
            Returns:
                dict: Mapping from identifier (4D_Number or Name) to status_desc.
            """
            pattern = re.compile(
                r'(4D\d{3,4}|[A-Za-z]+(?:\s[A-Za-z]+)*)\s*(?:\(([^)]+)\))?',
                re.IGNORECASE
            )
            outliers_dict = {}
            for match in pattern.finditer(existing_outliers):
                identifier = match.group(1).strip()
                status_desc = match.group(2).strip() if match.group(2) else ''
                outliers_dict[identifier.lower()] = {
                    'original': identifier,
                    'status_desc': status_desc
                }
            return outliers_dict

        def reconstruct_outliers(outliers_dict, edited_data):
            """
            Reconstructs the outliers string from the dictionary with names included.
            
            Returns:
                str: Comma-separated outliers with names.
            """
            outliers_list = []
            
            # Helper function to find a person's name by 4D number
            def find_name_by_4d(four_d, data):
                for row in data:
                    if row.get("4D_Number", "").lower() == four_d.lower():
                        return row.get("Name", "")
                return ""
            
            for entry in outliers_dict.values():
                identifier = entry['original']
                status_desc = entry['status_desc']
                
                # Check if identifier is a 4D number
                if re.match(r'^4D\d{3,4}$', identifier, re.IGNORECASE):
                    # Find the person's name
                    name = find_name_by_4d(identifier, edited_data)
                    if name:
                        formatted_entry = f"{identifier} {name}"
                    else:
                        formatted_entry = identifier
                else:
                    # Identifier is already a name
                    formatted_entry = identifier
                
                # Add status description if available
                if status_desc:
                    formatted_entry += f" ({status_desc})"
                    
                outliers_list.append(formatted_entry)
                    
            return ", ".join(outliers_list) if outliers_list else "None"

        def update_outliers(edited_data, conduct_record, platoon):
            # Determine which outlier column key to use based on platoon
            if platoon != "Coy HQ":
                outlier_key = f"PLT{platoon} Outliers"
            else:
                outlier_key = "Coy HQ Outliers"
                
            existing_outliers_str = ensure_str(conduct_record.get(outlier_key, ''))
            existing_outliers = parse_existing_outliers(existing_outliers_str)
            
            processed_identifiers = set()
            
            for row in edited_data:
                four_d = is_valid_4d(row.get("4D_Number", ""))
                name = ensure_str(row.get("Name", ""))
                status_desc = ensure_str(row.get("StatusDesc", ""))
                is_outlier = row.get("Is_Outlier", False)
                
                identifier = four_d if four_d else name
                identifier_key = identifier.lower()
                
                if is_outlier:
                    processed_identifiers.add(identifier_key)
                    if identifier_key in existing_outliers:
                        if existing_outliers[identifier_key]['status_desc'] != status_desc:
                            existing_outliers[identifier_key]['status_desc'] = status_desc
                    else:
                        existing_outliers[identifier_key] = {
                            'original': identifier,
                            'status_desc': status_desc
                        }
                else:
                    if identifier_key in existing_outliers:
                        del existing_outliers[identifier_key]
            
            updated_outliers = reconstruct_outliers(existing_outliers, edited_data)
            return updated_outliers
        updated_outliers = update_outliers(edited_data, conduct_record, platoon)

        new_pt_value = f"{new_participating}/{new_total}"
        # Add after the attendance_data extraction but before the sheet updates
# Initialize recruit and commander counts
        rec_counts = 0
        cmd_counts = 0

        # Count participating recruits and commanders
        for row in edited_data:
            if not row.get('Is_Outlier', False):
                if row.get('Rank', '').upper() == 'REC':
                    rec_counts += 1
                else:
                    cmd_counts += 1

        # Get total counts from nominal roll for this platoon
        #records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        rec_totals = 0
        cmd_totals = 0

        for person in records_nominal:
            plt = person.get("platoon", "")
            if plt == platoon:  # Only count for the selected platoon
                if person.get("rank", "").upper() == "REC":
                    rec_totals += 1
                else:
                    cmd_totals += 1

        # Format the detailed P/T value
        new_pt_value = f"REC: {rec_counts}/{rec_totals}\nCMD: {cmd_counts}/{cmd_totals}\nTOTAL: {new_participating}/{new_total}"

        try:
            # Extract both date and conduct name from the selected_conduct
            conduct_parts = selected_conduct.split(" - ")
            conduct_date = conduct_parts[0].strip()  # Extract the date part
            conduct_name = conduct_parts[1].strip()  # Extract the name part
            
            # Find all instances of the conduct name in column 2
            matching_cells = []
            all_values = SHEET_CONDUCTS.get_all_values()
            
            # Start from row 2 (assuming row 1 is header)
            for row_idx, row_values in enumerate(all_values[1:], start=2):
                if len(row_values) >= 2 and row_values[1] == conduct_name:
                    # Also check if the date matches (assuming date is in column 1)
                    if len(row_values) >= 1 and row_values[0] == conduct_date:
                        matching_cells.append(row_idx)
            
            if not matching_cells:
                st.error("Conduct not found in the sheet.")
                logger.error(f"Conduct '{selected_conduct}' not found in the sheet.")
                st.stop()
                
            # Use the first matching cell (should be only one if date+name combination is unique)
            row_number = matching_cells[0]
            
            # Log if multiple matches were found (shouldn't happen if dates are unique)
            if len(matching_cells) > 1:
                logger.warning(f"Multiple matches found for conduct '{selected_conduct}'. Using the first match.")
                
        except Exception as e:
            st.error(f"Error locating Conduct in the sheet: {e}")
            logger.error(f"Exception while locating Conduct '{selected_conduct}': {e}")
            st.stop()
        try:
            # Determine the column index based on the selected platoon
            if platoon != "Coy HQ":
                # For Platoons 1-4, map to columns 3-6 respectively
                platoon_num = int(platoon)  # Convert platoon to integer
                column_index = 2 + platoon_num  # Platoon 1 -> Column 3, Platoon 2 -> Column 4, etc.
                pt_field = f"P/T PLT{platoon_num}"
            else:
                # For 'Coy HQ', assume it maps to Column 7
                column_index = 7
                pt_field = "P/T Coy HQ"
            
            # Update the specified cell in the Google Sheet
            SHEET_CONDUCTS.update_cell(row_number, column_index, new_pt_value)
            
            # Log the update action
            logger.info(
                f"Updated {pt_field} to {new_pt_value} for conduct '{selected_conduct}' "
                f"in company '{selected_company}' by user '{st.session_state.username}'."
            )
            
        except ValueError:
            # Handle the case where platoon is not an integer and not 'Coy HQ'
            st.error(f"Invalid platoon selection: '{platoon}'. Please select a valid platoon.")
            logger.error(f"Invalid platoon selection: '{platoon}'. Update aborted.")
            st.stop()
            
        except Exception as e:
            # Handle other potential exceptions
            st.error(f"An error occurred while updating the conduct: {e}")
            logger.error(f"Exception while updating conduct '{selected_conduct}': {e}")
            st.stop()

        try:
            if platoon != "Coy HQ":
                # For platoon "1" through "4"
                outlier_column_index = 8 + int(platoon)  # e.g., platoon "1": 8+1 = 9
            else:
                outlier_column_index = 13  # Coy HQ

            SHEET_CONDUCTS.update_cell(row_number, outlier_column_index, updated_outliers if updated_outliers else "None")

            #SHEET_CONDUCTS.update_cell(row_number, 9, updated_outliers if updated_outliers else "None")
            logger.info(
                f"Updated Outliers to '{updated_outliers}' for conduct '{selected_conduct}' "
                f"in company '{selected_company}' by user '{st.session_state.username}'."
            )
        except Exception as e:
            st.error(f"Error updating Outliers: {e}")
            logger.error(f"Exception while updating Outliers: {e}")
            st.stop()

        if new_pointers:
            try:
                SHEET_CONDUCTS.update_cell(row_number, 14, new_pointers)
                logger.info(
                    f"Updated Pointers to '{new_pointers}' for conduct '{selected_conduct}' "
                    f"in company '{selected_company}' by user '{st.session_state.username}'."
                )
            except Exception as e:
                st.error(f"Error updating Pointers: {e}")
                logger.error(f"Exception while updating Pointers for conduct '{selected_conduct}': {e}")
                st.stop()

        try:
            # Get all platoon participation data
            pt1 = SHEET_CONDUCTS.cell(row_number, 3).value
            pt2 = SHEET_CONDUCTS.cell(row_number, 4).value
            pt3 = SHEET_CONDUCTS.cell(row_number, 5).value
            pt4 = SHEET_CONDUCTS.cell(row_number, 6).value
            pt5 = SHEET_CONDUCTS.cell(row_number, 7).value
            
            platoon_values = [pt1, pt2, pt3, pt4, pt5]
            
            # Initialize counters
            total_rec_part = 0
            total_rec = 0
            total_cmd_part = 0 
            total_cmd = 0
            
            # Process each platoon's data
            for pt in platoon_values:
                lines = pt.split('\n')
                if len(lines) >= 3:  # Check if we have the detailed format
                    # Parse REC line
                    rec_line = lines[0]
                    if rec_line.startswith("REC:"):
                        rec_parts = rec_line.replace("REC:", "").strip().split('/')
                        if len(rec_parts) == 2 and rec_parts[0].isdigit() and rec_parts[1].isdigit():
                            total_rec_part += int(rec_parts[0])
                            total_rec += int(rec_parts[1])
                    
                    # Parse CMD line
                    cmd_line = lines[1]
                    if cmd_line.startswith("CMD:"):
                        cmd_parts = cmd_line.replace("CMD:", "").strip().split('/')
                        if len(cmd_parts) == 2 and cmd_parts[0].isdigit() and cmd_parts[1].isdigit():
                            total_cmd_part += int(cmd_parts[0])
                            total_cmd += int(cmd_parts[1])
            
            # Calculate the overall total
            total_part = total_rec_part + total_cmd_part
            total_strength = total_rec + total_cmd
            
            # Format the company-wide P/T total in the detailed format
            pt_total = f"REC: {total_rec_part}/{total_rec}\nCMD: {total_cmd_part}/{total_cmd}\nTOTAL: {total_part}/{total_strength}"
            
            # Update the P/T Total cell
            SHEET_CONDUCTS.update_cell(row_number, 8, pt_total)
            
        except Exception as e:
            st.error(f"Error calculating/updating P/T Total: {e}")
            logger.error(f"Exception while calculating/updating P/T Total for conduct '{selected_conduct}': {e}")
            st.stop()

        st.success(f"Conduct '{selected_conduct}' updated successfully.")
        logger.info(
            f"Conduct '{selected_conduct}' updated successfully in company '{selected_company}' "
            f"by user '{st.session_state.username}'."
        )


        # Optionally, clear the conduct table if desired
elif feature == "Update Parade":
    st.header("Update Parade State")

    st.session_state.parade_platoon = st.selectbox(
        "Platoon for Parade Update:",
        options=[1, 2, 3, 4, "Coy HQ"],
        format_func=lambda x: str(x)
    )

    submitted_by = st.session_state.username

    if st.button("Load Personnel"):
        platoon = str(st.session_state.parade_platoon).strip()
        if not platoon:
            st.error("Please select a valid platoon.")
            st.stop()

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        data = get_company_personnel(platoon, records_nominal, records_parade)
        st.session_state.parade_table = data
        st.info(f"Loaded {len(data)} personnel for Platoon {platoon} in company '{selected_company}'.")
        logger.info(f"Loaded personnel for Platoon {platoon} in company '{selected_company}' by user '{submitted_by}'.")

        current_statuses = [
            row for row in records_parade
            if normalize_name(row.get('platoon', '')) == normalize_name(platoon)
        ]
        if current_statuses:
            st.subheader("Current Parade Status")
            formatted_statuses = []
            for status in current_statuses:
                formatted_statuses.append({
                    "Name": status.get("name", ""),
                    "Platoon": status.get("platoon", ""),
                    "Status": status.get("status", ""),
                    "Start_Date": status.get("start_date_ddmmyyyy", ""),
                    "End_Date": status.get("end_date_ddmmyyyy", "")
                })
            #st.write(formatted_statuses)
            logger.info(
                f"Displayed current parade statuses for platoon {platoon} in company '{selected_company}' "
                f"by user '{submitted_by}'."
            )

    if st.session_state.parade_table:
        st.subheader("Edit Parade Data, Then Click 'Update'")
        st.write("Fill in 'Status', 'Start_Date (DDMMYYYY)', 'End_Date (DDMMYYYY)'")
        st.write("To delete an existing status, please delete the values in 'Status', 'Start_Date (DDMMYYYY)', 'End_Date (DDMMYYYY)' only.")
        sorted_conduct_table = sorted(st.session_state.parade_table, 
                                 key=lambda x: "ZZZ" if x.get("Rank", "").upper() == "REC" else x.get("Rank", ""))
        edited_data = st.data_editor(
            st.session_state.parade_table,
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Name": st.column_config.TextColumn("Name", disabled=True),
                "4D_Number": st.column_config.TextColumn("4D_Number", disabled=True),
                "Rank": st.column_config.TextColumn("Rank", disabled=True),
                "Number_of_Leaves_Left": st.column_config.TextColumn("Number_of_Leaves_Left", disabled=True),
                "Dates_Taken": st.column_config.TextColumn("Dates_Taken", disabled=True),
                "_row_num": st.column_config.TextColumn("_row_num", disabled=True),

            }
        )
    else:
        edited_data = None

    if st.button("Update Parade State") and edited_data is not None:
        rows_updated = 0
        platoon = str(st.session_state.parade_platoon).strip()

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        # Initialize lists to collect batch requests for each sheet
        delete_requests = []      # For all deletions in Parade_State
        update_requests = []      # For updates in Parade_State
        nominal_requests = []     # For updates to the Nominal_Roll (leaves)
        append_rows = []          # For any new rows to be appended to Parade_State

        # Retrieve the header to figure out column indices for updates
        try:
            header = SHEET_PARADE.row_values(1)
            header = [h.strip().lower() for h in header]
            name_col = header.index("name") + 1
            status_col = header.index("status") + 1
            start_date_col = header.index("start_date_ddmmyyyy") + 1
            end_date_col = header.index("end_date_ddmmyyyy") + 1
            submitted_by_col = header.index("submitted_by") + 1 if "submitted_by" in header else None
        except ValueError as ve:
            st.error(f"Required column missing in Parade_State: {ve}. Cannot proceed.")
            logger.error(f"Required column missing in Parade_State: {ve} in company '{selected_company}'.")
            st.stop()

        # Process each row from the data editor
        for idx, row in enumerate(edited_data):
            name_val = ensure_str(row.get("Name", "")).strip()
            status_val = ensure_str(row.get("Status", "")).strip()
            start_val = ensure_str(row.get("Start_Date", "")).strip()
            end_val = ensure_str(row.get("End_Date", "")).strip()
            four_d = is_valid_4d(row.get("4D_Number", ""))

            rank = ensure_str(row.get("Rank", "")).strip()
            parade_entry = st.session_state.parade_table[idx]
            row_num = parade_entry.get('_row_num')  # Existing row number (if any)

            # 1) If all key fields are empty on an existing row -> schedule deletion.
            if row_num and not status_val and not start_val and not end_val:
                delete_requests.append({
                    'deleteDimension': {
                        'range': {
                            'sheetId': SHEET_PARADE.id,
                            'dimension': 'ROWS',
                            'startIndex': row_num - 1,  # 0-indexed
                            'endIndex': row_num
                        }
                    }
                })
                logger.info(
                    f"Scheduled deletion of Parade_State row {row_num} for {name_val} in company '{selected_company}'."
                )
                rows_updated += 1
                continue

            # Ensure the name is provided if we are not deleting
            if not name_val:
                st.error(f"Name is required for row {idx}. Skipping.")
                logger.error(f"Name missing for row {idx} in company '{selected_company}'.")
                continue

            # 2) If an existing row has no status -> schedule deletion.
            if row_num and not status_val:
                delete_requests.append({
                    'deleteDimension': {
                        'range': {
                            'sheetId': SHEET_PARADE.id,
                            'dimension': 'ROWS',
                            'startIndex': row_num - 1,
                            'endIndex': row_num
                        }
                    }
                })
                logger.info(
                    f"Scheduled deletion of Parade_State row {row_num} for {name_val} in company '{selected_company}'."
                )
                rows_updated += 1
                continue

            # 3) If row has partial but not enough info (missing status/date), skip.
            if (status_val and (not start_val or not end_val)) or ((start_val or end_val) and not status_val):
                st.error(f"Missing fields (Status/Start/End) for {name_val}. Skipping.")
                logger.error(f"Missing fields for {name_val} in company '{selected_company}'. Skipping.")
                continue

            # If there's no status, start, end, it's a new blank row. Skip if no changes:
            if not row_num and not status_val and not start_val and not end_val:
                continue

            # 4) Validate dates if present
            formatted_start_val = ""
            formatted_end_val = ""
            if start_val and end_val:
                formatted_start_val = ensure_date_str(start_val)
                formatted_end_val = ensure_date_str(end_val)
                try:
                    start_dt = datetime.strptime(formatted_start_val, "%d%m%Y")
                    end_dt = datetime.strptime(formatted_end_val, "%d%m%Y")
                    if end_dt < start_dt:
                        st.error(f"End date is before start date for {name_val}. Skipping.")
                        logger.error(f"End date before start date for {name_val} in company '{selected_company}'.")
                        continue
                except ValueError:
                    st.error(f"Invalid date(s) for {name_val}, skipping.")
                    logger.error(
                        f"Invalid date format for {name_val}: Start={formatted_start_val}, End={formatted_end_val} "
                        f"in company '{selected_company}'."
                    )
                    continue

            # 5) If the status indicates leave, schedule the Nominal_Roll updates.
            leave_pattern = re.compile(r'\b(?:leave|ll|ol)\b', re.IGNORECASE)
            if leave_pattern.search(status_val):
                half_day = "(am)" in status_val.lower() or "(pm)" in status_val.lower()
                dates_str = (
                    f"{formatted_start_val}-{formatted_end_val}" 
                    if (formatted_start_val and formatted_end_val and formatted_start_val != formatted_end_val) 
                    else formatted_start_val
                )
                logger.debug(f"Constructed dates_str for {name_val}: {dates_str}")

                # Attempt to find a matching Nominal Record
                try:
                    nominal_record = SHEET_NOMINAL.find(name_val, in_column=2)
                except Exception as e:
                    st.error(f"Error finding {name_val} in Nominal_Roll: {e}. Skipping.")
                    logger.error(f"Exception while finding {name_val} in Nominal_Roll: {e}.")
                    continue

                if not nominal_record:
                    st.error(f"{name_val}/{four_d} not found in Nominal_Roll. Skipping.")
                    logger.error(f"{name_val}/{four_d} not found in Nominal_Roll in company '{selected_company}'.")
                    continue

                existing_dates = SHEET_NOMINAL.cell(nominal_record.row, 6).value
                if is_leave_accounted(existing_dates, dates_str):
                    logger.info(
                        f"Leave on {dates_str} for {name_val}/{four_d} already accounted for in company '{selected_company}'. Skipping."
                    )
                    continue

                leaves_used = 0.5 if half_day else calculate_leaves_used(dates_str)
                if leaves_used <= 0:
                    st.error(f"Invalid leave duration for {name_val}, skipping.")
                    logger.error(f"Invalid leave duration for {name_val}: {dates_str} in company '{selected_company}'.")
                    continue

                if has_overlapping_status(four_d, start_dt, end_dt, records_parade):
                    st.error(f"Leave dates overlap with existing record for {name_val}. Skipping.")
                    logger.error(f"Leave dates overlap for {name_val}: {dates_str} in company '{selected_company}'.")
                    continue

                # Update Leaves
                try:
                    current_leaves_left = SHEET_NOMINAL.cell(nominal_record.row, 5).value
                    try:
                        current_leaves_left = float(current_leaves_left)
                    except ValueError:
                        current_leaves_left = 14
                        logger.warning(f"Invalid 'Number of Leaves Left' for {name_val}/{four_d}. Resetting to 14.")

                    if leaves_used > current_leaves_left:
                        st.error(
                            f"{name_val}/{four_d} does not have enough leaves left. "
                            f"Available: {current_leaves_left}, Requested: {leaves_used}. Skipping."
                        )
                        logger.error(
                            f"{name_val}/{four_d} insufficient leaves. Available: {current_leaves_left}, Requested: {leaves_used}."
                        )
                        continue

                    new_leaves_left = current_leaves_left - leaves_used
                    # Nominal_Roll 'Number of Leaves Left' update
                    nominal_requests.append({
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_NOMINAL.id,
                                'startRowIndex': nominal_record.row - 1,
                                'endRowIndex': nominal_record.row,
                                'startColumnIndex': 4,  # Column E (0-indexed)
                                'endColumnIndex': 5,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'numberValue': new_leaves_left}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    })

                    # Nominal_Roll 'Dates Taken' update
                    new_dates_entry = dates_str
                    if existing_dates:
                        existing_dates = existing_dates.strip()
                        if existing_dates and not existing_dates.endswith(','):
                            updated_dates = f"{existing_dates},{new_dates_entry}"
                        else:
                            updated_dates = f"{existing_dates}{new_dates_entry}"
                    else:
                        updated_dates = new_dates_entry

                    nominal_requests.append({
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_NOMINAL.id,
                                'startRowIndex': nominal_record.row - 1,
                                'endRowIndex': nominal_record.row,
                                'startColumnIndex': 5,  # Column F (0-indexed)
                                'endColumnIndex': 6,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'stringValue': updated_dates}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    })
                    logger.info(
                        f"Scheduled updates for Nominal_Roll for {name_val}/{four_d}: "
                        f"New leaves left = {new_leaves_left}, Dates = {updated_dates}."
                    )
                except Exception as e:
                    st.error(f"Error updating leaves for {name_val}/{four_d}: {e}. Skipping.")
                    logger.error(f"Exception while updating leaves for {name_val}/{four_d}: {e}.")
                    continue

            # 6) Build the batch update requests for the Parade_State if this row is existing:
            if row_num:
                # Compare with original to see if changed
                original_entry = st.session_state.parade_table[idx]
                is_changed = (
                    row.get('Status', '') != original_entry.get('Status', '') or
                    row.get('Start_Date', '') != original_entry.get('Start_Date', '') or
                    row.get('End_Date', '') != original_entry.get('End_Date', '')
                )

                # Prepare separate "updateCells" requests for each column
                # (Name, Status, Start, End) to the same row.
                update_requests.extend([
                    # Update "Name"
                    {
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_PARADE.id,
                                'startRowIndex': row_num - 1,
                                'endRowIndex': row_num,
                                'startColumnIndex': name_col - 1,
                                'endColumnIndex': name_col,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'stringValue': name_val}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    },
                    # Update "Status"
                    {
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_PARADE.id,
                                'startRowIndex': row_num - 1,
                                'endRowIndex': row_num,
                                'startColumnIndex': status_col - 1,
                                'endColumnIndex': status_col,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'stringValue': status_val}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    },
                    # Update "Start_Date"
                    {
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_PARADE.id,
                                'startRowIndex': row_num - 1,
                                'endRowIndex': row_num,
                                'startColumnIndex': start_date_col - 1,
                                'endColumnIndex': start_date_col,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'stringValue': formatted_start_val}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    },
                    # Update "End_Date"
                    {
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_PARADE.id,
                                'startRowIndex': row_num - 1,
                                'endRowIndex': row_num,
                                'startColumnIndex': end_date_col - 1,
                                'endColumnIndex': end_date_col,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'stringValue': formatted_end_val}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    }
                ])

                # If status/dates changed, update "Submitted_By"
                if submitted_by_col and is_changed:
                    update_requests.append({
                        'updateCells': {
                            'range': {
                                'sheetId': SHEET_PARADE.id,
                                'startRowIndex': row_num - 1,
                                'endRowIndex': row_num,
                                'startColumnIndex': submitted_by_col - 1,
                                'endColumnIndex': submitted_by_col,
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredValue': {'stringValue': submitted_by}
                                }]
                            }],
                            'fields': 'userEnteredValue'
                        }
                    })

                rows_updated += 1

            else:
                # This is a new entry to be appended
                new_row = [
                    platoon,
                    rank,
                    name_val,
                    four_d,
                    status_val,
                    formatted_start_val,
                    formatted_end_val,
                    submitted_by
                ]
                append_rows.append(new_row)
                rows_updated += 1

        # =======================
        # Execute the batched operations in a safe order
        # =======================

        # 1) Nominal Roll updates (independent of row references in Parade sheet)
        if nominal_requests:
            SHEET_PARADE.spreadsheet.batch_update({"requests": nominal_requests})

        # 2) Parade updates (existing rows only)
        if update_requests:
            SHEET_PARADE.spreadsheet.batch_update({"requests": update_requests})

        # 3) Deletions in descending order, so row shifts do not break references
        if delete_requests:
            # Sort by 'startIndex' descending so we delete from bottom to top
            delete_requests = sorted(
                delete_requests,
                key=lambda r: r['deleteDimension']['range']['startIndex'],
                reverse=True
            )
            SHEET_PARADE.spreadsheet.batch_update({"requests": delete_requests})

        # 4) Append brand-new rows
        if append_rows:
            SHEET_PARADE.append_rows(append_rows, value_input_option='USER_ENTERED')

        st.success("Parade State updated.")
        logger.info(
            f"Parade State updated for {rows_updated} row(s) for platoon {platoon} in company '{selected_company}' "
            f"by user '{submitted_by}'."
        )

        # Reset the session state
        st.session_state.parade_platoon = 1
        st.session_state.parade_table = []


# ------------------------------------------------------------------------------
# 11) Feature D: Queries with Multiple Options
# ------------------------------------------------------------------------------
elif feature == "Queries":
    st.subheader("Query Person Information")
    
    # Add tabs for different query types
    tab1, tab2, tab3, tab4, tab5,tab6 = st.tabs(["Medical Statuses", "Leave Counter", "MC Counter", "Threshold Alerts" ,"Ration Requirements", "Daily Attendance"])
    
    # Common input field for all tabs
    person_input = st.text_input("Enter the 4D Number or partial Name", key="query_person_input")
    search_button = st.button("Search", key="btn_query_person")
    def parse_ddmmyyyy(d):
        try:
            return datetime.strptime(str(d), "%d%m%Y")
        except ValueError:
            return datetime.min
    if search_button:
        if not person_input:
            st.error("Please enter a 4D Number or Name.")
            st.stop()
            
        # Common data fetching for all tabs
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_allparade_records(selected_company, SHEET_PARADE)
        parade_data = records_parade

        # Check if input is a valid 4D
        four_d_input_clean = is_valid_4d(person_input)
        person_rows = []

        if four_d_input_clean:
            # Search by 4D
            person_rows = [
                row for row in parade_data
                if is_valid_4d(row.get("4d_number", "")) == four_d_input_clean
            ]
        else:
            # Otherwise, partial match by name (case-insensitive)
            person_input_lower = person_input.strip().lower()
            person_rows = [
                row for row in parade_data
                if person_input_lower in row.get("name", "").strip().lower()
            ]

        if not person_rows:
            st.warning(f"No Parade_State records found for '{person_input}'")
            logger.info(f"No Parade_State records found for '{person_input}' in company '{selected_company}'.")
        else:
            def parse_ddmmyyyy(d):
                try:
                    return datetime.strptime(str(d), "%d%m%Y")
                except ValueError:
                    return datetime.min
                    
            # Get person details for display in all tabs
            four_d_val = ""
            name_val = ""
            rank_val = ""
            
            if person_rows:
                first_row = person_rows[0]
                four_d_val = is_valid_4d(first_row.get("4d_number", "")) or ""
                
                # Try to get more details from nominal roll
                if four_d_val:
                    for nominal in records_nominal:
                        if nominal['4d_number'].upper() == four_d_val.upper():
                            name_val = nominal['name']
                            rank_val = nominal['rank']
                            break
                
                # If we don't have a name yet, use the one from parade state
                if not name_val:
                    name_val = first_row.get("name", "")
            
            # TAB 1: MEDICAL STATUSES (Original Functionality)
            with tab1:
                st.subheader(f"Medical Statuses for {rank_val} {name_val} ({four_d_val})")
                
                valid_status_prefixes = ("ex", "rib", "ld", "mc", "ml")
                filtered_person_rows = [
                    row for row in person_rows if row.get("status", "").lower().startswith(valid_status_prefixes)
                ]
                filtered_person_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("start_date_ddmmyyyy", "")))

                enhanced_rows = []
                for row in filtered_person_rows:
                    # Grab 4D or empty
                    row_four_d = is_valid_4d(row.get("4d_number", "")) or ""
                    # Grab rank from nominal if possible
                    row_rank = ""
                    row_name = ""
                    
                    # We can look up 4D if it exists, or match by name if not
                    if row_four_d:
                        # If 4D is valid, match by 4D
                        for nominal in records_nominal:
                            if nominal['4d_number'].upper() == row_four_d.upper():
                                row_rank = nominal['rank']
                                break
                        row_name = find_name_by_4d(row.get("4d_number", ""), records_nominal)
                    else:
                        # If no 4D, match by name partially
                        name_from_parade = ensure_str(row.get("name", ""))
                        for nominal in records_nominal:
                            # For partial match, ensure the parade name is quite specific
                            # but here we'll just do a direct equality ignoring case
                            if nominal['name'].strip().lower() == name_from_parade.strip().lower():
                                row_four_d = nominal['4d_number']
                                row_rank = nominal['rank']
                                break
                        row_name = name_from_parade

                    enhanced_rows.append({
                        "Rank": row_rank,
                        "Name": row_name,
                        "4D_Number": row_four_d,
                        "Status": row.get("status", ""),
                        "Start_Date": row.get("start_date_ddmmyyyy", ""),
                        "End_Date": row.get("end_date_ddmmyyyy", "")
                    })

                if enhanced_rows:
                    st.table(enhanced_rows)
                else:
                    st.info("No medical status records found")

            # TAB 2: LEAVE COUNTER
            with tab2:
                st.subheader(f"Leave Status for {rank_val} {name_val} ({four_d_val})")
                
                # Filter for leave-related statuses
                leave_prefixes = ("ll", "ol", "leave")
                leave_rows = [
                    row for row in person_rows if any(row.get("status", "").lower().startswith(prefix) for prefix in leave_prefixes)
                ]
                leave_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("start_date_ddmmyyyy", "")))

                # Calculate total leaves taken and remaining
                total_leave_days = 0
                leave_details = []
                
                for row in leave_rows:
                    start_date = parse_ddmmyyyy(row.get("start_date_ddmmyyyy", ""))
                    end_date = parse_ddmmyyyy(row.get("end_date_ddmmyyyy", ""))
                    
                    # If end_date is valid, calculate duration
                    if end_date != datetime.min and start_date != datetime.min:
                        duration = (end_date - start_date).days + 1  # inclusive of both start and end dates
                        total_leave_days += duration
                    else:
                        duration = "Unknown"  # Handle case where dates are missing
                    
                    leave_details.append({
                        "Status": row.get("status", ""),
                        "Start_Date": row.get("start_date_ddmmyyyy", ""),
                        "End_Date": row.get("end_date_ddmmyyyy", ""),
                        "Duration": duration if isinstance(duration, str) else f"{duration} days"
                    })
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Leaves Taken", f"{total_leave_days} days")
                with col2:
                    st.metric("Annual Leave Entitled", "14 days")
                with col3:
                    remaining_leaves = 14 - total_leave_days
                    st.metric("Remaining Leaves", f"{max(0, remaining_leaves)} days", 
                              delta=f"{'-' if remaining_leaves < 0 else ''}{abs(remaining_leaves)}" if remaining_leaves != 14 else None)
                
                # Display leave history
                if leave_details:
                    st.subheader("Leave History")
                    st.table(leave_details)
                else:
                    st.info("No leave records found")

            # TAB 3: MC COUNTER
            with tab3:
                st.subheader(f"Medical Status for {rank_val} {name_val} ({four_d_val})")
                
                # Filter for medical-related statuses
                medical_prefixes = ("mc", "ml")
                medical_rows = [
                    row for row in person_rows if any(row.get("status", "").lower().startswith(prefix) for prefix in medical_prefixes)
                ]
                medical_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("start_date_ddmmyyyy", "")))

                # Calculate total medical leave days
                total_mc_days = 0
                total_ml_days = 0
                medical_details = []
                
                for row in medical_rows:
                    start_date = parse_ddmmyyyy(row.get("start_date_ddmmyyyy", ""))
                    end_date = parse_ddmmyyyy(row.get("end_date_ddmmyyyy", ""))
                    status = row.get("status", "").lower()
                    
                    # If end_date is valid, calculate duration
                    if end_date != datetime.min and start_date != datetime.min:
                        duration = (end_date - start_date).days + 1  # inclusive of both start and end dates
                        
                        # Count by type
                        if status.startswith("mc"):
                            total_mc_days += duration
                        elif status.startswith("ml"):
                            total_ml_days += duration
                    else:
                        duration = "Unknown"  # Handle case where dates are missing
                    
                    medical_details.append({
                        "Status": row.get("status", ""),
                        "Start_Date": row.get("start_date_ddmmyyyy", ""),
                        "End_Date": row.get("end_date_ddmmyyyy", ""),
                        "Duration": duration if isinstance(duration, str) else f"{duration} days"
                    })

                
                total_medical = total_mc_days + total_ml_days
                st.metric("Total MC/ML", f"{total_medical} days")
                
                # Display medical history
                if medical_details:
                    st.subheader("Medical History")
                    st.table(medical_details)
                else:
                    st.info("No medical records found")
    with tab4:
        st.subheader("📋 MR & Report Sick Threshold Alerts")
        st.write("MR (Medical Reporting) Threshold: Counts every calendar day covered by MC or ML within a rolling 30‑day window (default: the last 30 days including today).")
        st.write("RSO Threshold: Once they hit the MR threshold and they have 3 or more MC/ML periods tagged \"RSO\"")

        # ── date‑range picker ─────────────────────────────────────────────
        today = datetime.today().date()
        default_start = today - timedelta(days=29)             # NEW
        start_date, end_date = st.date_input(                 # NEW
            "MR rolling‑window (inclusive)",
            (default_start, today),
            key="mr_window"
        )
        if start_date > end_date:
            st.error("Start date must be on or before end date.")
            st.stop()

        # pull data …
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        all_parade_records = get_allparade_records(selected_company, SHEET_PARADE)

        nominal_lookup = {rec['4d_number'].upper(): rec for rec in records_nominal}
        persons_by_4d = defaultdict(list)
        for row in all_parade_records:
            four_d = is_valid_4d(row.get("4d_number", ""))
            if four_d:
                persons_by_4d[four_d].append(row)

        flagged_persons = []
        for four_d, rows in persons_by_4d.items():
            mc_ml_days = set()                                # NEW use set to avoid dup days
            weekend_sick_count = 0

            for row in rows:
                status = row.get("status", "").strip().lower()
                if not status.startswith(("mc", "ml")):
                    continue

                # parse dates
                sdt = parse_ddmmyyyy(row.get("start_date_ddmmyyyy", ""))
                edt = parse_ddmmyyyy(row.get("end_date_ddmmyyyy", ""))
                if sdt == datetime.min or edt == datetime.min:
                    continue

                # iterate through covered days
                cur = sdt.date()
                while cur <= edt.date():
                    if start_date <= cur <= end_date:         # NEW window filter
                        mc_ml_days.add(cur)
                    cur += timedelta(days=1)

                # count RSO on Fri‑Sun only inside window      # NEW
                if status.startswith(("mc(rso)", "ml(rso)", "mc (rso)", "ml (rso)")):
                    if start_date <= sdt.date() <= end_date:
                        weekend_sick_count += 1

            # MR threshold = 8+ days in window
            if len(mc_ml_days) >= 8:
                nom = nominal_lookup.get(four_d.upper(), {})
                flagged_persons.append({
                    "4D Number": four_d,
                    "Rank": nom.get("rank", ""),
                    "Name": nom.get("name", ""),
                    "MR Threshold": "✅",
                    "Weekend Sick Count": weekend_sick_count,
                    "Report Sick Threshold": "✅" if weekend_sick_count >= 3 else "❌"
                })

        # ── display ─────────────────────────────────────────────────────
        if flagged_persons:
            st.success(f"Found {len(flagged_persons)} personnel who hit MR threshold "
                    f"({start_date:%d %b %Y} → {end_date:%d %b %Y}).")
            st.table(flagged_persons)
        else:
            st.info("✅ No one hit the MR or Report Sick thresholds "
                    f"for the selected window.")


# ───────────────────────────────────────────────
# TAB 6 – DAILY ATTENDANCE   (patched)
# ───────────────────────────────────────────────
    with tab6:
        st.subheader("📊 Daily Attendance")

        # NEW population selector
        pop6 = st.radio(
            "Count population:",
            ("Only personnel with 4‑D", "All personnel"),
            index=0, horizontal=True, key="pop_tab6"
        )
        only_4d_att = (pop6 == "Only personnel with 4‑D")       # NEW

        att_date = st.date_input(
            "Date to view attendance for", datetime.today().date(),
            key="attendance_date"
        )

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        parade_records  = get_allparade_records(selected_company, SHEET_PARADE)

        # filter nominal according to selector
        records_nominal = [
            r for r in records_nominal
            if r["company"] == selected_company and
            (not only_4d_att or r["4d_number"])               # NEW
        ]
        total_nominal = len(records_nominal)

        def absent_set(the_date):
            ids = set()
            for row in parade_records:
                if row.get("company", "") != selected_company:
                    continue
                four_d = is_valid_4d(row.get("4d_number", ""))
                if only_4d_att and not four_d:                   # NEW
                    continue
                status_prefix = row.get("status", "").lower().split()[0]
                if status_prefix not in LEGEND_STATUS_PREFIXES:
                    continue
                sd = parse_ddmmyyyy(row.get("start_date_ddmmyyyy", ""))
                ed = parse_ddmmyyyy(row.get("end_date_ddmmyyyy",   ""))
                if sd.date() <= the_date <= ed.date():
                    uid = four_d or row.get("name", "").lower()
                    ids.add(uid)
            return ids

        # ------- A) single‑day ---------------------------------
        abs_today   = absent_set(att_date)
        present_cnt = total_nominal - len(abs_today)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Present", f"{present_cnt:02d}")
        with col2:
            st.metric("Absent",  f"{len(abs_today):02d}")
        with col3:
            st.metric("Nominal", f"{total_nominal:02d}")

        if abs_today:
            st.write("Absent personnel:")
            st.table(pd.DataFrame(sorted(abs_today), columns=["UID"]))

        # ------- B) overall‑to‑date ----------------------------
        st.subheader("Overall Attendance (1 Apr 2025 → today)")

        start_overall = datetime(2025, 4, 1).date()
        end_overall   = datetime.today().date()
        first_3wk_end = start_overall + timedelta(weeks=3) - timedelta(days=1)

        total_present_days = total_nominal_days = 0
        current = start_overall
        while current <= end_overall:
            if current > first_3wk_end and current.weekday() >= 5:
                current += timedelta(days=1)
                continue
            abs_ids = absent_set(current)
            total_present_days += (total_nominal - len(abs_ids))
            total_nominal_days += total_nominal
            current += timedelta(days=1)

        overall_pct = (
            total_present_days / total_nominal_days * 100
            if total_nominal_days else 0
        )
        st.metric("Attendance Rate", f"{overall_pct:.2f}%")
# ------------------------------------------------------------------------------
# 14) Feature F: Generate WhatsApp Message
# ------------------------------------------------------------------------------
elif feature == "Generate WhatsApp Message":
    st.header("Generate WhatsApp Message")

    # --- 1) Existing WhatsApp Message Generation for Selected Company ---
    records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
    records_parade2 = get_allparade_records(selected_company, SHEET_PARADE)

    selected_date = st.date_input("Select Parade Date", datetime.now(TIMEZONE).date())
    target_datetime = datetime.combine(selected_date, datetime.min.time())
    # Fetch nominal and parade records for the selected company
    company_nominal = [record for record in records_nominal if record['company'] == selected_company]
    company_parade = [record for record in records_parade2 if record['company'] == selected_company]

    if not company_nominal:
        st.warning(f"No nominal records found for company '{selected_company}'.")
        st.stop()

    # Generate the company-specific message
    company_message = generate_company_message(selected_company, company_nominal, company_parade, target_date=target_datetime)
    st.code(company_message, language='text')




elif feature == "Overall View":
    st.header("Overall View of All Conducts")
    # (a) Fetch all conducts
    conducts = get_conduct_records(selected_company, SHEET_CONDUCTS)
    if not conducts:
        st.info("No conducts available to display.")
    else:
        # (b) Convert to DataFrame
        df = pd.DataFrame(conducts)

        # (c) Convert the 'date' field to datetime objects for sorting
        # Assumes date format is "DDMMYYYY" as produced by ensure_date_str
        def parse_date(date_str):
            try:
                return datetime.strptime(date_str, "%d%m%Y")
            except ValueError:
                return None

        df['Date'] = df['date'].apply(parse_date)

        # (d) Warn if any dates could not be parsed
        invalid_dates = df['Date'].isnull()
        if invalid_dates.any():
            st.warning(f"{invalid_dates.sum()} conduct(s) have invalid date formats and will appear at the bottom.")
            logger.warning(f"{invalid_dates.sum()} conduct(s) have invalid date formats in company '{selected_company}'.")

        # (e) Sort the DataFrame by Date (latest first)
        df_sorted = df.sort_values(by='Date', ascending=False)

        # (f) Format the 'Date' column for display
        df_sorted['Date'] = df_sorted['Date'].dt.strftime("%d-%m-%Y")

        # (g) Select and rename columns for display (remove outlier and pointer columns)
        display_columns = {
            'Date': 'Date',
            'conduct_name': 'Conduct Name',
            'p/t plt1': 'P/T PLT1',
            'p/t plt2': 'P/T PLT2',
            'p/t plt3': 'P/T PLT3',
            'p/t plt4': 'P/T PLT4',
            'p/t coy hq': 'P/T Coy HQ',
            'p/t total': 'P/T Total',
            'submitted_by': 'Submitted By',
            'pointers': 'Safety PAR'
        }
        df_display = df_sorted.rename(columns=display_columns)[list(display_columns.values())]

        # -------------------------------------------------------------------------
        # **Added: Filtering and Sorting Options**
        # -------------------------------------------------------------------------
        st.subheader("Filter and Sort Conducts")
        with st.expander("🔍 Filter Conducts"):
            search_term = st.text_input(
                "Search by Conduct Name or Date (DDMMYYYY):",
                value="",
                help="Enter a keyword to filter conducts by name or date."
            )
            sort_field = st.selectbox("Sort By", options=["Date", "Conduct Name"], index=0)
            sort_order = st.radio("Sort Order", options=["Ascending", "Descending"], index=1)
            
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
                df_display['Date_Sort'] = pd.to_datetime(df_display['Date'], format="%d-%m-%Y", errors='coerce')
                df_display = df_display.sort_values(by='Date_Sort', ascending=ascending)
                df_display = df_display.drop(columns=['Date_Sort'])
            elif sort_field == "Conduct Name":
                df_display = df_display.sort_values(by='Conduct Name', ascending=ascending)

        # -------------------------------------------------------------------------
        st.subheader("All Conducts")
        st.dataframe(df_display, use_container_width=True)
        # -------------------------------------------------------------------------
        # **Added: Individuals' Missed Conducts**
        # -------------------------------------------------------------------------
        st.subheader("Individuals' Missed Conducts")
        # Build a dictionary to track which individuals (by their 4D number) missed which conducts
        missed_conducts_dict = defaultdict(set)
        # List of outlier columns in the normalized conduct records
        outlier_columns = [
            "plt1 outliers",
            "plt2 outliers",
            "plt3 outliers",
            "plt4 outliers",
            "coy hq outliers"
        ]
        for conduct in conducts:
            conduct_name = ensure_str(conduct.get('conduct_name', ''))
            conduct_outliers = set()
            for col in outlier_columns:
                outliers_str = ensure_str(conduct.get(col, ''))
                if outliers_str.lower() == 'none' or not outliers_str.strip():
                    continue  # Skip if no outliers listed
                # Split the outliers string by comma
                outliers = [o.strip() for o in outliers_str.split(',') if o.strip()]
                for outlier in outliers:
                    # Extract the 4D number using regex (e.g., "4D1204")
                    match = re.match(r'(4D\d{3,4})(?:\s*\(.*\))?', outlier, re.IGNORECASE)
                    if match:
                        four_d = match.group(1).upper()
                        conduct_outliers.add(four_d)
            for four_d in conduct_outliers:
                missed_conducts_dict[four_d].add(conduct_name)

        # Build a list of dictionaries for each individual
        missed_conducts_data = []
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        # Build a lookup: normalized 4D number (uppercase) to the person's name
        four_d_to_name = {row['4d_number'].upper(): row['name'] for row in records_nominal}
        for four_d, conducts_missed in missed_conducts_dict.items():
            name = four_d_to_name.get(four_d, "Unknown")
            missed_conducts_data.append({
                "4D_Number": four_d,
                "Name": name,
                "Missed Conducts Count": len(conducts_missed),
                "Missed Conducts": ", ".join(sorted(conducts_missed))
            })

        if missed_conducts_data:
            df_missed = pd.DataFrame(missed_conducts_data)
            # Sort individuals from most missed conducts to least
            df_missed = df_missed.sort_values(by="Missed Conducts Count", ascending=False).reset_index(drop=True)

            # Apply styling to bold the top 3 individuals
            def highlight_top3(row):
                return ['font-weight: bold' if row.name < 3 else '' for _ in row]
            styled_df = df_missed.style.apply(highlight_top3, axis=1)
            st.subheader("Missed Conducts by Individuals (Most to Least)")
            st.dataframe(styled_df, use_container_width=True)
        else:
            st.info("✅ **No individuals have missed any conducts.**")
            logger.info(f"No missed conducts recorded in company '{selected_company}' by user '{st.session_state.username}'.")

        # -------------------------------------------------------------------------
        logger.info(f"Displayed overall view of all conducts in company '{selected_company}' by user '{st.session_state.username}'.")

    st.header("Attendance Analytics")

    # Get Everything sheet data (assumes you have a gspread Worksheet object `sheet_everything`)
    SHEET_EVERYTHING = worksheets["everything"]
    everything_data = SHEET_EVERYTHING.get_all_values()
    if not everything_data:
        st.error("Everything sheet is empty!")
    else:
        # Assuming that the first three columns are static (e.g., Rank, 4D_Number, Name),
        # the remaining columns are conduct columns.
        conduct_headers = everything_data[0][3:]
        if not conduct_headers:
            st.error("No conduct columns found!")
        else:
            selected_conduct = st.selectbox("Select Conduct Column", conduct_headers)
            
            # Fetch the nominal records
            records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
            
            # Compute the analytics for the selected conduct column
            try:
                analytics = analyze_attendance(everything_data, records_nominal, selected_conduct)
            except ValueError as e:
                st.error(str(e))
                analytics = None
            
            if analytics:
                # Let the user choose which level to view
                view_options = [
                    "Overall Attendance",
                    "Platoon-Level Attendance",
                    "Section-Level Attendance",
                    "Individual-Level Attendance",
                    "Conduct Summary",
                    "Training-Wide Attendance"  # New option for combined conduct analytics
                ]
                selected_view = st.radio("Select View", view_options)
                
                if selected_view == "Overall Attendance":
                    overall = analytics['overall']
                    st.subheader("Overall Attendance")
                    st.write(
                        f"**Total Individuals:** {overall['total']}  |  "
                        f"**Present:** {overall['present']}  |  "
                        f"**Attendance:** {overall['percentage']:.2f}%"
                    )
                
                elif selected_view == "Platoon-Level Attendance":
                    platoon_summary = analytics['platoon_summary']
                    st.subheader("Platoon-Level Attendance")
                    # Create a DataFrame for easy display
                    df_platoon = pd.DataFrame.from_dict(platoon_summary, orient='index') \
                        .reset_index() \
                        .rename(columns={'index': 'Platoon'})
                    st.dataframe(df_platoon, use_container_width=True)
                    
                elif selected_view == "Section-Level Attendance":
                    section_summary = analytics['section_summary']
                    st.subheader("Section-Level Attendance")
                    # Build a DataFrame from the section summary
                    df_section = pd.DataFrame([
                        {
                            'Platoon': key[0],
                            'Section': key[1],
                            'Total': val['total'],
                            'Present': val['present']
                        }
                        for key, val in section_summary.items()
                    ])
                    st.dataframe(df_section, use_container_width=True)
                    
                elif selected_view == "Individual-Level Attendance":
                    individual_details = analytics['individual_details']
                    st.subheader("Individual-Level Attendance")
                    # Create a DataFrame for individual details
                    df_individual = pd.DataFrame([
                        {
                            'Name': name,
                            'Platoon': details['platoon'],
                            'Section': details['section'],
                            'Roll': details['roll'],
                            'Attendance': details['attendance']
                        }
                        for name, details in individual_details.items()
                    ])
                    st.dataframe(df_individual, use_container_width=True)
                    
                elif selected_view == "Conduct Summary":
                    conduct_summary = analytics['conduct_summary']
                    st.subheader("Conduct Summary")
                    # Build a DataFrame from the conduct summary
                    df_conduct = pd.DataFrame([
                        {
                            'Conduct Column': col,
                            'Total Nominal': summary['total'],
                            'Present': summary['present'],
                            'Attendance (%)': f"{summary['percentage']:.2f}"
                        }
                        for col, summary in conduct_summary.items()
                    ])
                    st.dataframe(df_conduct, use_container_width=True)
                    
                elif selected_view == "Training-Wide Attendance":
                    st.subheader("Training-Wide Attendance (Combined Conducts)")
                    # Determine the number of conduct columns (all columns after the first three)
                    headers = everything_data[0]
                    conduct_types = [classify_conduct(h) for h in headers[3:]]
                    
                    # Build a mapping from name to their attendance row (for fast lookup)
                    attendance_mapping = {}
                    for row in everything_data[1:]:
                        name = row[2].strip()
                        attendance_mapping[name] = row
                    
                    # Initialize aggregates
                    training_overall_total = 0
                    training_overall_present = 0
                    training_platoon_summary = {}   # {platoon: {'total': X, 'present': Y}}
                    training_section_summary = {}   # {(platoon, section): {'total': X, 'present': Y}}
                    training_individual_details = {}  # {name: {platoon, section, roll, yes_count, total, percentage}}
                    
                    for record in records_nominal:
                        name = record['name'].strip()
                        status = record.get('bmt_ptp', 'combined')  # NEW
                        row = attendance_mapping.get(name)
                        yes_count = 0
                        denom = 0
                        # Iterate over all conduct columns
                        for idx, ctype in enumerate(conduct_types, start=3):
                            applies = (
                                ctype == 'combined' or
                                status == 'combined' or
                                status == ctype
                            )
                            if not applies:
                                continue

                            denom += 1
                            value = row[idx].strip().lower() if row and len(row) > idx else ""
                            if value == "yes":
                                yes_count += 1

                        # nothing applied? skip the record entirely
                        if denom == 0:
                            continue
                        training_overall_total += denom
                        training_overall_present += yes_count
                        
                        # Parse platoon, section, and roll using the helper function
                        num_str = record.get('4d_number', '')
                        platoon, section, roll = parse_4d_number(num_str)
                        
                        # Aggregate only if both platoon and section are available
                        if platoon and section:
                            if platoon not in training_platoon_summary:
                                training_platoon_summary[platoon] = {'total': 0, 'present': 0}
                            training_platoon_summary[platoon]['total'] += denom
                            training_platoon_summary[platoon]['present'] += yes_count
                            
                            key = (platoon, section)
                            if key not in training_section_summary:
                                training_section_summary[key] = {'total': 0, 'present': 0}
                            training_section_summary[key]['total'] += denom
                            training_section_summary[key]['present'] += yes_count
                        
                        individual_percentage = (yes_count / denom * 100) if denom else 0
                        training_individual_details[name] = {
                            'platoon': platoon,
                            'section': section,
                            'roll': roll,
                            'yes_count': yes_count,
                            'total': denom,
                            'percentage': individual_percentage
                        }
                    
                    overall_percentage = (training_overall_present / training_overall_total * 100) if training_overall_total else 0
                    st.write(
                        f"**Attendance:** {overall_percentage:.2f}%"
                    )
                    
                    st.subheader("Platoon-Level Training-Wide Attendance")
                    df_platoon_tw = pd.DataFrame([
                        {
                            'Platoon': platoon,
                            'Attendance (%)': f"{(summary['present']/summary['total']*100):.2f}" if summary['total'] else "0.00"
                        }
                        for platoon, summary in training_platoon_summary.items()
                    ])
                    st.dataframe(df_platoon_tw, use_container_width=True)
                    
                    st.subheader("Section-Level Training-Wide Attendance")
                    df_section_tw = pd.DataFrame([
                        {
                            'Platoon': key[0],
                            'Section': key[1],
                            'Attendance (%)': f"{(val['present']/val['total']*100):.2f}" if val['total'] else "0.00"
                        }
                        for key, val in training_section_summary.items()
                    ])
                    st.dataframe(df_section_tw, use_container_width=True)
                    
                    st.subheader("Individual-Level Training-Wide Attendance")
                    df_individual_tw = pd.DataFrame([
                        {
                            'Name': name,
                            'Platoon': details['platoon'],
                            'Section': details['section'],
                            'Roll': details['roll'],
                            'Attendance (%)': f"{details['percentage']:.2f}"
                        }
                        for name, details in training_individual_details.items()
                    ])
                    def highlight_below_threshold(val):
                        try:
                            if float(val) < 75:
                                return 'background-color: #ff9999'
                        except Exception:
                            pass
                        return ''
                    
                    styled_df = df_individual_tw.style.applymap(highlight_below_threshold, subset=['Attendance (%)'])
                    st.dataframe(styled_df, use_container_width=True)
