import streamlit as st  # type: ignore
import gspread  # type: ignore
from oauth2client.service_account import ServiceAccountCredentials  # type: ignore
from datetime import datetime
from collections import defaultdict
import re
import pandas as pd  # type: ignore
import logging
import json
import os
from typing import List, Dict, Optional
from zoneinfo import ZoneInfo  # type: ignore
from datetime import timedelta

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
TIMEZONE = ZoneInfo('Asia/Singapore')  
USER_DB_PATH = "users.json"
NON_CMD_RANKS = ["PTE", "LCP", "CPL", "CFC", "REC"]
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


if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False
if 'username' not in st.session_state:
    st.session_state.username = ""
if 'user_companies' not in st.session_state:
    st.session_state.user_companies = []

def login():
    st.title("🔒 1SIRTracker")
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

if not st.session_state.authenticated:
    login()
    st.stop()

st.set_page_config(page_title="1SIRTracker", layout="centered")

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPES)

COMPANY_SPREADSHEETS = {
    "Alpha": "Alpha",
    "Bravo": "Bravo",
    "Charlie": "Charlie",
    "Support": "Support",
    "MSC": "MSC",
    "HQ": "HQ",
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
    Analyzes attendance for a specific conduct.
    """
    # ── setup ───────────────────────────────────────────────────────────────────
    nominal_mapping = {r['name'].strip(): r for r in nominal_data}
    headers = everything_data[0]
    if conduct_header not in headers:
        raise ValueError(f"Conduct column '{conduct_header}' not found.")
    conduct_idx = headers.index(conduct_header)

    attendance_mapping = {row[2].strip(): row for row in everything_data[1:]}

    overall_total = overall_present = 0
    platoon_summary, section_summary, individual_details = {}, {}, {}

    # ── main loop ───────────────────────────────────────────────────────────────
    for rec in nominal_data:
        name = rec['name'].strip()

        overall_total += 1
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
        total = present = 0

        for rec in nominal_data:
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
    sorted_platoons = sorted(all_platoons, key=lambda x: (x.lower() not in ('coy hq', 'hq'), x))
    for platoon in sorted_platoons:
        records = active_parade_by_platoon.get(platoon, [])

        # Determine platoon label
        if platoon.lower() in ('coy hq', 'hq'):
            platoon_label = "Coy HQ"
        elif selected_company == "Support":
            support_platoon_map = {
                "1": "SIGNAL PL",
                "2": "SCOUT PL",
                "3": "PIONEER PL",
                "4": "OPFOR PL"
            }
            platoon_label = support_platoon_map.get(platoon, f"Platoon {platoon}")
        else:
            platoon_label = f"Platoon {platoon}"

        # Total nominal strength for this platoon
        platoon_nominal = len([
            record for record in company_nominal_records
            if record.get('platoon', 'Coy HQ') == platoon
        ])

        # Initialize lists for conformant absentees split into commander and non-cmd,
        # plus non-conformant parade records (to be shown under "Pl Statuses")
        commander_absentees = []
        non_cmd_absentees = []
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
                # Split conformant absentees by whether their rank indicates a non-cmd
                if rank.upper() in NON_CMD_RANKS:
                    non_cmd_absentees.append({
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

        non_cmd_group = defaultdict(list)
        for absentee in non_cmd_absentees:
            key = (absentee['4d'].strip(), absentee['rank'].strip(), absentee['name'].strip())
            non_cmd_group[key].append(f"{absentee['status']} {absentee['details']}")

        if platoon.lower() not in ('coy hq', 'hq'):
            platoon_absent = len(commander_group) + len(non_cmd_group)
        else:
            # For Coy HQ, combine both groups
            combined_group = defaultdict(list)
            for absentee in (commander_absentees + non_cmd_absentees):
                key = (absentee['4d'].strip(), absentee['rank'].strip(), absentee['name'].strip())
                combined_group[key].append(f"{absentee['status']} {absentee['details']}")
            platoon_absent = len(combined_group)
        total_absent += platoon_absent

        # For platoons (other than Coy HQ), calculate nominal breakdown based on rank
        if platoon.lower() not in ('coy hq', 'hq'):
            platoon_nominal_records = [
                r for r in company_nominal_records
                if r.get('platoon', 'Coy HQ') == platoon
            ]
            commander_nominal = sum(
                1 for r in platoon_nominal_records if r.get('rank', '').upper() not in NON_CMD_RANKS
            )
            non_cmd_nominal = sum(
                1 for r in platoon_nominal_records if r.get('rank', '').upper() in NON_CMD_RANKS
            )
        else:
            commander_nominal = None
            non_cmd_nominal = None

        platoon_details.append({
            'label': platoon_label,
            'nominal': platoon_nominal,
            'unique_absent': platoon_absent,  # use the grouped count here
            'present': platoon_nominal - platoon_absent,
            'commander_group': commander_group,
            'non_cmd_group': non_cmd_group,
            'non_conformant': non_conformant_absentees,
            'commander_nominal': commander_nominal,
            'non_cmd_nominal': non_cmd_nominal
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

        # For platoons other than Coy HQ, show commander/non-cmd breakdown
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
                f"Non-Commander Absent Strength: {len(detail['non_cmd_group']):02d}/{detail['non_cmd_nominal']:02d}"
            )
            for (d, rank, name), details_list in detail['non_cmd_group'].items():
                details_str = ", ".join(details_list)
                if d:
                    message_lines.append(f"> {d} {rank} {name} ({details_str})")
                else:
                    message_lines.append(f"> {rank} {name} ({details_str})")
        else:
            # For Coy HQ, combine commander and non-cmd into one list
            combined_group = defaultdict(list)
            for key, details_list in detail['commander_group'].items():
                combined_group[key].extend(details_list)
            for key, details_list in detail['non_cmd_group'].items():
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
        normalized_row['dates taken'] = ensure_str(normalized_row.get('dates taken', ''))
        normalized_row['company'] = selected_company  # Add company information
        normalized_records.append(normalized_row)
    
    return normalized_records

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
        if person.get("Rank", "").upper() in NON_CMD_RANKS:
            person["Personnel_Type"] = "non-cmd"
        else:
            person["Personnel_Type"] = "cmd"
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
        if person.get("Rank", "").upper() in NON_CMD_RANKS:
            person["Personnel_Type"] = "non-cmd"
        else:
            person["Personnel_Type"] = "cmd"
    return data



st.title("1SIRTracker")

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

if "adhoc_personnel" not in st.session_state:
    st.session_state.adhoc_personnel = []
if "adhoc_conduct_name" not in st.session_state:
    st.session_state.adhoc_conduct_name = ""
if "adhoc_conduct_date" not in st.session_state:
    st.session_state.adhoc_conduct_date = ""

feature = st.sidebar.selectbox(
    "Select Feature",
    ["Add Conduct", "Add Ad-Hoc Conduct", "Update Conduct", "Update Parade", "Analytics", "Message"]
)

def add_pointer():
    st.session_state.conduct_pointers.append(
        {"observation": "", "reflection": "", "recommendation": ""}
    )
def add_update_pointer():
    st.session_state.update_conduct_pointers.append(
        {"observation": "", "reflection": "", "recommendation": ""}
    )

if feature == "Add Conduct":
    st.header("Add Conduct")
    st.info("""Please key in name of conduct in all caps and properly with reference to training programme and choose the session number accurately!!
            
Examples:
- ENDURANCE RUN
- GPMG LF""")

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
    st.session_state.conduct_name = st.text_input(
        "Conduct Name",
        value=st.session_state.conduct_name
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
        final_conduct_name = f"{st.session_state.conduct_name} {st.session_state.conduct_session}"
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
                                 key=lambda x: "ZZZ" if x.get("Rank", "").upper() in NON_CMD_RANKS else x.get("Rank", ""))
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

        # Initialize non-cmd and cmd counts for each platoon
        non_cmd_counts = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}
        cmd_counts = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}
        non_cmd_totals = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}
        cmd_totals = {"1": 0, "2": 0, "3": 0, "4": 0, "Coy HQ": 0}

        # Calculate total non-cmd and cmd for each platoon
        for person in records_nominal:
            plt = person.get("platoon", "")
            if plt in platoon_options:
                if person.get("rank", "").upper() in NON_CMD_RANKS:
                    non_cmd_totals[plt] += 1
                else:
                    cmd_totals[plt] += 1

        # Count participating non-cmd and cmd
        for row in edited_data:
            if not row.get('Is_Outlier', False):
                plt = platoon
                if plt in platoon_options:
                    if row.get('Rank', '').upper() in NON_CMD_RANKS:
                        non_cmd_counts[plt] += 1
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
            
            non_cmd_ratio = f"{non_cmd_counts[platoon]}/{non_cmd_totals[platoon]}"
            cmd_ratio = f"{cmd_counts[platoon]}/{cmd_totals[platoon]}"
            total_ratio = f"{non_cmd_counts[platoon] + cmd_counts[platoon]}/{total_strength_platoons[platoon]}"
            
            pt_plts[index] = f"non-cmd: {non_cmd_ratio}\ncmd: {cmd_ratio}\nTOTAL: {total_ratio}"

        # Calculate total participants and total strength
        total_non_cmd_part = sum(non_cmd_counts.values())
        total_non_cmd = sum(non_cmd_totals.values())
        total_cmd_part = sum(cmd_counts.values())
        total_cmd = sum(cmd_totals.values())
        total_part = total_non_cmd_part + total_cmd_part
        total_strength = sum(total_strength_platoons.values())

        # Format the totals
        pt_total = f"non-cmd: {total_non_cmd_part}/{total_non_cmd}\ncmd: {total_cmd_part}/{total_cmd}\nTOTAL: {total_part}/{total_strength}"

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
        st.session_state.conduct_name = ""
        st.session_state.conduct_table = []
        st.session_state.conduct_pointers = [
            {"observation": "", "reflection": "", "recommendation": ""}
        ]

elif feature == "Add Ad-Hoc Conduct":
    st.header("Add Ad-Hoc Conduct")
    st.info("Select a group of personnel to record a conduct. Their current on-status information for the selected date will be pre-loaded.")

    st.session_state.adhoc_conduct_name = st.text_input(
        "Ad-Hoc Conduct Name",
        value=st.session_state.adhoc_conduct_name
    )
    st.session_state.adhoc_conduct_date = st.text_input(
        "Date (DDMMYYYY)",
        value=st.session_state.adhoc_conduct_date
    )

    records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
    personnel_options = sorted([p['name'] for p in records_nominal if p.get('name')])
    
    selected_personnel_names = st.multiselect(
        "Select Personnel for this Conduct",
        options=personnel_options
    )

    if st.button("Load Personnel & Status"):
        date_str = st.session_state.adhoc_conduct_date.strip()
        if not selected_personnel_names:
            st.warning("Please select at least one person.")
            st.stop()
        if not date_str:
            st.error("Please enter a Date.")
            st.stop()
        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format (use DDMMYYYY).")
            st.stop()

        records_parade = get_allparade_records(selected_company, SHEET_PARADE)
        
        # Build conduct table for the selected personnel
        parade_map = defaultdict(list)
        for row in records_parade:
            person_name = row.get('name', '').strip().upper()
            parade_map[person_name].append(row)
        
        adhoc_data = []
        nominal_map = {p['name']: p for p in records_nominal}

        for name in selected_personnel_names:
            person = nominal_map.get(name)
            if not person: continue

            active_statuses = []
            for parade in parade_map.get(name.strip().upper(), []):
                try:
                    start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                    end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                    if start_dt <= date_obj.date() <= end_dt:
                        status = parade.get('status', '').strip().upper()
                        if status: active_statuses.append(status)
                except ValueError:
                    continue
            
            is_outlier = len(active_statuses) > 0
            status_desc = ", ".join(active_statuses) if is_outlier else ""
            adhoc_data.append({
                'Rank': person.get('rank', ''), 'Name': name, '4D_Number': person.get('4d_number', ''),
                'Is_Outlier': is_outlier, 'StatusDesc': status_desc
            })

        st.session_state.adhoc_personnel = adhoc_data
        logger.info(f"Loaded {len(adhoc_data)} personnel for ad-hoc conduct by user '{st.session_state.username}'.")

    if st.session_state.adhoc_personnel:
        st.write("Review outlier status. Uncheck 'Is_Outlier' if the person participated.")
        edited_data = st.data_editor(
            st.session_state.adhoc_personnel,
            use_container_width=True, num_rows="fixed", hide_index=True,
            column_config={
                "Name": st.column_config.TextColumn("Name", disabled=True),
                "4D_Number": st.column_config.TextColumn("4D_Number", disabled=True),
                "Rank": st.column_config.TextColumn("Rank", disabled=True),
            }
        )
    else:
        edited_data = None

    if st.button("Finalize Ad-Hoc Conduct"):
        conduct_name = st.session_state.adhoc_conduct_name.strip()
        conduct_date = st.session_state.adhoc_conduct_date.strip()

        if not conduct_name or not conduct_date or not edited_data:
            st.error("Please fill all fields and load personnel before finalizing.")
            st.stop()
        
        try:
            formatted_date = datetime.strptime(conduct_date, "%d%m%Y").strftime("%d%m%Y")
        except ValueError:
            st.error("Invalid date format. Please use DDMMYYYY.")
            st.stop()

        # Update 'Everything' sheet
        SHEET_EVERYTHING = worksheets["everything"]
        all_everything_data = SHEET_EVERYTHING.get_all_values()
        new_col_header = f"{formatted_date}, {conduct_name}"

        if new_col_header in all_everything_data[0]:
            st.error(f"A conduct with the name '{new_col_header}' already exists.")
            st.stop()

        new_col_index = len(all_everything_data[0]) + 1
        SHEET_EVERYTHING.update_cell(1, new_col_index, new_col_header)

        participation_map = {row["Name"]: "No" if row["Is_Outlier"] else "Yes" for row in edited_data}
        
        updates = []
        for row_idx, row in enumerate(all_everything_data[1:], start=2):
            name = row[2].strip()
            value = participation_map.get(name, "")
            cell = gspread.utils.rowcol_to_a1(row_idx, new_col_index)
            updates.append({'range': cell, 'values': [[value]]})

        if updates:
            SHEET_EVERYTHING.batch_update(updates)

        # Update 'Conducts' sheet
        non_cmd_part = sum(1 for p in edited_data if not p["Is_Outlier"] and p["Rank"].upper() in NON_CMD_RANKS)
        cmd_part = sum(1 for p in edited_data if not p["Is_Outlier"] and p["Rank"].upper() not in NON_CMD_RANKS)
        non_cmd_total = sum(1 for p in edited_data if p["Rank"].upper() in NON_CMD_RANKS)
        cmd_total = sum(1 for p in edited_data if p["Rank"].upper() not in NON_CMD_RANKS)
        pt_total_str = f"non-cmd: {non_cmd_part}/{non_cmd_total}\ncmd: {cmd_part}/{cmd_total}\nTOTAL: {non_cmd_part + cmd_part}/{len(edited_data)}"

        outliers_by_platoon = defaultdict(list)
        name_to_platoon_map = {p['name']: p['platoon'] for p in records_nominal}
        for person in edited_data:
            if person["Is_Outlier"]:
                platoon = name_to_platoon_map.get(person["Name"], "Coy HQ")
                status = f" ({person['StatusDesc']})" if person['StatusDesc'] else ""
                outliers_by_platoon[platoon].append(f"{person.get('4D_Number', '')} {person['Name']}{status}".strip())
        
        platoon_options = ["1", "2", "3", "4", "Coy HQ"]
        outliers_list = [", ".join(outliers_by_platoon.get(p, [])) or "None" for p in platoon_options]
        
        SHEET_CONDUCTS.append_row([
            formatted_date, conduct_name, "N/A", "N/A", "N/A", "N/A", "N/A", pt_total_str,
            outliers_list[0], outliers_list[1], outliers_list[2], outliers_list[3], outliers_list[4],
            "", st.session_state.username
        ])

        st.success(f"Ad-Hoc Conduct '{conduct_name}' on {formatted_date} has been finalized.")
        logger.info(f"Ad-Hoc Conduct '{conduct_name}' added by user '{st.session_state.username}'.")
        
        st.session_state.adhoc_personnel, st.session_state.adhoc_conduct_name, st.session_state.adhoc_conduct_date = [], "", ""


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

    # Disable platoon selection for ad-hoc conducts
    is_adhoc_conduct = conduct_record.get('p/t plt1', '').strip() == "N/A"
    if is_adhoc_conduct:
        st.info("Ad-hoc conduct detected. Platoon selection is not applicable.")
        st.session_state.conduct_platoon = "Ad-Hoc"
        platoon_display_options = ["Not Applicable"]
        platoon_disabled = True
    else:
        platoon_display_options = ["1", "2", "3", "4", "Coy HQ"]
        platoon_disabled = False

    st.subheader("Select Platoon to Update")
    st.session_state.conduct_platoon = st.selectbox( 
        "Select Platoon",
        options=platoon_display_options,
        index=platoon_display_options.index(str(st.session_state.conduct_platoon)) if not platoon_disabled and str(st.session_state.conduct_platoon) in platoon_display_options else 0,
        disabled=platoon_disabled,
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

        # Check again if it's an ad-hoc conduct to decide the loading logic
        is_adhoc_conduct_check = conduct_record.get('p/t plt1', '').strip() == "N/A"
        
        if is_adhoc_conduct_check:
            # Logic for loading Ad-Hoc conducts from the 'Everything' sheet
            st.info("Loading only the personnel involved in this ad-hoc conduct.")
            everything_data = worksheets["everything"].get_all_values()
            target_col_header = f"{conduct_record.get('date')}, {conduct_record.get('conduct_name')}"
            conduct_data = []

            if everything_data and len(everything_data) > 1:
                headers = everything_data[0]
                try:
                    col_idx = headers.index(target_col_header)
                    
                    # Consolidate all outlier strings to parse their status descriptions
                    outlier_keys = [f"plt{i} outliers" for i in range(1, 5)] + ["coy hq outliers"]
                    all_outliers_str = ", ".join(
                        [conduct_record.get(key, '') for key in outlier_keys if conduct_record.get(key, '').lower().strip() not in ('none', '')]
                    )
                    parsed_outliers = parse_existing_outliers(all_outliers_str)
                    nominal_map = {p['name'].lower(): p for p in records_nominal}

                    # Iterate through 'Everything' sheet to find participants
                    for row_data in everything_data[1:]:
                        if len(row_data) > col_idx:
                            name = row_data[2].strip()
                            attendance_status = row_data[col_idx].strip().lower()

                            if attendance_status in ("yes", "no"):
                                person_nominal = nominal_map.get(name.lower())
                                if person_nominal:
                                    is_outlier = attendance_status == "no"
                                    # Look up status description from parsed outliers
                                    status_desc = parsed_outliers.get(name.lower(), {}).get('status_desc', '')
                                    
                                    conduct_data.append({
                                        'Rank': person_nominal.get('rank', ''),
                                        'Name': name,
                                        '4D_Number': person_nominal.get('4d_number', ''),
                                        'Is_Outlier': is_outlier,
                                        'StatusDesc': status_desc
                                    })
                except ValueError:
                    st.error(f"Could not find conduct column '{target_col_header}' in Everything sheet.")
                    logger.error(f"Could not find conduct column '{target_col_header}'.")
            
        else:
            # Logic for loading regular, platoon-based conducts
            if platoon == "Ad-Hoc":
                st.error("A platoon must be selected to update this conduct.")
                st.stop()
                
            load_from_parade_state = False
            if platoon != "Coy HQ":
                pt_col_key = f'p/t plt{platoon}'
                outlier_col_key = f'plt{platoon} outliers'
            else:
                pt_col_key = 'p/t coy hq'
                outlier_col_key = 'coy hq outliers'

            pt_value = conduct_record.get(pt_col_key, '0/0')
            outliers_value = conduct_record.get(outlier_col_key, '')

            if "0/0" in pt_value and (not outliers_value or outliers_value.strip().lower() == "none"):
                load_from_parade_state = True

            if load_from_parade_state:
                conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)
                st.info("Loading personnel from parade state since no existing data found for this platoon.")
            else:
                conduct_data = build_fake_conduct_table(platoon, date_obj, records_nominal, records_parade)
                existing_outliers_str = conduct_record.get(outlier_col_key, "")
                existing_outliers = parse_existing_outliers(existing_outliers_str)
                
                # Merge existing outliers into the table
                for _, outlier_info in existing_outliers.items():
                    name_to_find = outlier_info["original"]
                    status_desc = outlier_info["status_desc"]
                    
                    for row in conduct_data:
                        # Match by name (case-insensitive) as the primary identifier
                        if row.get("Name", "").strip().lower() == name_to_find.strip().lower():
                            row["Is_Outlier"] = True
                            if status_desc:
                                row["StatusDesc"] = status_desc
                            break

        st.session_state.update_conduct_table = conduct_data
        st.success(f"Loaded {len(conduct_data)} personnel for the selected conduct.")
        logger.info(
            f"Loaded conduct personnel for '{selected_conduct}' "
            f"in company '{selected_company}' by user '{st.session_state.username}'."
        )

    if "update_conduct_table" in st.session_state and st.session_state.update_conduct_table:
        #st.subheader(f"Edit Conduct Data for Platoon {st.session_state.conduct_platoon}")
        #st.write("Toggle 'Is_Outlier' if not participating, or add new rows for extra people.")
        sorted_conduct_table = sorted(st.session_state.update_conduct_table, 
                                 key=lambda x: "ZZZ" if x.get("Rank", "").upper() in NON_CMD_RANKS else x.get("Rank", ""))
        st.write("In order to update, make sure correct platoon chosen and then press load on status for the table to reflect correct platoon. Hence, whenever changing platoon make sure to press load after that to reflect accordingly.")
        edited_data = st.data_editor(
            st.session_state.update_conduct_table,
            num_rows="fixed",
            hide_index=True,
        )
    else:
        edited_data = None

    if st.button("Update Conduct Data") and edited_data is not None:
        # --- COMMON SETUP ---
        # Get the conduct record to determine its type and find its row number
        try:
            conduct_parts = selected_conduct.split(" - ")
            conduct_date, conduct_name = conduct_parts[0].strip(), conduct_parts[1].strip()
            
            all_conduct_values = SHEET_CONDUCTS.get_all_values()
            row_number = -1
            for i, row in enumerate(all_conduct_values):
                if row[0] == conduct_date and row[1] == conduct_name:
                    row_number = i + 1
                    break
            if row_number == -1:
                st.error("Could not find the conduct to update. It may have been moved or deleted.")
                st.stop()
        except Exception as e:
            st.error(f"Error finding conduct row: {e}")
            st.stop()

        # Update the 'Everything' sheet (common to both ad-hoc and regular)
        SHEET_EVERYTHING = worksheets["everything"]
        formatted_date_str = ensure_date_str(conduct_record['date'])
        attendance_data = extract_attendance_data(edited_data)
        update_conduct_column_everything(
            SHEET_EVERYTHING, formatted_date_str, conduct_record['conduct_name'], attendance_data
        )

        # Update pointers (common to both)
        pointers_list = []
        for idx, pointer in enumerate(st.session_state.update_conduct_pointers, start=1):
            obs = pointer.get("observation", "").strip()
            refl = pointer.get("reflection", "").strip()
            rec = pointer.get("recommendation", "").strip()
            pointer_str = ""
            if obs: pointer_str += f"Observation {idx}:\n{obs}\n"
            if refl: pointer_str += f"Reflection {idx}:\n{refl}\n"
            if rec: pointer_str += f"Recommendation {idx}:\n{rec}\n"
            pointers_list.append(pointer_str.strip())
        new_pointers = "\n\n".join(pointers_list)
        SHEET_CONDUCTS.update_cell(row_number, 14, new_pointers)

        # --- LOGIC SPLIT: AD-HOC vs. REGULAR ---
        is_adhoc = conduct_record.get('p/t plt1', '').strip() == "N/A"

        if is_adhoc:
            # --- Ad-Hoc Conduct Update Logic ---
            st.info("Updating Ad-Hoc Conduct...")
            
            # 1. Calculate P/T Total for the ad-hoc group
            non_cmd_participating = sum(1 for p in edited_data if not p["Is_Outlier"] and p["Rank"].upper() in NON_CMD_RANKS)
            cmd_participating = sum(1 for p in edited_data if not p["Is_Outlier"] and p["Rank"].upper() not in NON_CMD_RANKS)
            non_cmd_total_group = sum(1 for p in edited_data if p["Rank"].upper() in NON_CMD_RANKS)
            cmd_total_group = sum(1 for p in edited_data if p["Rank"].upper() not in NON_CMD_RANKS)
            new_pt_total_value = f"non-cmd: {non_cmd_participating}/{non_cmd_total_group}\ncmd: {cmd_participating}/{cmd_total_group}\nTOTAL: {non_cmd_participating + cmd_participating}/{len(edited_data)}"
            SHEET_CONDUCTS.update_cell(row_number, 8, new_pt_total_value)

            # 2. Calculate and update outliers for all relevant platoons
            records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
            outliers_by_platoon = defaultdict(list)
            name_to_platoon_map = {p['name']: p['platoon'] for p in records_nominal}
            for person in edited_data:
                if person["Is_Outlier"]:
                    platoon_of_person = name_to_platoon_map.get(person["Name"], "Coy HQ")
                    status = f" ({person['StatusDesc']})" if person['StatusDesc'] else ""
                    outliers_by_platoon[platoon_of_person].append(f"{person.get('4D_Number', '')} {person['Name']}{status}".strip())
            
            platoon_options = ["1", "2", "3", "4", "Coy HQ"]
            for i, p_opt in enumerate(platoon_options):
                outlier_col_idx = 9 + i
                outliers_str = ", ".join(outliers_by_platoon.get(p_opt, [])) or "None"
                SHEET_CONDUCTS.update_cell(row_number, outlier_col_idx, outliers_str)
            
        else:
            # --- Regular Platoon Conduct Update Logic ---
            st.info("Updating Platoon Conduct...")
            platoon = str(st.session_state.conduct_platoon).strip()
            
            # 1. Calculate and update the specific platoon's P/T value
            records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
            non_cmd_counts = sum(1 for row in edited_data if not row.get('Is_Outlier', False) and row.get('Rank', '').upper() in NON_CMD_RANKS)
            cmd_counts = sum(1 for row in edited_data if not row.get('Is_Outlier', False) and row.get('Rank', '').upper() not in NON_CMD_RANKS)
            
            non_cmd_totals_platoon = sum(1 for p in records_nominal if p.get("platoon", "") == platoon and p.get("rank", "").upper() in NON_CMD_RANKS)
            cmd_totals_platoon = sum(1 for p in records_nominal if p.get("platoon", "") == platoon and p.get("rank", "").upper() not in NON_CMD_RANKS)
            new_participating = len([r for r in edited_data if not r['Is_Outlier']])
            new_total_platoon = len(edited_data)
            
            new_pt_value = f"non-cmd: {non_cmd_counts}/{non_cmd_totals_platoon}\ncmd: {cmd_counts}/{cmd_totals_platoon}\nTOTAL: {new_participating}/{new_total_platoon}"
            
            if platoon != "Coy HQ":
                pt_column_index = 2 + int(platoon)
                outlier_column_index = 8 + int(platoon)
            else:
                pt_column_index = 7
                outlier_column_index = 13
            SHEET_CONDUCTS.update_cell(row_number, pt_column_index, new_pt_value)

            # 2. Calculate and update the specific platoon's outliers
            outliers_for_platoon = [
                (f"{row.get('4D_Number', '')} {row['Name']}" + 
                 (f" ({row.get('StatusDesc', '')})" if row.get('StatusDesc') else "")).strip()
                for row in edited_data if row.get('Is_Outlier', False)
            ]
            SHEET_CONDUCTS.update_cell(row_number, outlier_column_index, ", ".join(outliers_for_platoon) or "None")

            # 3. Recalculate and update the overall P/T Total in column 8
            all_conduct_values_updated = SHEET_CONDUCTS.get_all_values()
            current_row_values = all_conduct_values_updated[row_number - 1]
            
            total_non_cmd_part, total_non_cmd, total_cmd_part, total_cmd = 0, 0, 0, 0
            # Columns 3 to 7 (P/T PLT1 to P/T Coy HQ)
            for pt_cell in current_row_values[2:7]:
                if pt_cell and pt_cell != "N/A":
                    lines = pt_cell.split('\n')
                    try:
                        non_cmd_line = lines[0]
                        if non_cmd_line.startswith("non-cmd:"):
                            non_cmd_parts = non_cmd_line.replace("non-cmd:", "").strip().split('/')
                            total_non_cmd_part += int(non_cmd_parts[0])
                            total_non_cmd += int(non_cmd_parts[1])
                        
                        cmd_line = lines[1]
                        if cmd_line.startswith("cmd:"):
                            cmd_parts = cmd_line.replace("cmd:", "").strip().split('/')
                            total_cmd_part += int(cmd_parts[0])
                            total_cmd += int(cmd_parts[1])
                    except (IndexError, ValueError):
                        continue # Ignore malformed cells
            
            total_part = total_non_cmd_part + total_cmd_part
            total_strength = total_non_cmd + total_cmd
            pt_total = f"non-cmd: {total_non_cmd_part}/{total_non_cmd}\ncmd: {total_cmd_part}/{total_cmd}\nTOTAL: {total_part}/{total_strength}"
            SHEET_CONDUCTS.update_cell(row_number, 8, pt_total)

        st.success(f"Conduct '{selected_conduct}' updated successfully.")
        logger.info(
            f"Conduct '{selected_conduct}' updated successfully in company '{selected_company}' "
            f"by user '{st.session_state.username}'."
        )


        # Optionally, clear the conduct table if desired
elif feature == "Update Parade":
    st.header("Update Parade State")

    st.markdown(
        """
        **Please use one of the following standard prefixes for the status.** 
        **Please make sure to update the RSI/RSO status in the Parade State sheet.** 
        
        You can add any additional details in parentheses `()` after the prefix. For `RSI` and `RSO`, please write it like `MC (RSI)`.

        **Examples:**
        - `MC (RSO)`
        - `ML (RSI)`
        
        ---
        
        **Standard Prefixes:**
        - `OL` - Overseas Leave
        - `LL` - Local Leave
        - `ML` - Medical Leave
        - `MC` - Medical Certificate
        - `AO` - Attached Out
        - `OIL` - Off in Lieu
        - `MA` - Medical Appointment
        - `SO` - Stay Out
        - `CL` - Compassionate Leave
        - `I/A` - Interview / Appointment
        - `AWOL` - AWOL
        - `HL` - Hospitalisation Leave
        - `Others` - Others
        """
    )

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
                                 key=lambda x: "ZZZ" if x.get("Rank", "").upper() in NON_CMD_RANKS else x.get("Rank", ""))
        edited_data = st.data_editor(
            st.session_state.parade_table,
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Name": st.column_config.TextColumn("Name", disabled=True),
                "4D_Number": st.column_config.TextColumn("4D_Number", disabled=True),
                "Rank": st.column_config.TextColumn("Rank", disabled=True),
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
elif feature == "Analytics":
    
    query_mode = st.radio(
        "Select Analytics Mode",
        ("By Personnel", "By Conduct"),
        horizontal=True,
        key="analytics_mode"
    )

    if query_mode == "By Personnel":
        st.subheader("Query by Personnel")
        # 1. Get all personnel from nominal roll for the multiselect
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        personnel_names = sorted([p['name'] for p in records_nominal if p['name']])
        commanders = sorted([p['name'] for p in records_nominal if p['name'] and p['rank'].upper() not in NON_CMD_RANKS])
        non_commanders = sorted([p['name'] for p in records_nominal if p['name'] and p['rank'].upper() in NON_CMD_RANKS])

        # 2. Selection UI
        all_personnel_option = "ALL PERSONNEL"
        commanders_option = "ALL COMMANDERS"
        non_commanders_option = "ALL NON-COMMANDERS"
        special_options = [all_personnel_option, commanders_option, non_commanders_option]

        selected_options = st.multiselect(
            "Select groups or individuals to query.",
            options=special_options + personnel_names,
            default=[]
        )

        # Determine the list of people to query
        names_to_query_set = set()
        if all_personnel_option in selected_options:
            names_to_query_set.update(personnel_names)
        if commanders_option in selected_options:
            names_to_query_set.update(commanders)
        if non_commanders_option in selected_options:
            names_to_query_set.update(non_commanders)
        
        # Add any individually selected people
        for option in selected_options:
            if option not in special_options:
                names_to_query_set.add(option)
        
        names_to_query = sorted(list(names_to_query_set))

        if not names_to_query:
            st.info("Please select personnel from the list above to see their information.")
            st.stop()
            
        # Data fetching for all tabs
        records_parade = get_allparade_records(selected_company, SHEET_PARADE)
        sheet_everything = worksheets.get("everything")
        everything_data = sheet_everything.get_all_values() if sheet_everything else []

        # Create a mapping from name to nominal record for easy lookup
        nominal_map = {p['name'].lower(): p for p in records_nominal}
        
        # Create tabs
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Medical Statuses", "Leaves", "RSI/RSO", "Training Attendance", "Conduct Records", "Daily Attendance"])

        # Helper function to parse dates
        def parse_ddmmyyyy(d):
            try:
                return datetime.strptime(str(d), "%d%m%Y")
            except (ValueError, TypeError):
                return None

        # TAB 1: MEDICAL STATUSES
        with tab1:
            st.subheader("Medical Statuses")
            display_prefixes = ("ex", "rib", "ld", "mc", "ml")

            all_medical_summary = []
            group_totals = defaultdict(int)

            for name in names_to_query:
                person_parade_records = [
                    r for r in records_parade 
                    if r.get('name', '').strip().lower() == name.strip().lower()
                ]
                
                person_totals = defaultdict(int)
                medical_details = []

                for record in person_parade_records:
                    status = record.get("status", "").lower()
                    for prefix in display_prefixes:
                        if status.startswith(prefix):
                            start_date = parse_ddmmyyyy(record.get("start_date_ddmmyyyy", ""))
                            end_date = parse_ddmmyyyy(record.get("end_date_ddmmyyyy", ""))
                            
                            duration = "Unknown"
                            if start_date and end_date and end_date >= start_date:
                                days = (end_date - start_date).days + 1
                                duration = f"{days} day(s)"
                                person_totals[prefix] += days

                            medical_details.append({
                                "Status": record.get("status", ""),
                                "Start Date": record.get("start_date_ddmmyyyy", ""),
                                "End Date": record.get("end_date_ddmmyyyy", ""),
                                "Duration": duration
                            })
                            break # Move to next record
                
                for prefix, total in person_totals.items():
                    group_totals[prefix] += total

                nominal_info = nominal_map.get(name.lower(), {})
                all_medical_summary.append({
                    "Rank": nominal_info.get('rank', 'N/A'),
                    "Name": name,
                    "EX Days": person_totals['ex'],
                    "RIB Days": person_totals['rib'],
                    "LD Days": person_totals['ld'],
                    "MC Days": person_totals['mc'],
                    "ML Days": person_totals['ml'],
                })

                if medical_details:
                    with st.expander(f"View medical history for {name}"):
                        st.table(medical_details)
            
            if any(opt in selected_options for opt in special_options) and names_to_query:
                st.subheader("Group Summary (Medical)")
                num_people = len(names_to_query)
                st.metric("Selected Personnel", f"{num_people}")
                st.markdown("---")
                
                prefix_map = {
                    "ex": "Excuse", "rib": "RIB", "ld": "Light Duty", "mc": "MC", "ml": "Med Leave"
                }
                cols = st.columns(len(display_prefixes))

                for i, prefix in enumerate(display_prefixes):
                    total_days = group_totals.get(prefix, 0)
                    avg_days = total_days / num_people if num_people > 0 else 0
                    label = prefix_map.get(prefix, prefix.upper())
                    with cols[i]:
                        st.metric(f"Total {label} Days", total_days)
                        st.metric(f"Avg {label} Days", f"{avg_days:.2f}")
                st.markdown("---")


            if all_medical_summary:
                df_medical_summary = pd.DataFrame(all_medical_summary)
                st.dataframe(df_medical_summary, use_container_width=True, hide_index=True)
            else:
                st.info("No medical status records found for the selected personnel.")

        # TAB 2: LEAVE COUNTER
        with tab2:
            st.subheader("Leaves")
            leave_prefixes = ("ll", "ol", "leave")
            
            all_leave_records = []
            group_total_leaves = 0

            for name in names_to_query:
                person_parade_records = [
                    r for r in records_parade 
                    if r.get('name', '').strip().lower() == name.strip().lower()
                ]
                
                total_leave_days = 0
                leave_details = []
                    
                for record in person_parade_records:
                    status = record.get("status", "").lower()
                    if any(status.startswith(p) for p in leave_prefixes):
                        start_date = parse_ddmmyyyy(record.get("start_date_ddmmyyyy", ""))
                        end_date = parse_ddmmyyyy(record.get("end_date_ddmmyyyy", ""))
                        
                        duration = "Unknown"
                        if start_date and end_date and end_date >= start_date:
                            days = (end_date - start_date).days + 1
                            total_leave_days += days
                            duration = f"{days} day(s)"
                        
                        leave_details.append({
                            "Status": record.get("status", ""),
                            "Start Date": record.get("start_date_ddmmyyyy", ""),
                            "End Date": record.get("end_date_ddmmyyyy", ""),
                            "Duration": duration
                        })
                
                group_total_leaves += total_leave_days
                nominal_info = nominal_map.get(name.lower(), {})
                remaining_leaves = 14 - total_leave_days
                
                all_leave_records.append({
                    "Rank": nominal_info.get('rank', 'N/A'),
                    "Name": name,
                    "Leaves Taken (days)": total_leave_days,
                    "Leaves Remaining (days)": max(0, remaining_leaves)
                })

                if leave_details:
                    with st.expander(f"View leave history for {name}"):
                        st.table(leave_details)
            
            if any(opt in selected_options for opt in special_options) and names_to_query:
                st.subheader("Group Summary (Leave)")
                num_people = len(names_to_query)
                avg_leaves = group_total_leaves / num_people if num_people > 0 else 0
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Selected Personnel", num_people)
                with col2:
                    st.metric("Total Leave Days Taken", group_total_leaves)
                with col3:
                    st.metric("Avg Leave / Person", f"{avg_leaves:.2f}")

            if all_leave_records:
                df_leave_summary = pd.DataFrame(all_leave_records)
                st.dataframe(df_leave_summary, use_container_width=True, hide_index=True)
            else:
                st.info("No leave records found for the selected personnel.")

        # TAB 3: RSI/RSO
        with tab3:
            st.subheader("RSI/RSO Records")
            rsi_rso_prefixes = ("rsi", "rso")
            
            all_rsi_rso_summary = []
            group_total_rsi = 0
            group_total_rso = 0

            for name in names_to_query:
                person_parade_records = [
                    r for r in records_parade 
                    if r.get('name', '').strip().lower() == name.strip().lower()
                ]
                
                total_rsi = 0
                total_rso = 0
                rsi_rso_details = []
                    
                for record in person_parade_records:
                    status = record.get("status", "").lower()
                    
                    is_rsi_or_rso = False
                    if "rsi" in status:
                        total_rsi += 1
                        is_rsi_or_rso = True
                    elif "rso" in status:
                        total_rso += 1
                        is_rsi_or_rso = True

                    if is_rsi_or_rso:
                        start_date = parse_ddmmyyyy(record.get("start_date_ddmmyyyy", ""))
                        end_date = parse_ddmmyyyy(record.get("end_date_ddmmyyyy", ""))
                        
                        duration = "Unknown"
                        if start_date and end_date and end_date >= start_date:
                            days = (end_date - start_date).days + 1
                            duration = f"{days} day(s)"

                        rsi_rso_details.append({
                            "Status": record.get("status", ""),
                            "Start Date": record.get("start_date_ddmmyyyy", ""),
                            "End Date": record.get("end_date_ddmmyyyy", ""),
                            "Duration": duration
                        })
                
                group_total_rsi += total_rsi
                group_total_rso += total_rso
                nominal_info = nominal_map.get(name.lower(), {})
                all_rsi_rso_summary.append({
                    "Rank": nominal_info.get('rank', 'N/A'),
                    "Name": name,
                    "RSI Count": total_rsi,
                    "RSO Count": total_rso
                })

                if rsi_rso_details:
                    with st.expander(f"View RSI/RSO history for {name}"):
                        st.table(rsi_rso_details)
            
            if any(opt in selected_options for opt in special_options) and names_to_query:
                st.subheader("Group Summary (RSI/RSO)")
                num_people = len(names_to_query)
                avg_rsi = group_total_rsi / num_people if num_people > 0 else 0
                avg_rso = group_total_rso / num_people if num_people > 0 else 0

                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    st.metric("Selected Personnel", num_people)
                with col2:
                    st.metric("Total RSIs", group_total_rsi)
                with col3:
                    st.metric("Avg RSI / Person", f"{avg_rsi:.2f}")
                with col4:
                    st.metric("Total RSOs", group_total_rso)
                with col5:
                    st.metric("Avg RSO / Person", f"{avg_rso:.2f}")

            if all_rsi_rso_summary:
                df_summary = pd.DataFrame(all_rsi_rso_summary)
                st.dataframe(df_summary, use_container_width=True, hide_index=True)
            else:
                st.info("No RSI/RSO records found for the selected personnel.")

        # TAB 4: ATTENDANCE HISTORY
        with tab4:
            st.subheader("Training Attendance")
            
            if not everything_data or len(everything_data) < 2:
                st.warning("The 'Everything' sheet is empty or has no data, so attendance history cannot be displayed.")
            else:
                headers = everything_data[0]
                conduct_headers = headers[3:]
                
                attendance_map = {row[2].strip().lower(): row for row in everything_data[1:]}

                all_attendance_records = []
                
                for name in names_to_query:
                    person_row = attendance_map.get(name.lower())
                    
                    nominal_info = nominal_map.get(name.lower(), {})
                    rank = nominal_info.get('rank', 'N/A')

                    if person_row:
                        attended_count = 0
                        total_conducts = 0
                        missed_conducts_list = []
                        
                        for i, conduct_name in enumerate(conduct_headers):
                            col_idx = i + 3
                            attendance_status = person_row[col_idx].strip().lower() if len(person_row) > col_idx else ""
                            
                            if attendance_status in ("yes", "no"):
                                total_conducts += 1
                                if attendance_status == 'yes':
                                    attended_count += 1
                                else:
                                    missed_conducts_list.append(conduct_name)
                        
                        attendance_percentage = (attended_count / total_conducts * 100) if total_conducts > 0 else 0
                        
                        all_attendance_records.append({
                            "Rank": rank,
                            "Name": name,
                            "Attendance": f"{attended_count}/{total_conducts}",
                            "Percentage": f"{attendance_percentage:.2f}%"
                        })

                        if missed_conducts_list:
                            with st.expander(f"View missed conducts for {name}"):
                                st.write(", ".join(missed_conducts_list))
                    else:
                        all_attendance_records.append({
                            "Rank": rank,
                            "Name": name,
                            "Attendance": "N/A",
                            "Percentage": "N/A"
                        })


                if any(opt in selected_options for opt in special_options) and names_to_query:
                    st.subheader("Group Summary (Training Attendance)")
                    group_attended_count = sum(int(record["Attendance"].split("/")[0]) for record in all_attendance_records)
                    group_total_conducts = sum(int(record["Attendance"].split("/")[1]) for record in all_attendance_records)
                    group_attendance_percentage = (group_attended_count / group_total_conducts * 100) if group_total_conducts > 0 else 0
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Selected Personnel", len(names_to_query))
                    with col2:
                        st.metric("Overall Group Attendance", f"{group_attendance_percentage:.2f}%")

                if all_attendance_records:
                    df_attendance = pd.DataFrame(all_attendance_records)
                    st.dataframe(df_attendance, use_container_width=True, hide_index=True)
                else:
                    st.info("No attendance records found for the selected personnel in the 'Everything' sheet.")
        
        # TAB 5: CONDUCT RECORDS
        with tab5:
            st.subheader("Individual Conduct Records")

            if not everything_data or len(everything_data) < 2:
                st.warning("The 'Everything' sheet is empty, so conduct records cannot be displayed.")
            else:
                headers = everything_data[0]
                conduct_headers = headers[3:]
                attendance_map = {row[2].strip().lower(): row for row in everything_data[1:]}

                conduct_filter = st.text_input("Filter conducts by name:", key="conduct_record_filter").lower()

                # Pre-process headers to group conduct series
                all_conduct_series = defaultdict(dict)
                one_off_conducts = []
                for header in conduct_headers:
                    try:
                        conduct_name_part = header.split(', ')[1]
                    except IndexError:
                        conduct_name_part = header
                    match = re.match(r'^(.*\S)\s+(\d+)$', conduct_name_part)
                    if match:
                        base_name, session = match.groups()
                        all_conduct_series[base_name.strip()][int(session)] = header
                    else:
                        one_off_conducts.append(header)

                for name in names_to_query:
                    person_row = attendance_map.get(name.lower())
                    if not person_row:
                        continue

                    credited_conducts = []
                    
                    # Process one-off conducts
                    for header in one_off_conducts:
                        col_idx = headers.index(header)
                        status = person_row[col_idx].strip().lower() if len(person_row) > col_idx else ""
                        if status == 'yes':
                            credited_conducts.append(header)

                    # Process conduct series
                    for base_name, sessions in all_conduct_series.items():
                        yes_count = 0
                        for session_num, header in sessions.items():
                            col_idx = headers.index(header)
                            status = person_row[col_idx].strip().lower() if len(person_row) > col_idx else ""
                            if status == 'yes':
                                yes_count += 1
                        
                        # Credit for the first `yes_count` sessions
                        sorted_sessions = sorted(sessions.keys())
                        for i in range(min(yes_count, len(sorted_sessions))):
                            session_to_credit = sorted_sessions[i]
                            credited_conducts.append(sessions[session_to_credit])

                    # Apply the filter
                    filtered_conducts = [c for c in credited_conducts if conduct_filter in c.lower()]

                    nominal_info = nominal_map.get(name.lower(), {})
                    rank = nominal_info.get('rank', 'N/A')
                    
                    with st.expander(f"View conduct records for {rank} {name}"):
                        if filtered_conducts:
                            st.write(", ".join(sorted(list(set(filtered_conducts)))))
                        else:
                            st.write("No matching conducts found.")

        # TAB 6: DAILY ATTENDANCE
        with tab6:
            st.subheader("Daily Attendance")

            # Date range is fixed from June 14 of the current year to today.
            today = datetime.today().date()
            start_date = datetime(today.year, 6, 14).date()
            end_date = today
            
            st.write(f"Displaying attendance percentage from {start_date.strftime('%d %b %Y')} to {end_date.strftime('%d %b %Y')}.")

            if start_date > end_date:
                st.warning(f"The fixed start date ({start_date.strftime('%d %b %Y')}) is in the future. No data to display.")
                st.stop()

            all_attendance_summary = []
            group_attendance_percentages = []
            total_days_in_range = (end_date - start_date).days + 1

            for name in names_to_query:
                # Get all parade records for the person
                person_parade_records = [
                    r for r in records_parade 
                    if r.get('name', '').strip().lower() == name.strip().lower()
                ]

                absent_dates = set()
                for record in person_parade_records:
                    status_prefix = record.get("status", "").lower().split(' ')[0]
                    if status_prefix in LEGEND_STATUS_PREFIXES:
                        record_start = parse_ddmmyyyy(record.get("start_date_ddmmyyyy", ""))
                        record_end = parse_ddmmyyyy(record.get("end_date_ddmmyyyy", ""))

                        if record_start and record_end:
                            # Find the intersection of the record's date range and the overall query range
                            overlap_start = max(start_date, record_start.date())
                            overlap_end = min(end_date, record_end.date())

                            # If they overlap, add all dates in the overlap period to the set
                            if overlap_start <= overlap_end:
                                current_date = overlap_start
                                while current_date <= overlap_end:
                                    absent_dates.add(current_date)
                                    current_date += timedelta(days=1)
                
                num_absent_days = len(absent_dates)
                present_days = total_days_in_range - num_absent_days
                attendance_percentage = (present_days / total_days_in_range * 100) if total_days_in_range > 0 else 0
                group_attendance_percentages.append(attendance_percentage)
                
                nominal_info = nominal_map.get(name.lower(), {})
                all_attendance_summary.append({
                    "Rank": nominal_info.get('rank', 'N/A'),
                    "Name": name,
                    "Attendance (%)": f"{attendance_percentage:.2f}%"
                })

            if any(opt in selected_options for opt in special_options) and names_to_query:
                st.subheader("Group Summary (Daily Attendance)")
                avg_group_percentage = sum(group_attendance_percentages) / len(group_attendance_percentages) if group_attendance_percentages else 0
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Selected Personnel", len(names_to_query))
                with col2:
                    st.metric("Average Daily Attendance", f"{avg_group_percentage:.2f}%")
                
            if all_attendance_summary:
                df_summary = pd.DataFrame(all_attendance_summary)
                st.dataframe(df_summary, use_container_width=True, hide_index=True)

    elif query_mode == "By Conduct":
        st.subheader("Query by Conduct")
        st.info("Select a conduct to see the status of all personnel, including back-crediting to allow for progressive tracking.")

        sheet_everything = worksheets.get("everything")
        everything_data = sheet_everything.get_all_values() if sheet_everything else []

        if not everything_data or len(everything_data) < 2:
            st.warning("The 'Everything' sheet is empty, so conducts cannot be queried.")
            st.stop()
        
        headers = everything_data[0]
        conduct_headers = headers[3:]
        
        selected_conducts = st.multiselect(
            "Select one or more conducts to view:",
            options=conduct_headers
        )

        if not selected_conducts:
            st.info("Please select a conduct from the list above.")
            st.stop()

        # Data fetching needed for this mode
        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        nominal_map = {p['name'].lower(): p for p in records_nominal}
        attendance_map = {row[2].strip().lower(): row for row in everything_data[1:]}

        # Pre-process headers to group conduct series
        all_conduct_series = defaultdict(dict)
        for header in conduct_headers:
            try:
                conduct_name_part = header.split(', ')[1]
            except IndexError:
                conduct_name_part = header
            match = re.match(r'^(.*\S)\s+(\d+)$', conduct_name_part)
            if match:
                base_name, session = match.groups()
                all_conduct_series[base_name.strip()][int(session)] = header
        
        for conduct_header in selected_conducts:
            st.markdown(f"#### Results for: `{conduct_header}`")

            # Determine if the selected conduct is part of a series
            base_name_selected, session_selected = None, None
            is_series = False
            try:
                conduct_name_part = conduct_header.split(', ')[1]
            except IndexError:
                conduct_name_part = conduct_header
            match = re.match(r'^(.*\S)\s+(\d+)$', conduct_name_part)
            if match:
                base_name_selected, session_selected = match.groups()
                base_name_selected = base_name_selected.strip()
                session_selected = int(session_selected)
                is_series = True

            results = []
            for person in records_nominal:
                name_lower = person['name'].lower()
                person_row = attendance_map.get(name_lower)
                
                status = "Not Marked"
                
                if person_row:
                    original_status = "Not Marked"
                    try:
                        col_idx = headers.index(conduct_header)
                        original_status = person_row[col_idx].strip().lower()
                    except (ValueError, IndexError):
                        pass # Keep as Not Marked

                    if not is_series:
                        status = original_status
                    else:
                        # Back-crediting logic
                        series_sessions = all_conduct_series.get(base_name_selected, {})
                        
                        # Find which sessions were actually attended
                        actually_attended_sessions = {}
                        for s_num, s_header in series_sessions.items():
                            try:
                                s_col_idx = headers.index(s_header)
                                if len(person_row) > s_col_idx and person_row[s_col_idx].strip().lower() == 'yes':
                                    actually_attended_sessions[s_num] = s_header
                            except (ValueError, IndexError):
                                continue
                        
                        yes_count = len(actually_attended_sessions)
                        source_of_credit_header = ""
                        if actually_attended_sessions:
                            latest_attended_num = max(actually_attended_sessions.keys())
                            source_of_credit_header = actually_attended_sessions[latest_attended_num]

                        # Determine which sessions the person is credited for
                        sorted_sessions_nums = sorted(series_sessions.keys())
                        credited_sessions = sorted_sessions_nums[:yes_count]

                        if session_selected in credited_sessions:
                            if session_selected in actually_attended_sessions:
                                status = "Yes"
                            else:
                                status = f"Yes (from {source_of_credit_header})"
                        else:
                            status = "No" if original_status in ('yes', 'no') else "Not Marked"

                results.append({
                    "Rank": person.get('rank', 'N/A'),
                    "Name": person.get('name', 'N/A'),
                    "Status": status.capitalize()
                })
            
            if results:
                df_results = pd.DataFrame(results)
                st.dataframe(df_results, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# 14) Feature F: Generate Message
# ------------------------------------------------------------------------------
elif feature == "Message":
    st.header("Parade State Message")

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

