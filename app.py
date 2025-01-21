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
from typing import List, Dict
from zoneinfo import ZoneInfo  # type: ignore
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
            "conducts": sh.worksheet("Conducts"),
            "safety": sh.worksheet("Safety")
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

def safety_sharing_app_form(SHEET_SAFETY):
    """
    Safety Sharing UI with st.data_editor inside a form,
    so the app won't re-run on every checkbox tick.
    The sheet is updated only when the form is submitted.
    """


    # 1) Read Safety sheet header
    header_row = SHEET_SAFETY.row_values(1)
    if not header_row:
        st.error("Safety sheet is empty. Ensure row 1 has at least 'Rank' and 'Name'.")
        return

    if len(header_row) < 2:
        st.error("Safety sheet header must have at least 2 columns: 'Rank', 'Name'.")
        return

    # Collect existing columns (Weeks) after "Rank" & "Name"
    existing_cols = header_row[2:]

    # Keep track of which column is selected in session_state
    if "safety_selected_col" not in st.session_state:
        st.session_state.safety_selected_col = None

    # ──────────────────────────────────────────────────────────────────────────
    # SECTION A: Select or Create a Week Column
    # ──────────────────────────────────────────────────────────────────────────
    st.subheader("1) Choose or Create a Week Column")

    if st.session_state.safety_selected_col:
        # Already have a selected column
        st.info(f"Currently selected column: **{st.session_state.safety_selected_col}**")
        if st.button("Change/Reset Column"):
            st.session_state.safety_selected_col = None
            st.rerun()
    else:
        # No column chosen yet → radio to select or create
        week_mode = st.radio("Week Mode:", ["Select Existing", "Create New"], horizontal=True)

        if week_mode == "Select Existing":
            if not existing_cols:
                st.warning("No existing columns found. Create one first.")
            else:
                chosen_col = st.selectbox("Choose a Column:", options=existing_cols)
                if st.button("Use This Column"):
                    st.session_state.safety_selected_col = chosen_col
                    st.rerun()

        else:
            # Creating a new column
            user_week = st.text_input("Week # (e.g. 'Week 1')")
            user_pointers = st.text_input("Pointers (short description)")

            if st.button("Create New Column"):
                if not user_week.strip():
                    st.error("Please enter something like 'Week 1'.")
                    return

                new_col_name = user_week.strip()
                if user_pointers.strip():
                    new_col_name += f" ({user_pointers.strip()})"

                try:
                    current_header = SHEET_SAFETY.row_values(1)
                    new_col_index = len(current_header) + 1  # next free column
                    SHEET_SAFETY.update_cell(1, new_col_index, new_col_name)
                    st.success(f"Created new column: '{new_col_name}'")

                    # store in session & re-run
                    st.session_state.safety_selected_col = new_col_name
                    st.rerun()

                except Exception as e:
                    st.error(f"Error creating column '{new_col_name}': {e}")
                    return

    # If still no column is chosen, stop
    if not st.session_state.safety_selected_col:
        return

    # ──────────────────────────────────────────────────────────────────────────
    # SECTION B: Data Editor & Date Input in a Form
    # ──────────────────────────────────────────────────────────────────────────
    selected_col_name = st.session_state.safety_selected_col
    updated_header = SHEET_SAFETY.row_values(1)
    if selected_col_name not in updated_header:
        st.error(f"Column '{selected_col_name}' not found in the updated header.")
        return

    col_index = updated_header.index(selected_col_name) + 1  # 1-based index

    st.subheader(f"2) Enter Date & Tick Attendees in '{selected_col_name}'")

    # Grab all rows
    all_values = SHEET_SAFETY.get_all_values()
    if len(all_values) < 2:
        st.warning("No data rows below the header. Ensure your sheet has personnel data in row 2+.")
        return

    # Build data for the editor
    data_for_editor = []
    for row_idx, row_vals in enumerate(all_values[1:], start=2):
        rank_val = row_vals[0] if len(row_vals) > 0 else ""
        name_val = row_vals[1] if len(row_vals) > 1 else ""

        existing_attendance = ""
        if len(row_vals) >= col_index:
            existing_attendance = row_vals[col_index - 1].strip()

        # If it starts with "Yes", interpret as attended
        attended_bool = existing_attendance.startswith("Yes")

        data_for_editor.append({
            "RowIndex": row_idx,
            "Rank": rank_val,
            "Name": name_val,
            "Attended": attended_bool
        })

    # ──────────────────────────────────────────────────────────────────────────
    # PUT EVERYTHING INSIDE A FORM
    # ──────────────────────────────────────────────────────────────────────────
    with st.form("safety_form", clear_on_submit=False):
        st.info("Tick 'Attended' for each person. No reruns will happen until you click 'Submit'.")
        date_input = st.text_input("Date (DDMMYYYY)")
        edited_data = st.data_editor(
            data_for_editor,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed"
        )

        # Pressing this button will "submit" the form (one-time run).
        submitted = st.form_submit_button("Update Attendance")

    if submitted:
        # The sheet is updated ONLY when this button is clicked
        if not date_input.strip():
            st.error("Please enter a date in DDMMYYYY format before submitting.")
            return

        # Validate date
        try:
            datetime.strptime(date_input, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format. Must be DDMMYYYY.")
            return

        rows_updated = 0
        for entry in edited_data:
            row_idx = entry["RowIndex"]
            # Fetch the existing value from the sheet
            existing_attendance = all_values[row_idx - 1][col_index - 1].strip() if len(all_values[row_idx - 1]) >= col_index else ""

            # Determine if there's a change in the checkbox state
            checkbox_state_changed = (entry["Attended"] and not existing_attendance.startswith("Yes")) or \
                                    (not entry["Attended"] and existing_attendance.startswith("Yes"))

            if checkbox_state_changed:
                # Update the new value based on the checkbox state
                if entry["Attended"]:
                    new_value = f"Yes, {date_input}"
                else:
                    new_value = ""

                # Update the sheet only if there's a meaningful change
                try:
                    SHEET_SAFETY.update_cell(row_idx, col_index, new_value)
                    rows_updated += 1
                except Exception as e:
                    st.error(f"Failed to update row {row_idx} for {entry['Name']}: {e}")

        if rows_updated > 0:
            st.success(
                f"Updated."
            )
        else:
            st.info("No rows were updated as there were no changes to the data.")
# ------------------------------------------------------------------------------
# 3) Helper Functions + Caching
# ------------------------------------------------------------------------------

def generate_company_message(selected_company: str, nominal_records: List[Dict], parade_records: List[Dict]) -> str:
    """
    Generate a company-specific message in the specified format.

    Parameters:
    - selected_company: The company name.
    - nominal_records: List of nominal records from Nominal_Roll.
    - parade_records: List of parade records from Parade_State.

    Returns:
    - A formatted string message.
    """
    # Define the legend-based status prefixes

    # Get current date and day of the week
    today = datetime.now(TIMEZONE)
    date_str = today.strftime("%d%m%y, %A")

    # Filter nominal records for the selected company
    company_nominal_records = [record for record in nominal_records if record['company'] == selected_company]

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
        company = parade.get('company', '')
        if company != selected_company:
            continue

        platoon = parade.get('platoon', 'Coy HQ')  # Default to 'Coy HQ' if platoon not specified

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

    # Initialize counters for total nominal and absent strengths
    total_nominal = len(company_nominal_records)
    total_absent = 0

    # Initialize storage for platoon-wise details
    platoon_details = []

    sorted_platoons = sorted(all_platoons, key=lambda x: (0, x) if x.lower() == 'coy hq' else (1, x))
    # Iterate through all platoons to gather data
    for platoon in sorted_platoons:
        records = active_parade_by_platoon.get(platoon, [])  # Get parade records for this platoon, if any

        # Determine if the platoon is 'Coy HQ' or a regular platoon
        if platoon.lower() == 'coy hq':
            platoon_label = "Coy HQ"
        else:
            platoon_label = f"Platoon {platoon}"

        # Calculate platoon nominal strength
        platoon_nominal = len([
            record for record in company_nominal_records
            if record.get('platoon', 'Coy HQ') == platoon
        ])

        # Initialize lists for categorizing absentees
        conformant_absentees = []
        non_conformant_absentees = []

        # List absentees with reasons
        for parade in records:
            name = parade.get('name', '')
            name_key = name.strip().lower()
            status = parade.get('status', '').upper()
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

            rank = name_to_rank.get(name_key, "N/A")
            # Check if the status conforms to the legend-based statuses
            status_prefix = status.lower().split()[0]  # Extract the prefix
            if status_prefix in LEGEND_STATUS_PREFIXES:
                conformant_absentees.append({
                    'rank': rank,
                    'name': name,
                    'status': status,
                    'details': details
                })
            else:
                non_conformant_absentees.append({
                    'rank': rank,
                    'name': name,
                    'status': status,
                    'details': details
                })

        # Update total_absent based on conformant absentees
        platoon_absent = len(conformant_absentees)
        total_absent += platoon_absent

        # Calculate platoon present strength based on conformant absentees
        platoon_present = platoon_nominal - platoon_absent

        # Store the details for this platoon
        platoon_details.append({
            'label': platoon_label,
            'present': platoon_present,
            'nominal': platoon_nominal,
            'absent': platoon_absent,
            'conformant': conformant_absentees,
            'non_conformant': non_conformant_absentees
        })

    # Calculate total_present after determining total_absent
    total_present = total_nominal - total_absent

    # Start building the message
    message_lines = []
    message_lines.append(f"*🐆 {selected_company.upper()} COY*")
    message_lines.append("*🗒️ FIRST PARADE STATE*")
    message_lines.append(f"*🗓️ {date_str}*\n")

    # Add the overall strength
    message_lines.append(f"Coy Present Strength: {total_present:02d}/{total_nominal:02d}")
    message_lines.append(f"Coy Absent Strength: {total_absent:02d}/{total_nominal:02d}\n")

    # Iterate through stored platoon details to build the message
    for detail in platoon_details:
        message_lines.append(f"_*{detail['label']}*_")
        message_lines.append(f"Pl Present Strength: {detail['present']:02d}/{detail['nominal']:02d}")
        message_lines.append(f"Pl Absent Strength: {detail['absent']:02d}/{detail['nominal']:02d}")

        # Add conformant absentees to the message
        if detail['conformant']:
            for absentee in detail['conformant']:
                message_lines.append(f"> {absentee['rank']} {absentee['name']} ({absentee['status'].upper().split()[0]} {absentee['details']})")

        # Add Pl Statuses count
        status_group = defaultdict(list)
        # Add non-conformant absentees if any
        if detail['non_conformant']:
            for person in detail['non_conformant']:
                rank = person['rank']
                name = person['name']
                status_code = person['status']
                details = person['details']
                key = (rank, name)
                # Combine Status Code with Details for clarity
                status_entry = f"{status_code} {details}"
                status_group[key].append(status_entry)

        pl_status_count = len(status_group)
        message_lines.append(f"\nPl Statuses: {pl_status_count:02d}/{detail['nominal']:02d}")
        if detail['non_conformant']:
            # Iterate through the grouped statuses and append consolidated lines
            for (rank, name), details_list in status_group.items():
                if rank and name:
                    line_prefix = f"> {rank} {name}"
                else:
                    line_prefix = f"> {name}"
                consolidated_details = ", ".join(details_list)
                message_lines.append(f"{line_prefix} ({consolidated_details})")

        message_lines.append("")  # Add a blank line for separation

    # Combine all lines into a single message
    final_message = "\n".join(message_lines)
    return final_message


def generate_leopards_message(all_records_nominal, all_records_parade):
    """
    Generate the Leopards-style parade message as a single string.
    Aggregates data for all companies provided in the records.
    Considers specific status prefixes (AO, LL, etc.) as absences.
    """
    # --- 1) Define the Status Prefixes and Legend Mapping ---


    # List of prefixes to check
    ABSENT_STATUS_PREFIXES = list(LEGEND_STATUS_PREFIXES.keys())

    # --- 2) Initialize Data Structures ---
    # Extract unique companies from nominal records
    company_order = sorted(list(set(row['company'] for row in all_records_nominal)))
    company_data = {comp: {"absent": [], "present_count": 0, "total_count": 0} for comp in company_order}
    overall_total = 0
    overall_absent = 0
    overall_present = 0

    # --- 3) Map Names to Their Details from Nominal Roll ---
    name_map = {}
    for row in all_records_nominal:
        name_upper = row["name"].strip().upper()
        company = row["company"]
        name_map[(name_upper, company)] = {
            "rank": row["rank"],
            "4d": row["4d_number"],
        }

    # --- 4) Determine Active Parade Statuses ---
    tz = ZoneInfo('Asia/Singapore')  # Replace with your timezone, e.g., 'America/New_York'
    today_date = datetime.now(tz).date()  # Ensure timezone-aware date
    active_parade = []
    for p in all_records_parade:
        start_str = p.get("start_date_ddmmyyyy", "")
        end_str = p.get("end_date_ddmmyyyy", "")
        status = p.get("status", "").strip().lower()
        name_val = p.get("name", "").strip().upper()
        company = p.get("company", "")
        try:
            sd = datetime.strptime(start_str, "%d%m%Y").date()
            ed = datetime.strptime(end_str, "%d%m%Y").date()
            if sd <= today_date <= ed:
                # This parade record is active
                active_parade.append({
                    "name_upper": name_val,
                    "status": status,
                    "start_dt": sd,
                    "end_dt": ed,
                    "company": company
                })
        except ValueError:
            # Skip invalid date formats
            logger.warning(
                f"Invalid date format for {name_val}: {start_str} - {end_str} in company '{company}'"
            )
            continue

    # --- 5) Map Each Person to Their Active Status ---
    person_status_map = {}
    for rec in active_parade:
        name_u = rec["name_upper"]
        company = rec["company"]
        key = (name_u, company)
        if key not in person_status_map:
            person_status_map[key] = rec  # Store the first encountered active status
        else:
            # Optional: Implement priority if multiple statuses exist
            pass

    # --- 6) Tally Present and Absent Personnel per Company ---
    for person_row in all_records_nominal:
        comp_name = person_row.get("company", "").strip()
        if comp_name not in company_data:
            logger.warning(f"Company '{comp_name}' not recognized. Skipping record for '{person_row['name']}'.")
            continue  # Skip if company is not in the list

        company_data[comp_name]["total_count"] += 1
        overall_total += 1

        name_upper = person_row["name"].strip().upper()
        rank_val = person_row["rank"]
        four_d_val = person_row["4d_number"]

        # Check if the person has an active absent status
        key = (name_upper, comp_name)
        if key in person_status_map:
            status_data = person_status_map[key]
            status = status_data["status"]
            # Check if status starts with any of the absent prefixes
            is_absent = False
            for prefix in ABSENT_STATUS_PREFIXES:
                if status.startswith(prefix):
                    legend_code = LEGEND_STATUS_PREFIXES[prefix]
                    is_absent = True
                    break
            if is_absent:
                # Format: "Rank Name [Legend Code] (Date Range)"
                start_dt = status_data["start_dt"].strftime("%d%m%y")
                end_dt = status_data["end_dt"].strftime("%d%m%y")
                date_range = f"{start_dt} - {end_dt}" if start_dt != end_dt else start_dt
                display_str = f"{rank_val} {person_row['name']} {legend_code} ({date_range})"
                company_data[comp_name]["absent"].append(display_str)
                overall_absent += 1
            else:
                company_data[comp_name]["present_count"] += 1
                overall_present += 1
        else:
            # No active absent status; person is present
            company_data[comp_name]["present_count"] += 1
            overall_present += 1

    # --- 7) Build the Message Lines ---
    lines = []
    # Header
    now_str = datetime.now(tz).strftime("%d%m%y %A, %H%M HRS")  # e.g., "160125 Thursday, 1500 HRS"
    lines.append("🐆 Leopards Parade Report")
    lines.append(f"🗓️ {now_str}\n")

    # Legend
    lines.append("> Legend")
    lines.append(" 1. [OL] - Overseas Leave")
    lines.append(" 2. [LL] - Local Leave")
    lines.append(" 3. [ML] - Medical Leave")
    lines.append(" 4. [AO] - Attached Out (use for on course too)")
    lines.append(" 5. [OIL] - Off in Lieu")
    lines.append(" 6. [MA] - Medical Appointment (Govt)")
    lines.append(" 7. [SO] - Stay Out")

    lines.append("---------------------------\n")

    # Overall Stats
    lines.append(f"🔢 Total Strength: {overall_total}")
    lines.append(f"❌ Absent: {overall_absent} pax")
    lines.append(f"✅ Present: {overall_present}\n")
    lines.append("> 🪖 Duty Personnel")
    # Duty Personnel - Omitted as per user request
    lines.append("> 🪖 Duty Personnel")
    # lines.append("- DOO: CPT ARAVIND")
    # lines.append("- BDS: 1SG RAJA")
    # lines.append("")  # Add a blank line for separation

    # Company-wise Absent Personnel
    for idx, comp_name in enumerate(company_order, start=1):
        cdata = company_data[comp_name]
        present = cdata["present_count"]
        total = cdata["total_count"]
        absent_list = cdata["absent"]
        absent_count = len(absent_list)

        lines.append(f"> {idx}️⃣ {comp_name} - {present}/{total}")

        if absent_list:
            for i, absent in enumerate(absent_list, start=1):
                lines.append(f"{i}. {absent}")
        else:
            lines.append("None")

        lines.append("")  # Add a blank line for separation

    # --- 8) Combine Lines into a Single Message ---
    final_message = "\n".join(lines)
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
    Handles case-insensitive and whitespace-trimmed headers.
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

        active_status = False
        status_desc = ""

        for parade in parade_map.get(name_key, []):
            try:
                start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
                end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
                if start_dt <= date_obj.date() <= end_dt:
                    active_status = True
                    status_desc = parade.get('status', '')
                    break
            except ValueError:
                logger.warning(
                    f"Invalid date format for {name_key}: "
                    f"{parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')}"
                )
                continue
        data.append({
            'Rank': rank,
            'Name': name,
            '4D_Number': four_d,
            'Is_Outlier': active_status,
            'StatusDesc': status_desc if active_status else ""
        })
    logger.info(f"Built conduct table with {len(data)} personnel for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
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
    st.session_state.conduct_platoon = st.selectbox(
        "Your Platoon",
        options=[1, 2, 3, 4],
        format_func=lambda x: str(x)
    )
    st.session_state.conduct_name = st.text_input(
        "Conduct Name (e.g. IPPT)",
        value=st.session_state.conduct_name
    )

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
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)

        st.session_state.conduct_table = conduct_data
        st.success(f"Loaded {len(conduct_data)} personnel for Platoon {platoon} ({date_obj.strftime('%d%m%Y')}).")
        logger.info(
            f"Loaded conduct personnel for Platoon {platoon} on {date_obj.strftime('%d%m%Y')} "
            f"in company '{selected_company}' by user '{submitted_by}'."
        )

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

    if st.button("Finalize Conduct"):
        date_str = st.session_state.conduct_date.strip()
        platoon = str(st.session_state.conduct_platoon).strip()
        cname = st.session_state.conduct_name.strip()
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
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

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
                    all_outliers.append(f"{four_d} ({status_desc})" if four_d else f"{name_} ({status_desc})")
                else:
                    all_outliers.append(f"{four_d}" if four_d else f"{name_}")

        for (rank, nm, fd, p_) in new_people:
            final_fd = fd if fd else ""
            SHEET_NOMINAL.append_row([rank, nm, final_fd, p_, 14, ""])  
            logger.info(
                f"Added new person to Nominal_Roll: Rank={rank}, Name={nm}, 4D_Number={final_fd}, "
                f"Platoon={p_} in company '{selected_company}' by user '{submitted_by}'."
            )

        total_strength_platoons = {}
        for plt in [1, 2, 3, 4]:
            strength = get_company_strength(str(plt), records_nominal)
            total_strength_platoons[plt] = strength

        pt_plts = ['0/0', '0/0', '0/0', '0/0']
        participating = 0
        for row in edited_data:
            if not row.get('Is_Outlier', False):
                participating += 1

        pt_plts[int(platoon)-1] = f"{participating}/{total_strength_platoons[int(platoon)]}"

        x_total = 0
        for pt in pt_plts:
            x = int(pt.split('/')[0]) if '/' in pt and pt.split('/')[0].isdigit() else 0
            x_total += x
        y_total = sum(total_strength_platoons.values())
        pt_total = f"{x_total}/{y_total}"

        formatted_date_str = ensure_date_str(date_str)
        SHEET_CONDUCTS.append_row([
            formatted_date_str,
            cname,
            pt_plts[0],
            pt_plts[1],
            pt_plts[2],
            pt_plts[3],
            pt_total,
            ", ".join(all_outliers) if all_outliers else "None",
            pointers,
            submitted_by
        ])
        logger.info(
            f"Appended Conduct: {formatted_date_str}, {cname}, "
            f"P/T PLT1: {pt_plts[0]}, P/T PLT2: {pt_plts[1]}, P/T PLT3: {pt_plts[2]}, "
            f"P/T PLT4: {pt_plts[3]}, P/T Total: {pt_total}, Outliers: {', '.join(all_outliers) if all_outliers else 'None'}, "
            f"Pointers: {pointers}, Submitted_By: {submitted_by} in company '{selected_company}'."
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
            SHEET_CONDUCTS.update_cell(conduct_row, 7, pt_total)
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
            f"P/T Total: {pt_total}\n"
            f"Outliers: {', '.join(all_outliers) if all_outliers else 'None'}\n"
            f"Pointers:\n{pointers if pointers else 'None'}\n"
            f"Submitted By: {submitted_by}"
        )

        st.session_state.conduct_date = ""
        st.session_state.conduct_platoon = 1
        st.session_state.conduct_name = ""
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

    conduct_index = conduct_names.index(selected_conduct) if selected_conduct in conduct_names else -1
    if conduct_index == -1:
        st.error("Selected conduct not found.")
        st.stop()



    conduct_record = records_conducts[conduct_index]

    st.subheader("Select Platoon to Update")
    selected_platoon = st.selectbox(
        "Select Platoon",
        options=[1, 2, 3, 4],
        format_func=lambda x: f"Platoon {x}",
        key="update_conduct_platoon_select"
    )

        # Initialize a session state variable to track the previous selection
    if 'update_conduct_selected_prev' not in st.session_state:
        st.session_state.update_conduct_selected_prev = None

    current_selected_conduct = selected_conduct

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
        platoon = str(selected_platoon).strip()
        date_str = conduct_record['date']
        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format in selected Conduct.")
            st.stop()

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)
        st.session_state.update_conduct_table = conduct_data
        st.success(
            f"Loaded {len(conduct_data)} personnel for Platoon {platoon} from Conduct '{selected_conduct}'."
        )
        logger.info(
            f"Loaded conduct personnel for Platoon {platoon} from Conduct '{selected_conduct}' "
            f"in company '{selected_company}' by user '{st.session_state.username}'."
        )

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

    if st.button("Update Conduct Data") and edited_data is not None:
        rows_updated = 0
        platoon = str(selected_platoon).strip()
        pt_field = f"P/T PLT{platoon}"
        new_participating = sum([1 for row in edited_data if not row.get('Is_Outlier', False)])
        new_total = len(edited_data)
        new_outliers = []
        pointers_list = []

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
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        for row in edited_data:
            four_d = is_valid_4d(row.get("4D_Number", ""))
            status_desc = ensure_str(row.get("StatusDesc", ""))
            if row.get("Is_Outlier", False):
                if status_desc:
                    new_outliers.append(f"{four_d} ({status_desc})" if four_d else row.get("Name", ""))
                else:
                    new_outliers.append(f"{four_d}" if four_d else row.get("Name", ""))

        existing_outliers = ensure_str(conduct_record.get('outliers', ''))
        if existing_outliers.lower() != 'none' and existing_outliers:
            pattern = re.compile(r'4D\d{3,4}\s*\([^)]*\)')
            existing_outliers_list = pattern.findall(existing_outliers)
            updated_outliers = ", ".join(new_outliers)
        else:
            updated_outliers = ", ".join(new_outliers) if new_outliers else "None"

        new_pt_value = f"{new_participating}/{new_total}"

        try:
            conduct_name = selected_conduct.split(" - ")[1]
            cell = SHEET_CONDUCTS.find(conduct_name, in_column=2)
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
            SHEET_CONDUCTS.update_cell(row_number, 3 + int(platoon) - 1, new_pt_value)
            logger.info(
                f"Updated {pt_field} to {new_pt_value} for conduct '{selected_conduct}' "
                f"in company '{selected_company}' by user '{st.session_state.username}'."
            )
        except Exception as e:
            st.error(f"Error updating {pt_field}: {e}")
            logger.error(f"Exception while updating {pt_field}: {e}")
            st.stop()

        try:
            SHEET_CONDUCTS.update_cell(row_number, 8, updated_outliers if updated_outliers else "None")
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
                SHEET_CONDUCTS.update_cell(row_number, 9, new_pointers)
                logger.info(
                    f"Updated Pointers to '{new_pointers}' for conduct '{selected_conduct}' "
                    f"in company '{selected_company}' by user '{st.session_state.username}'."
                )
            except Exception as e:
                st.error(f"Error updating Pointers: {e}")
                logger.error(f"Exception while updating Pointers for conduct '{selected_conduct}': {e}")
                st.stop()

        try:
            pt1 = SHEET_CONDUCTS.cell(row_number, 3).value
            pt2 = SHEET_CONDUCTS.cell(row_number, 4).value
            pt3 = SHEET_CONDUCTS.cell(row_number, 5).value
            pt4 = SHEET_CONDUCTS.cell(row_number, 6).value

            pt1_part = int(pt1.split('/')[0]) if '/' in pt1 and pt1.split('/')[0].isdigit() else 0
            pt2_part = int(pt2.split('/')[0]) if '/' in pt2 and pt2.split('/')[0].isdigit() else 0
            pt3_part = int(pt3.split('/')[0]) if '/' in pt3 and pt3.split('/')[0].isdigit() else 0
            pt4_part = int(pt4.split('/')[0]) if '/' in pt4 and pt4.split('/')[0].isdigit() else 0

            x_total = pt1_part + pt2_part + pt3_part + pt4_part
            y_total = sum([
                int(p.split('/')[1]) if '/' in p and p.split('/')[1].isdigit() else 0 
                for p in [pt1, pt2, pt3, pt4]
            ])

            pt_total = f"{x_total}/{y_total}"

            SHEET_CONDUCTS.update_cell(row_number, 7, pt_total)
            logger.info(
                f"Updated P/T Total to {pt_total} for conduct '{selected_conduct}' in company '{selected_company}' "
                f"by user '{st.session_state.username}'."
            )
        except Exception as e:
            st.error(f"Error calculating/updating P/T Total: {e}")
            logger.error(f"Exception while calculating/updating P/T Total for conduct '{selected_conduct}': {e}")
            st.stop()

        st.success(f"Conduct '{selected_conduct}' updated successfully.")
        logger.info(
            f"Conduct '{selected_conduct}' updated successfully in company '{selected_company}' "
            f"by user '{st.session_state.username}'."
        )

        st.session_state.update_conduct_pointers = [
             {"observation": "", "reflection": "", "recommendation": ""}
        ]
        # Optionally, clear the conduct table if desired

# ------------------------------------------------------------------------------
# 10) Feature C: Update Parade
# ------------------------------------------------------------------------------
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
            logger.info(
                f"Displayed current parade statuses for platoon {platoon} in company '{selected_company}' "
                f"by user '{submitted_by}'."
            )

    if st.session_state.parade_table:
        st.subheader("Edit Parade Data, Then Click 'Update'")
        st.write("Fill in 'Status', 'Start_Date (DDMMYYYY)', 'End_Date (DDMMYYYY)'")
        edited_data = st.data_editor(
            st.session_state.parade_table,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True
        )
    else:
        edited_data = None

    if st.button("Update Parade State") and edited_data is not None:
        rows_updated = 0
        platoon = str(st.session_state.parade_platoon).strip()

        records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
        records_parade = get_parade_records(selected_company, SHEET_PARADE)

        for idx, row in enumerate(edited_data):
            name_val = ensure_str(row.get("Name", "")).strip()
            status_val = ensure_str(row.get("Status", "")).strip()
            start_val = ensure_str(row.get("Start_Date", "")).strip()
            end_val = ensure_str(row.get("End_Date", "")).strip()
            four_d = is_valid_4d(row.get("4D_Number", ""))

            rank = ensure_str(row.get("Rank", "")).strip()
            parade_entry = st.session_state.parade_table[idx]
            row_num = parade_entry.get('_row_num')

            if not name_val:
                st.error(f"Name is required for row {idx}. Skipping.")
                logger.error(f"Name missing for row {idx} in company '{selected_company}'.")
                continue

            if not status_val and row_num:
                try:
                    SHEET_PARADE.delete_rows(row_num)
                    logger.info(
                        f"Deleted Parade_State row {row_num} for {name_val} in company '{selected_company}' "
                        f"by user '{submitted_by}'."
                    )
                    rows_updated += 1
                    continue
                except Exception as e:
                    st.error(f"Error deleting row for {name_val}: {e}. Skipping.")
                    logger.error(f"Exception while deleting row for {name_val}: {e}.")
                    continue

            if not status_val or not start_val or not end_val:
                logger.error(f"Missing fields for {name_val} in company '{selected_company}'. Skipping.")
                continue

            formatted_start_val = ensure_date_str(start_val)
            formatted_end_val = ensure_date_str(end_val)

            try:
                start_dt = datetime.strptime(formatted_start_val, "%d%m%Y")
                end_dt = datetime.strptime(formatted_end_val, "%d%m%Y")
                if end_dt < start_dt:
                    st.error(f"End date is before start date for {name_val}, skipping.")
                    logger.error(
                        f"End date before start date for {name_val} in company '{selected_company}'."
                    )
                    continue
            except ValueError:
                st.error(f"Invalid date(s) for {name_val}, skipping.")
                logger.error(
                    f"Invalid date format for {name_val}: Start={formatted_start_val}, End={formatted_end_val} "
                    f"in company '{selected_company}'."
                )
                continue

            # If status is "leave" and they DO have a valid 4D, update leaves
            if status_val.lower() in ["leave", "ll", "ol"]:
                if formatted_start_val and formatted_end_val:
                    if formatted_start_val != formatted_end_val:
                        dates_str = f"{formatted_start_val}-{formatted_end_val}"
                    else:
                        dates_str = formatted_start_val
                else:
                    dates_str = formatted_start_val or formatted_end_val

                logger.debug(f"Constructed dates_str for {name_val}: {dates_str}") 
                nominal_record = SHEET_NOMINAL.find(name_val, in_column=2) # Debug statement
                existing_dates = SHEET_NOMINAL.cell(nominal_record.row, 6).value
                if is_leave_accounted(existing_dates, dates_str):
                    #st.warning(f"Leave for {name_val} on {dates_str} already exists. Skipping update.")
                    logger.info(
                        f"Leave on {dates_str} for {name_val}/{four_d} already accounted for in company '{selected_company}'. Skipping."
                    )
                    continue  
                leaves_used = calculate_leaves_used(dates_str)
                if leaves_used <= 0:
                    st.error(f"Invalid leave duration for {name_val}, skipping.")
                    logger.error(
                        f"Invalid leave duration for {name_val}: {dates_str} in company '{selected_company}'."
                    )
                    continue

                if has_overlapping_status(four_d, start_dt, end_dt, records_parade):
                    logger.error(f"Leave dates overlap for {name_val}: {dates_str} in company '{selected_company}'.")
                    continue

                try:
                    nominal_record = SHEET_NOMINAL.find(name_val, in_column=2)  # 4D_Number is column C
                    if nominal_record:
                        current_leaves_left = SHEET_NOMINAL.cell(nominal_record.row, 5).value
                        try:
                            current_leaves_left = int(current_leaves_left)
                        except ValueError:
                            current_leaves_left = 14
                            logger.warning(
                                f"Invalid 'Number of Leaves Left' for {name_val}/{four_d}. Resetting to 14."
                            )

                        if leaves_used > current_leaves_left:
                            st.error(
                                f"{name_val}/{four_d} does not have enough leaves left. "
                                f"Available: {current_leaves_left}, Requested: {leaves_used}. Skipping."
                            )
                            logger.error(
                                f"{name_val}/{four_d} insufficient leaves. "
                                f"Available: {current_leaves_left}, Requested: {leaves_used}."
                            )
                            continue

                        new_leaves_left = current_leaves_left - leaves_used
                        SHEET_NOMINAL.update_cell(nominal_record.row, 5, new_leaves_left)
                        logger.info(
                            f"Updated 'Number of Leaves Left' for {name_val}/{four_d}: {new_leaves_left} "
                            f"in company '{selected_company}' by user '{submitted_by}'."
                        )

                        existing_dates = SHEET_NOMINAL.cell(nominal_record.row, 6).value
                        new_dates_entry = dates_str
                        if existing_dates:
                            # Ensure comma separation without duplication
                            if existing_dates.strip() and existing_dates.strip()[-1] != ',':
                                updated_dates = f"{existing_dates},{new_dates_entry}"
                            else:
                                updated_dates = f"{existing_dates}{new_dates_entry}"
                        else:
                            updated_dates = new_dates_entry

                        logger.debug(f"Updated 'Dates Taken' for {name_val}/{four_d}: {updated_dates}")  # Debug statement

                        SHEET_NOMINAL.update_cell(nominal_record.row, 6, updated_dates)
                        logger.info(
                            f"Updated 'Dates Taken' for {name_val}/{four_d}: {updated_dates} in company '{selected_company}' "
                            f"by user '{submitted_by}'."
                        )
                    else:
                        st.error(f"{name_val}/{four_d} not found in Nominal_Roll. Skipping.")
                        logger.error(
                            f"{name_val}/{four_d} not found in Nominal_Roll in company '{selected_company}'."
                        )
                        continue
                except Exception as e:
                    st.error(f"Error updating leaves for {name_val}/{four_d}: {e}. Skipping.")
                    logger.error(f"Exception while updating leaves for {name_val}/{four_d}: {e}.")
                    continue

            header = SHEET_PARADE.row_values(1)
            header = [h.strip().lower() for h in header]

            if row_num:
                try:
                    name_col = header.index("name") + 1
                    status_col = header.index("status") + 1
                    start_date_col = header.index("start_date_ddmmyyyy") + 1
                    end_date_col = header.index("end_date_ddmmyyyy") + 1
                    submitted_by_col = header.index("submitted_by") + 1 if "submitted_by" in header else None
                except ValueError as ve:
                    st.error(f"Required column missing in Parade_State: {ve}.")
                    logger.error(f"Required column missing in Parade_State: {ve} in company '{selected_company}'.")
                    continue

                SHEET_PARADE.update_cell(row_num, name_col, name_val)
                SHEET_PARADE.update_cell(row_num, status_col, status_val)
                SHEET_PARADE.update_cell(row_num, start_date_col, formatted_start_val)
                SHEET_PARADE.update_cell(row_num, end_date_col, formatted_end_val)

                original_entry = st.session_state.parade_table[idx]
                is_changed = (
                    row.get('Status', '') != original_entry.get('Status', '') or
                    row.get('Start_Date', '') != original_entry.get('Start_Date', '') or
                    row.get('End_Date', '') != original_entry.get('End_Date', '')
                )
                if submitted_by_col and is_changed:
                    SHEET_PARADE.update_cell(row_num, submitted_by_col, submitted_by)

                rows_updated += 1
            else:
                SHEET_PARADE.append_row([
                    platoon,
                    rank,
                    name_val,
                    four_d,
                    status_val,
                    formatted_start_val,
                    formatted_end_val,
                    submitted_by
                ])
                logger.info(
                    f"Appended Parade_State for {name_val}/{four_d}: "
                    f"Status={status_val}, Start={formatted_start_val}, End={formatted_end_val}, Submitted_By={submitted_by} "
                    f"in company '{selected_company}' by user '{submitted_by}'."
                )
                rows_updated += 1

        st.success(f"Parade State updated.")
        logger.info(
            f"Parade State updated for {rows_updated} row(s) for platoon {platoon} in company '{selected_company}' "
            f"by user '{submitted_by}'."
        )

        st.session_state.parade_platoon = 1
        st.session_state.parade_table = []

# ------------------------------------------------------------------------------
# 11) Feature D: Queries
# ------------------------------------------------------------------------------
elif feature == "Queries":
    st.subheader("Query All Medical Statuses for a Person")
    # Modified to allow either 4D or partial Name
    person_input = st.text_input("Enter the 4D Number or partial Name", key="query_person_input")
    if st.button("Get Statuses", key="btn_query_person"):
        if not person_input:
            st.error("Please enter a 4D Number or Name.")
            st.stop()

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
            valid_status_prefixes = ("ex", "rib", "ld", "mc", "ma")
            filtered_person_rows = [
                row for row in person_rows if row.get("status", "").lower().startswith(valid_status_prefixes)
            ]
            person_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("start_date_ddmmyyyy", "")))

            enhanced_rows = []
            for row in person_rows:
                # Grab 4D or empty
                four_d_val = is_valid_4d(row.get("4d_number", "")) or ""
                # Grab rank from nominal if possible
                rank_val = ""
                # We can look up 4D if it exists, or match by name if not
                if four_d_val:
                    # If 4D is valid, match by 4D
                    for nominal in records_nominal:
                        if nominal['4d_number'].upper() == four_d_val.upper():
                            rank_val = nominal['rank']
                            break
                    name_val = find_name_by_4d(row.get("4d_number", ""), records_nominal)
                else:
                    # If no 4D, match by name partially
                    name_from_parade = ensure_str(row.get("name", ""))
                    for nominal in records_nominal:
                        # For partial match, ensure the parade name is quite specific
                        # but here we'll just do a direct equality ignoring case
                        if nominal['name'].strip().lower() == name_from_parade.strip().lower():
                            four_d_val = nominal['4d_number']
                            rank_val = nominal['rank']
                            break
                    name_val = name_from_parade

                enhanced_rows.append({
                    "Rank": rank_val,
                    "Name": name_val,
                    "4D_Number": four_d_val,
                    "Status": row.get("status", ""),
                    "Start_Date": row.get("start_date_ddmmyyyy", ""),
                    "End_Date": row.get("end_date_ddmmyyyy", "")
                })

            st.subheader(f"Statuses for '{person_input}'")
            st.table(enhanced_rows)
            logger.info(f"Displayed statuses for '{person_input}' in company '{selected_company}'.")

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
            'p/t total': 'P/T Total',
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
# 12) Feature E: Overall View (With Fix)
# ------------------------------------------------------------------------------
elif feature == "Overall View":
    st.header("Overall View of All Conducts")

    conducts = get_conduct_records(selected_company, SHEET_CONDUCTS)
    if not conducts:
        st.info("No conducts available to display.")
    else:
        df = pd.DataFrame(conducts)

        def parse_date(date_str):
            try:
                return datetime.strptime(date_str, "%d%m%Y")
            except ValueError:
                return None

        df['Date'] = df['date'].apply(parse_date)
        invalid_dates = df['Date'].isnull()
        if invalid_dates.any():
            st.warning(f"{invalid_dates.sum()} conduct(s) have invalid date formats and will appear at the bottom.")
            logger.warning(f"{invalid_dates.sum()} conduct(s) have invalid date formats in company '{selected_company}'.")

        df_sorted = df.sort_values(by='Date', ascending=False)
        df_sorted['Date'] = df_sorted['Date'].dt.strftime("%d-%m-%Y")

        display_columns = {
            'date': 'Date',
            'conduct_name': 'Conduct Name',
            'p/t plt1': 'P/T PLT1',
            'p/t plt2': 'P/T PLT2',
            'p/t plt3': 'P/T PLT3',
            'p/t plt4': 'P/T PLT4',
            'p/t total': 'P/T Total',
            'outliers': 'Outliers',
            'pointers': 'Pointers',
            'submitted_by': 'Submitted By'
        }

        df_display = df_sorted.rename(columns=display_columns)[list(display_columns.values())]

        st.subheader("Filter and Sort Conducts")
        with st.expander("🔍 Filter Conducts"):
            search_term = st.text_input(
                "Search by Conduct Name or Date (DDMMYYYY):",
                value="",
                help="Enter a keyword to filter conducts by name or date."
            )
            sort_field = st.selectbox(
                "Sort By",
                options=["Date", "Conduct Name"],
                index=0
            )
            sort_order = st.radio(
                "Sort Order",
                options=["Ascending", "Descending"],
                index=1
            )

            if search_term:
                search_term_upper = search_term.upper()
                df_display = df_display[
                    df_display['Conduct Name'].str.upper().str.contains(search_term_upper) |
                    df_display['Date'].str.contains(search_term)
                ]

            ascending = True if sort_order == "Ascending" else False
            if sort_field == "Date":
                # Safely create a helper column then drop it
                if '_Date_Sort' in df_display.columns:
                    df_display.drop(columns=['_Date_Sort'], inplace=True)
                df_display['_Date_Sort'] = pd.to_datetime(df_display['Date'], format="%d-%m-%Y", errors='coerce')
                df_display.sort_values(by='_Date_Sort', ascending=ascending, inplace=True)
                df_display.drop(columns=['_Date_Sort'], inplace=True)
            elif sort_field == "Conduct Name":
                df_display = df_display.sort_values(by='Conduct Name', ascending=ascending)

        st.subheader("All Conducts")
        st.dataframe(df_display, use_container_width=True)

        st.subheader("Individuals' Missed Conducts")
        from collections import defaultdict
        missed_conducts_dict = defaultdict(list)

        for conduct in conducts:
            conduct_name = ensure_str(conduct.get('conduct_name', ''))
            outliers_str = ensure_str(conduct.get('outliers', ''))
            if outliers_str.lower() == 'none' or not outliers_str.strip():
                continue
            outliers = [o.strip() for o in outliers_str.split(',') if o.strip()]
            for outlier in outliers:
                match = re.match(r'(4D\d{3,4})(?:\s*\([^)]*\))?', outlier, re.IGNORECASE)
                if match:
                    four_d = match.group(1).upper()
                    missed_conducts_dict[four_d].append(conduct_name)

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
            df_missed = df_missed.sort_values(by="Missed Conducts Count", ascending=False)
            df_missed = df_missed.reset_index(drop=True)

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
            logger.info(
                f"No missed conducts recorded in company '{selected_company}' by user '{st.session_state.username}'."
            )

        logger.info(
            f"Displayed overall view of all conducts in company '{selected_company}' "
            f"by user '{st.session_state.username}'."
        )

# ------------------------------------------------------------------------------
# 14) Feature F: Generate WhatsApp Message
# ------------------------------------------------------------------------------
elif feature == "Generate WhatsApp Message":
    st.header("Generate WhatsApp Message")

    # --- 1) Existing WhatsApp Message Generation for Selected Company ---
    records_nominal = get_nominal_records(selected_company, SHEET_NOMINAL)
    records_parade = get_parade_records(selected_company, SHEET_PARADE)
    today = datetime.now(TIMEZONE)
    today_date = today.date()
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
            logger.warning(
                f"Invalid date format for {parade.get('name', '')}: {start_date_str} - {end_date_str} in company '{selected_company}'"
            )
            continue
    records_parade_filtered = filtered_parade

    total_strength = len(records_nominal)
    mc_count = len([
        person for person in records_parade_filtered
        if person.get('status', '').strip().lower() == 'mc' or person.get('status', '').strip().lower() == 'ml'
    ])
    current_strength = total_strength - mc_count

    mc_list = []
    statuses_list = []
    others_list = []

    for parade in records_parade_filtered:
        status = ensure_str(parade.get('status', '')).lower()
        name_val = parade.get('name', '')
        rank = ""
        four_d_val = ""

        # Find rank & 4D from nominal if possible
        for nominal in records_nominal:
            if nominal['name'].strip().lower() == name_val.strip().lower():
                rank = nominal['rank']
                four_d_val = nominal['4d_number']  # might be empty
                break

        # Build detail string
        try:
            start_dt = datetime.strptime(parade.get('start_date_ddmmyyyy', ''), "%d%m%Y").date()
            end_dt = datetime.strptime(parade.get('end_date_ddmmyyyy', ''), "%d%m%Y").date()
            delta_days = (end_dt - start_dt).days + 1
            delta_str = f"{delta_days}D"
            if delta_days == 1:
                details = f"{delta_str} {status.upper()} ({start_dt.strftime('%d%m%y')})"
            else:
                details = f"{delta_str} {status.upper()} ({start_dt.strftime('%d%m%y')}-{end_dt.strftime('%d%m%y')})"
        except ValueError:
            details = f"{status.upper()} (Invalid Dates)"
            logger.warning(
                f"Invalid dates for {name_val}: {parade.get('start_date_ddmmyyyy', '')} - {parade.get('end_date_ddmmyyyy', '')} in company '{selected_company}'"
            )

        # Store in category
        person_dict = {"Rank": rank, "Name": name_val, "4D": four_d_val, "Details": details}
        if status == 'mc' or status == 'ml':
            mc_list.append(person_dict)
        elif status in ['rib', 'ld'] or status.startswith('ex'):
            statuses_list.append(person_dict)
        else:
            others_list.append(person_dict)

    today_str = today.strftime("%d%m%y")
    platoon = "1"  # Assuming platoon is fixed as "1" for the existing message
    company_name = selected_company

    message_lines = []
    message_lines.append(f"{today_str} {platoon} SIR ({company_name.upper()}) PARADE STRENGTH\n")
    message_lines.append(f"TOTAL STR: {total_strength}")
    message_lines.append(f"CURRENT STR: {current_strength}\n")

    # MC
    message_lines.append(f"MC: {len(mc_list)}")
    for idx, person in enumerate(mc_list, start=1):
        # If 4D is not empty, append it
        line_prefix = f"{idx}. {person['Rank']}/{person['Name']}"
        if person['4D']:
            line_prefix += f"/{person['4D']}"
        message_lines.append(line_prefix)
        message_lines.append(f"* {person['Details']}")
    message_lines.append("")

    # Statuses
    oot_count = sum(1 for s in statuses_list if re.search(r'\bOOT\b', s['Details'], re.IGNORECASE))
    ex_stay_in_count = sum(1 for s in statuses_list if re.search(r'\bEX STAY IN\b', s['Details'], re.IGNORECASE))
    total_statuses = len(statuses_list)
    message_lines.append(f"STATUSES: {total_statuses} [*{oot_count} OOTS ({ex_stay_in_count} EX STAY IN)]")
    for idx, person in enumerate(statuses_list, start=1):
        line_prefix = f"{idx}. {person['Rank']}/{person['Name']}"
        if person['4D']:
            line_prefix += f"/{person['4D']}"
        message_lines.append(line_prefix)
        message_lines.append(f"* {person['Details']}")
    message_lines.append("")

    # Others
    message_lines.append(f"OTHERS: {len(others_list):02d}")
    for idx, person in enumerate(others_list, start=1):
        line_prefix = f"{idx}. {person['Rank']}/{person['Name']}"
        if person['4D']:
            line_prefix += f"/{person['4D']}"
        message_lines.append(line_prefix)
        message_lines.append(f"* {person['Details']}")

    whatsapp_message = "\n".join(message_lines)

    # --- 2) Leopards Message Generation Across All Companies ---
    # Assuming 'get_nominal_records' and 'get_parade_records' include the 'company' field
    # Aggregate records across all companies
    all_records_nominal = []
    all_records_parade = []

    for company in st.session_state.user_companies:
        worksheets = get_sheets(company)
        if not worksheets:
            st.error(f"Failed to load spreadsheets for company '{company}'.")
            continue

        SHEET_NOMINAL = worksheets["nominal"]
        SHEET_PARADE = worksheets["parade"]

        # Retrieve nominal and parade records, which now include the 'company' field
        nominal = get_nominal_records(company, SHEET_NOMINAL)
        parade = get_parade_records(company, SHEET_PARADE)

        # Append to the aggregated lists
        all_records_nominal.extend(nominal)
        all_records_parade.extend(parade)

    if not all_records_nominal:
        st.warning("No nominal records found across your companies.")
    else:
        # Generate the Leopards message
        leopards_message = generate_leopards_message(all_records_nominal, all_records_parade)


    

    tab1, tab2, tab3 = st.tabs(["BMT Level", "BN Level", "Company Level"])

    with tab1:
        st.code(whatsapp_message, language='text')
    with tab2:
        st.code(leopards_message, language='text')
    with tab3:
        # Fetch nominal and parade records for the selected company
        company_nominal = [record for record in records_nominal if record['company'] == selected_company]
        company_parade = [record for record in records_parade if record['company'] == selected_company]

        if not company_nominal:
            st.warning(f"No nominal records found for company '{selected_company}'.")
            st.stop()

        # Generate the company-specific message
        company_message = generate_company_message(selected_company, company_nominal, company_parade)
        st.code(company_message, language='text')


elif feature == "Safety Sharing":
    st.header("Safety Sharing")
    SHEET_SAFETY = worksheets["safety"]
    safety_sharing_app_form(SHEET_SAFETY)
   