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
    Parses a string of outliers from the Conduct sheet.
    E.g. "4D123 (MC), 4D234 (Excused), John Doe" =>
         {
            "4d123": { "original": "4D123", "status_desc": "MC" },
            "4d234": { "original": "4D234", "status_desc": "Excused" },
            "john doe": { "original": "John Doe", "status_desc": "" }
         }
    This version handles nested parentheses in the status description.
    """
    if existing_outliers_str.strip().lower() == "none":
        return {}
    
    def split_outliers(s):
        """Splits the string on commas that are not inside parentheses."""
        parts = []
        current = []
        depth = 0
        for char in s:
            if char == '(':
                depth += 1
            elif char == ')':
                depth -= 1
            # When encountering a comma at depth 0, finish the current part.
            if char == ',' and depth == 0:
                parts.append(''.join(current).strip())
                current = []
            else:
                current.append(char)
        if current:
            parts.append(''.join(current).strip())
        return parts

    parts = split_outliers(existing_outliers_str)
    outliers_dict = {}
    for part in parts:
        # If the part contains a status in parentheses, extract it.
        if '(' in part and part.endswith(')'):
            # Find the first '(' and assume the last ')' is the closing of the outer group.
            idx = part.index('(')
            identifier = part[:idx].strip()
            status_desc = part[idx+1:-1].strip()  # remove the outer parentheses
        else:
            identifier = part.strip()
            status_desc = ''
        outliers_dict[identifier.lower()] = {
            'original': identifier,
            'status_desc': status_desc
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
    "Bravo": "Bravo",
    "Charlie": "Charlie",
    "Viper": "Viper",
    "MSC": "MSC"  # Added MSC
}

from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Set

def is_ha_activity(conduct_name: str, for_qualification: bool = True) -> bool:
    """Check if an activity counts for HA qualification or currency."""
    qualification_base_activities = {
        'FARTLEK',
        'ENDURANCE RUN',
        'ENDURANCE TEMPO RUN',
        'DISTANCE INTERVAL',
        'STRENGTH TRAINING',
        'METABOLIC CIRCUIT',
        'INTRO TO HEARTRATE',
        'CADENCE RUN',
    }
    
    currency_base_activities = {
        'FARTLEK',
        'ENDURANCE RUN',
        'ENDURANCE TEMPO RUN',
        'DISTANCE INTERVAL',
        'ROUTE MARCH',
        'STRENGTH TRAINING',
        'METABOLIC CIRCUIT',
        'INTRO TO HEARTRATE',
        'CADENCE RUN',
    }
    
    conduct_upper = conduct_name.upper()
    
    if for_qualification:
        return any(base_activity in conduct_upper for base_activity in qualification_base_activities)
    else:
        return any(base_activity in conduct_upper for base_activity in currency_base_activities)

def parse_date(date_str: str) -> datetime:
    """Parse date from DDMMYYYY format."""
    return datetime.strptime(date_str, "%d%m%Y")

def check_standard_ha_qualification(activities: List[Tuple[datetime, str]]) -> Tuple[bool, datetime, List[str]]:
    """
    Check for standard HA qualification: 10 periods across 10 days, max 2 days break between any two consecutive periods.
    Returns (is_qualified, qualification_date, qualifying_activities).
    Stops at the first valid qualification (earliest 10 activities that qualify).
    """
    # Sort activities by date to ensure they are in order
    activities = sorted(activities, key=lambda x: x[0])
    
    for i in range(len(activities)):
        current_sequence = [activities[i]]
        for j in range(i + 1, len(activities)):
            prev_date = current_sequence[-1][0]
            curr_date = activities[j][0]
            break_days = (curr_date - prev_date).days - 1
            
            if break_days <= 2:
                current_sequence.append(activities[j])
                if len(current_sequence) >= 10:
                    # Extract the first 10 qualifying activities
                    qualifying_activities = current_sequence[:10]
                    qualification_date = qualifying_activities[-1][0]
                    return True, qualification_date, [a[1] for a in qualifying_activities]
            else:
                # Break the inner loop if the break is too long
                break
    
    return False, None, []

def check_extended_ha_qualification(activities: List[Tuple[datetime, str]]) -> Tuple[bool, datetime, List[str]]:
    """
    Check for extended HA qualification: 14 periods across 14 days,
    max 5 days total break, no more than 3 consecutive break days between any two periods.
    Returns (is_qualified, qualification_date, qualifying_activities).
    Only called if standard qualification fails.
    """
    for i in range(len(activities)):
        current_activities = []
        total_break_days = 0
        for j in range(i, len(activities)):
            if not current_activities:
                current_activities.append(activities[j])
                continue
            
            previous_date = current_activities[-1][0]
            current_date = activities[j][0]
            date_diff = (current_date - previous_date).days
            break_days = date_diff - 1
            
            # Check individual break exceeds 3 days
            if break_days > 3:
                break
            
            new_total_break = total_break_days + break_days
            # Check total break days exceed 5
            if new_total_break > 5:
                break
            
            # Update the total break days and add the activity
            total_break_days = new_total_break
            current_activities.append(activities[j])
            
            if len(current_activities) >= 14:
                return True, current_activities[-1][0], [a[1] for a in current_activities[:14]]
    
    return False, None, []

def get_ha_qualification_status_from_data(all_data, person_name: str) -> Tuple[bool, datetime, List[str], bool]:
    """Modified version of get_ha_qualification_status that works with pre-fetched data."""
    if not all_data or len(all_data) < 2:
        return False, None, [], False
    
    headers = all_data[0]
    
    # Find the person's row
    person_row = None
    for row in all_data:
        if row[1].strip().upper() == person_name.strip().upper():
            person_row = row
            break
    
    if not person_row:
        return False, None, [], False
    
    activities = []
    for header, participation in zip(headers[2:], person_row[2:]):
        if participation.strip().upper() == 'YES':
            try:
                date_str, conduct_name = header.split(',', 1)
                date_str = date_str.strip()
                conduct_name = conduct_name.strip()
                
                if is_ha_activity(conduct_name, True):
                    activities.append((parse_date(date_str), conduct_name))
            except ValueError:
                continue
    
    if not activities:
        return False, None, [], False
        
    activities.sort(key=lambda x: x[0])
    
    # First try standard HA qualification
    qualified, qual_date, qual_activities = check_standard_ha_qualification(activities)
    is_extended = False
    
    # Only try extended HA if standard fails
    if not qualified:
        qualified, qual_date, qual_activities = check_extended_ha_qualification(activities)
        is_extended = True
    
    if not qualified or not qual_date:
        return False, None, [], False
    
    is_current, nil = check_ha(all_data, person_name, qual_date)
    if not is_current:
        return False, None, [], False
    
    return True, qual_date, qual_activities, is_extended
def check_ha(all_data, person_name: str, qualification_date: datetime) -> Tuple[bool, List[str]]:
    """
    Check HA currency across periods from qualification date to latest activity.
    For completed periods: if they exist and meet requirement (2 activities in 7-day window), return True
    For incomplete periods: automatically return True
    Returns (is_current, latest_qualifying_activities).
    """
    if not all_data or len(all_data) < 2:
        return False, []
    
    headers = all_data[0]
    
    # Find person's row
    person_row = None
    for row in all_data:
        if row[1].strip().upper() == person_name.strip().upper():
            person_row = row
            break
    
    if not person_row:
        return False, []
    
    # First find the latest activity date of ANY kind
    latest_activity_date = None
    for header, participation in zip(headers[2:], person_row[2:]):
        if participation.strip().upper() == 'YES':
            try:
                date_str, _ = header.split(',', 1)
                date_str = date_str.strip()
                activity_date = parse_date(date_str)
                
                if latest_activity_date is None or activity_date > latest_activity_date:
                    latest_activity_date = activity_date
            except ValueError:
                continue
    
    if not latest_activity_date:
        return False, []
    
    # Now collect all currency-eligible activities after qualification date
    currency_activities = []
    for header, participation in zip(headers[2:], person_row[2:]):
        if participation.strip().upper() == 'YES':
            try:
                date_str, conduct_name = header.split(',', 1)
                date_str = date_str.strip()
                conduct_name = conduct_name.strip()
                activity_date = parse_date(date_str)
                
                if activity_date > qualification_date and is_ha_activity(conduct_name, False):
                    currency_activities.append((activity_date, conduct_name))
            except ValueError:
                continue
    
    if not currency_activities:
        return True, []
    
    currency_activities.sort(key=lambda x: x[0])
    first_period_start = qualification_date + timedelta(days=1)
    
    # Calculate latest complete period end date (floor to 14 days from start)
    days_since_start = (latest_activity_date - first_period_start).days
    last_complete_period_end = first_period_start + timedelta(days=((days_since_start // 14) * 14) - 1)
    
    # If we have complete periods, check the last complete one
    if days_since_start >= 14:
        period_start = last_complete_period_end - timedelta(days=13)
        period_end = last_complete_period_end
        
        # Get activities in this period
        period_activities = [
            activity for activity in currency_activities
            if period_start <= activity[0] <= period_end
        ]
        
        # Check for any 7-day window with at least 2 activities
        valid_window_found = False
        for activity in period_activities:
            window_start = activity[0]
            window_end = window_start + timedelta(days=6)
            
            window_activities = [
                a for a in period_activities
                if window_start <= a[0] <= window_end
            ]
            
            if len(window_activities) >= 2:
                valid_window_found = True
                if period_end == last_complete_period_end:
                    return True, [a[1] for a in window_activities[-2:]]
                break
        
        if not valid_window_found:
            return False, []
    
    # If we're in an incomplete period or no complete periods exist, return True
    # with the latest activities
    return True, [a[1] for a in currency_activities[-2:]]
def check_ha_currency_from_data(all_data, person_name: str, qualification_date: datetime) -> Tuple[bool, List[str]]:
    """
    Check HA currency across all periods from qualification date to latest activity.
    Returns (is_current, latest_qualifying_activities).
    
    Currency requires 2 activities within any 7-day window in each 14-day period.
    If any period fails this requirement, currency is lost.
    """
    if not all_data or len(all_data) < 2:
        return False, []
    
    headers = all_data[0]
    
    # Find person's row
    person_row = None
    for row in all_data:
        if row[1].strip().upper() == person_name.strip().upper():
            person_row = row
            break
    
    if not person_row:
        return False, []
    
    # First find the latest activity date of ANY kind
    latest_activity_date = None
    for header, participation in zip(headers[2:], person_row[2:]):
        if participation.strip().upper() == 'YES':
            try:
                date_str, _ = header.split(',', 1)
                date_str = date_str.strip()
                activity_date = parse_date(date_str)
                
                if latest_activity_date is None or activity_date > latest_activity_date:
                    latest_activity_date = activity_date
            except ValueError:
                continue
    
    if not latest_activity_date:
        return False, []
    
    # Now collect all currency-eligible activities after qualification date
    currency_activities = []
    for header, participation in zip(headers[2:], person_row[2:]):
        if participation.strip().upper() == 'YES':
            try:
                date_str, conduct_name = header.split(',', 1)
                date_str = date_str.strip()
                conduct_name = conduct_name.strip()
                activity_date = parse_date(date_str)
                
                if activity_date > qualification_date and is_ha_activity(conduct_name, False):
                    currency_activities.append((activity_date, conduct_name))
            except ValueError:
                continue
    
    if not currency_activities:
        return False, []
    
    currency_activities.sort(key=lambda x: x[0])
    first_period_start = qualification_date + timedelta(days=1)
    # Calculate all periods from qualification date to latest activity of ANY kind
    days_since_qual = (latest_activity_date - first_period_start).days
    total_periods = (days_since_qual // 14) +1
    print(days_since_qual, total_periods)
    
    # Check each period
    for period_num in range(total_periods):
        period_start = first_period_start + timedelta(days=period_num * 14)
        period_end = period_start + timedelta(days=13)  # 14 days inclusive
        print(period_start)
        print(period_end)
        # Get activities in this period
        period_activities = [
            activity for activity in currency_activities
            if period_start <= activity[0] <= period_end
        ]
        
        if not period_activities:
            return False, []
        
        # Check for any 7-day window with at least 2 activities
        valid_window_found = False
        period_activities.sort(key=lambda x: x[0])
        
        for i, activity in enumerate(period_activities):
            window_start = activity[0]
            window_end = window_start + timedelta(days=6)
            
            # Find activities within this 7-day window
            window_activities = [
                a for a in period_activities
                if window_start <= a[0] <= window_end
            ]
            
            if len(window_activities) >= 2:
                valid_window_found = True
                # If this is the last period, store these activities as the latest qualifying ones
                if period_num == total_periods - 1:
                    return True, [a[1] for a in window_activities[-2:]]
                break
        
        if not valid_window_found:
            return False, []
    
    # If we've checked all periods and haven't returned False, the person is current
    return True, [a[1] for a in period_activities[-2:]]
# Update analyze_ha_status to add requalification information
def analyze_ha_status(everything_sheet, batch_size=100, start_row=1):
    """
    Analyze HA status for personnel in the Everything sheet, with optimized batch processing.
    """
    # Get all values in a single API call
    all_values = everything_sheet.get_all_values()
    if not all_values or len(all_values) < 2:
        return []
    
    headers = all_values[0]
    
    # Determine batch of rows to process
    if batch_size is None:
        rows_to_process = all_values[1:]
    else:
        rows_to_process = all_values[start_row:start_row + batch_size + 1]
    
    results = []
    for row in rows_to_process:
        person_name = row[1].strip()
        if not person_name:
            continue
        
        # Process qualification status using the already fetched data
        qualified, qual_date, qual_activities, is_extended = get_ha_qualification_status_from_data(all_values, person_name)
        
        # Determine currency status
        is_current = False
        currency_activities = []
        currency_end_date = None
        
        if qualified and qual_date:
            two_weeks = timedelta(days=14)
            currency_end_date = qual_date + two_weeks
            
            # Check currency using the already fetched data
            is_current, currency_activities = check_ha_currency_from_data(all_values, person_name, qual_date)
        
        results.append({
            'name': person_name,
            'is_qualified': qualified,
            'qualification_date': qual_date.strftime("%d%m%Y") if qual_date else None,
            'qualifying_activities': qual_activities,
            'is_extended': is_extended,
            'is_current': is_current,
            'currency_activities': currency_activities,
            'currency_end_date': currency_end_date.strftime("%d%m%Y") if currency_end_date else None
        })
    
    return results

def display_ha_status(batch_size=100, start_row=1):
    """
    Display HA status with optional batch processing.
    
    Args:
        batch_size (int, optional): Number of rows to process in a single batch
        start_row (int, optional): Starting row index (0-based, excluding headers)
    """
    st.header("HA Status Analysis")
    
    if 'everything' not in worksheets:
        st.error("Everything sheet not found")
        return
        
    with st.spinner(f"Analyzing HA status for batch starting at row {start_row + 1}..."):
        results = analyze_ha_status(worksheets['everything'], batch_size, start_row)
        
    st.subheader(f"HA Status Summary - Batch from Row {start_row + 1}")
    
    qualified_tab, current_tab, others_tab = st.tabs([
        "HA Qualified", "HA Current", "Others"
    ])
    
    with qualified_tab:
        qualified_personnel = [r for r in results if r['is_qualified']]
        if qualified_personnel:
            for person in qualified_personnel:
                qualification_type = "Extended HA" if person['is_extended'] else "Standard HA"
                with st.expander(f"{person['name']} - {qualification_type} Qualified on {person['qualification_date']}"):
                    st.write("Qualifying Activities:")
                    for activity in person['qualifying_activities']:
                        st.write(f"- {activity}")
                    
                    now = datetime.now()
                    qual_date = datetime.strptime(person['qualification_date'], "%d%m%Y")
                    
                    # Find latest activity date and calculate current period
                    def get_latest_activity_date(all_data, person_name):
                        headers = all_data[0]
                        person_row = None
                        for row in all_data:
                            if row[1].strip().upper() == person_name.strip().upper():
                                person_row = row
                                break
                        
                        if not person_row:
                            return None
                            
                        latest_date = None
                        for header, participation in zip(headers[2:], person_row[2:]):
                            if participation.strip().upper() == 'YES':
                                try:
                                    date_str, _ = header.split(',', 1)
                                    date_str = date_str.strip()
                                    activity_date = parse_date(date_str)
                                    if latest_date is None or activity_date > latest_date:
                                        latest_date = activity_date
                                except ValueError:
                                    continue
                        return latest_date
                    
                    latest_activity = get_latest_activity_date(worksheets['everything'].get_all_values(), person['name'])
                    if latest_activity:
                        #print(latest_activity)
                        days_since_qual = (latest_activity - qual_date).days - 1
                        #print(days_since_qual)
                        if days_since_qual == -1:
                            days_since_qual = 0
                        #print(days_since_qual)
                        current_period_number = days_since_qual // 14
                        
                        # Calculate the period boundaries
                        period_start = qual_date + timedelta(days=current_period_number * 14)
                        period_end = period_start + timedelta(days=14)
                        next_period_end = period_end + timedelta(days=14)
                        
                        # Format dates for comparison
                        period_end_str = period_end.strftime("%d%m%Y")
                        next_period_end_str = next_period_end.strftime("%d%m%Y")
                        latest_activity_str = latest_activity.strftime("%d%m%Y")
                        
                        if person['currency_end_date']:
                            if person['is_current']:
                                st.success("HA Currency: Maintained")
                                st.write("Current Period Currency Activities:")
                                for activity in person['currency_activities']:
                                    st.write(f"- {activity}")
                            else:
                                # Compare dates to determine which end date to show
                                latest_activity_date = datetime.strptime(latest_activity_str, "%d%m%Y")
                                period_end_date = datetime.strptime(period_end_str, "%d%m%Y")
                                
                                if latest_activity_date <= period_end_date:
                                    display_date = period_end_str
                                else:
                                    display_date = next_period_end_str
                                    
                                st.warning(f"HA Currency: Not Done - Still Have Time Until {display_date}")
                                st.info("Need 2 activity periods within 7 days to maintain currency")
                        else:
                            st.error("HA Qualification Error: No Currency End Date")
                    else:
                        st.error("Could not determine latest activity date")
                        
        else:
            st.info("No personnel have achieved HA qualification in this batch.")
            
    with current_tab:
        current_personnel = [r for r in results if r['is_qualified'] and r['is_current']]
        if current_personnel:
            for person in current_personnel:
                qualification_type = "Extended HA" if person['is_extended'] else "Standard HA"
                with st.expander(f"{person['name']} - {qualification_type}"):
                    st.write(f"Qualified on: {person['qualification_date']}")
                    st.write("Current Period Currency Activities:")
                    for activity in person['currency_activities']:
                        st.write(f"- {activity}")
        else:
            st.info("No personnel are currently maintaining HA currency in this batch.")
            
    with others_tab:
        other_personnel = [r for r in results if not r['is_qualified']]
        if other_personnel:
            for person in other_personnel:
                st.write(f"- {person['name']}: Not yet qualified")
        else:
            st.info("All personnel in this batch have achieved HA qualification.")
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
            "safety": sh.worksheet("Safety"),
            "everything": sh.worksheet("Everything"),
            "progressive": sh.worksheet("Progressive")
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

def safety_sharing_app_form(SHEET_SAFETY, SHEET_PARADE, selected_company):
    # === 1) Read Safety Sheet Header ===
    header_row = SHEET_SAFETY.row_values(1)
    if not header_row or len(header_row) < 3:
        st.error("Safety sheet must have at least 3 columns: 'Rank', '4D_Number', 'Name'.")
        return
    fixed_cols = header_row[:3]            # Fixed columns: Rank, 4D_Number, Name.
    existing_cols = header_row[3:]           # Attendance columns.

    # === 2) Choose or Create a Safety Column ===
    if "safety_selected_col" not in st.session_state:
        st.session_state.safety_selected_col = None

    st.subheader("Choose or Create a Safety Column")
    if st.session_state.safety_selected_col:
        st.info(f"Selected column: **{st.session_state.safety_selected_col}**")
        if st.button("Change/Reset Column"):
            st.session_state.safety_selected_col = None
            st.rerun()
    else:
        mode_choice = st.radio("Mode:", ["Select Existing", "Create New"], horizontal=True)
        if mode_choice == "Select Existing":
            if not existing_cols:
                st.warning("No existing safety columns found. Please create one.")
            else:
                selected = st.selectbox("Select a column:", options=existing_cols)
                if st.button("Use This Column"):
                    st.session_state.safety_selected_col = selected
                    st.rerun()
        else:
            # New column: include week number and description.
            week_number = st.number_input("Enter week number", min_value=1, step=1, value=1)
            description = st.text_input("Enter description (optional)", value="")
            if st.button("Create New Column"):
                header_value = f"Week {int(week_number)}"
                if description.strip():
                    header_value += f" ({description.strip()})"
                try:
                    current_header = SHEET_SAFETY.row_values(1)
                    new_col_index = len(current_header) + 1
                    SHEET_SAFETY.update_cell(row=1, col=new_col_index, value=header_value)
                    st.success(f"Created new column '{header_value}'")
                    st.session_state.safety_selected_col = header_value
                    st.rerun()
                except Exception as e:
                    st.error(f"Error creating column: {e}")
                    return

    selected_col = st.session_state.safety_selected_col
    updated_header = SHEET_SAFETY.row_values(1)
    if selected_col not in updated_header:
        #st.error("Selected safety column not found in the sheet header.")S
        return
    col_index = updated_header.index(selected_col) + 1  # 1-based index

    # === 3) Operation Mode Selection ===
    st.subheader("Operation Mode")
    op_mode = st.radio("Select Mode:", ["New Safety Column", "Update Safety Column"])
    st.write(f"Mode selected: **{op_mode}**")

    # === 4) Date Input (only for New mode) ===
    if op_mode == "New Safety Column":
        date_input = st.text_input("Enter date (DDMMYYYY)", value="")
    else:
        date_input = None  # No global date input for update mode.

    # === 5) Load Personnel Data ===
    if op_mode == "New Safety Column":
        # In New mode, the column should be blank.
        if not date_input.strip():
            st.error("Please enter a date for the new safety column.")
            return
        try:
            safety_date = datetime.strptime(date_input.strip(), "%d%m%Y").date()
        except ValueError:
            st.error("Invalid date format. Please use DDMMYYYY.")
            return

        # Load parade data.
        parade_data = SHEET_PARADE.get_all_values()
        parade_attendees = set()
        if parade_data and len(parade_data) >= 2:
            parade_header = [h.strip().lower() for h in parade_data[0]]
            try:
                name_idx = parade_header.index("name")
                start_idx = parade_header.index("start_date_ddmmyyyy")
                end_idx = parade_header.index("end_date_ddmmyyyy")
                status_idx = parade_header.index("status")
                for row in parade_data[1:]:
                    if len(row) <= max(name_idx, start_idx, end_idx):
                        continue
                    try:
                        start_dt = datetime.strptime(row[start_idx].strip(), "%d%m%Y").date()
                        end_dt = datetime.strptime(row[end_idx].strip(), "%d%m%Y").date()
                        status_val = row[status_idx].strip().upper()
                        if start_dt <= safety_date <= end_dt:
                            # Person appears in parade data → on status → NOT attended.
                            status_prefix = status_val.lower().split()[0]
                            if status_prefix in LEGEND_STATUS_PREFIXES:
                                parade_attendees.add(row[name_idx].strip().upper())
                    except ValueError:
                        continue
            except ValueError:
                st.error("Parade sheet missing required columns ('name', 'start_date_ddmmyyyy', 'end_date_ddmmyyyy').")
        else:
            st.warning("No parade data found.")

        # Load Safety sheet data.
        safety_values = SHEET_SAFETY.get_all_values()
        editor_data = []
        for i, row in enumerate(safety_values[1:], start=2):
            rank_val = row[0] if len(row) >= 1 else ""
            four_d_val = row[1] if len(row) >= 2 else ""
            name_val = row[2] if len(row) >= 3 else ""
            # For new mode, ignore any existing cell value.
            # Pre-populate: if the person's name is in parade_attendees, mark as NOT attended; else attended.
            attended = (name_val.strip().upper() not in parade_attendees)
            editor_data.append({
                "RowIndex": i,
                "Rank": rank_val,
                "4D_Number": four_d_val,
                "Name": name_val,
                "Attended": attended
            })
        st.session_state.safety_editor_data = editor_data
        st.session_state.safety_date = date_input.strip()
        st.success(f"Loaded {len(editor_data)} records for New Safety Column.")
    else:
        # Update mode: load current Safety sheet data.
        safety_values = SHEET_SAFETY.get_all_values()
        editor_data = []
        for i, row in enumerate(safety_values[1:], start=2):
            rank_val = row[0] if len(row) >= 1 else ""
            four_d_val = row[1] if len(row) >= 2 else ""
            name_val = row[2] if len(row) >= 3 else ""
            cell_val = row[col_index - 1].strip() if len(row) >= col_index else ""
            if cell_val.startswith("Yes,"):
                # Extract the date part (everything after "Yes,")
                date_val = cell_val[4:].strip()
                attended = True
            else:
                date_val = ""
                attended = False
            # In update mode, we simply load the current value.
            editor_data.append({
                "RowIndex": i,
                "Rank": rank_val,
                "4D_Number": four_d_val,
                "Name": name_val,
                "Attended": attended,
                "Date": date_val  # This field is shown for reference/editing if needed.
            })
        st.session_state.safety_editor_data = editor_data
        st.success(f"Loaded {len(editor_data)} records for Update Safety Column.")

    # === 6) Display Editor & Batch Update ===
    if "safety_editor_data" in st.session_state:
        st.subheader("Review & Update Attendance")
        edited_data = st.data_editor(
            st.session_state.safety_editor_data,
            key="safety_editor",
            use_container_width=True,
            hide_index=True,
            num_rows="fixed"
        )
        st.session_state.safety_editor_data = edited_data

        if st.button("Update Attendance"):
            update_requests = []
            if op_mode == "New Safety Column":
                # Every row is updated with the global date.
                date_str = st.session_state.safety_date
                for entry in st.session_state.safety_editor_data:
                    row_idx = entry["RowIndex"]
                    new_attended = entry["Attended"]
                    new_value = f"Yes, {date_str}" if new_attended else ""
                    update_requests.append({
                        "range": gspread.utils.rowcol_to_a1(row_idx, col_index),
                        "values": [[new_value]]
                    })
            else:
                # Update mode: update only rows that have changed.
                current_col_values = SHEET_SAFETY.col_values(col_index)[1:]  # Skip header.
                for entry in st.session_state.safety_editor_data:
                    row_idx = entry["RowIndex"]
                    new_attended = entry["Attended"]
                    # For update mode, use the per-row Date field.
                    new_date = entry["Date"].strip()
                    new_value = f"Yes, {new_date}" if new_attended else ""
                    orig_value = current_col_values[row_idx - 2] if (row_idx - 2) < len(current_col_values) else ""
                    if new_value != orig_value:
                        update_requests.append({
                            "range": gspread.utils.rowcol_to_a1(row_idx, col_index),
                            "values": [[new_value]]
                        })
            if update_requests:
                try:
                    SHEET_SAFETY.batch_update(update_requests)
                    st.success("Attendance updated successfully via batch update.")
                except Exception as e:
                    st.error(f"Batch update failed: {e}")
            else:
                st.info("No updates to apply.")
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



def add_conduct_column_progressive(sheet_progressive, conduct_date: str, conduct_name: str, attendance_data: List[tuple]):
    """
    Adds a new column (DATE, ConductName) to the 'Progressive' sheet, updates the Yes/No attendance,
    then REARRANGES the columns so that conduct columns are grouped/sorted by 'conduct_name'.
    Only keeps columns that match specified conduct names.

    :param sheet_progressive: gspread Worksheet object for the 'Progressive' sheet.
    :param conduct_date: e.g. '15012025'
    :param conduct_name: e.g. 'RESILIENCE LEARNING 4: Goal setting'
    :param attendance_data: list of tuples (name, rank, is_present)
    """
    # Define allowed conduct prefixes/names
    ALLOWED_CONDUCTS = {
        "FARTLEK",
        "Intro to Heartrate",
        "Speed Agility Quickness",
        "Metabolic circuit",
        "STRENGTH TRAINING",
        "AQUA",
        "DISTANCE INTERVAL",
        "ENDURANCE RUN",
        "ROUTE MARCH(3KM)",
        "BALANCING, FLEXIBILITY, MOBILITY",
        "CADENCE RUN",
        "GYM ORIENTATION",
        "GYM TRAINING",
        "ENDURANCE RUN TEMPO",
        "CPT"
    }

    def is_allowed_conduct(conduct_name: str) -> bool:
        """Check if a conduct name matches any of the allowed conducts"""
        conduct_upper = conduct_name.upper()
        return any(
            allowed.upper() in conduct_upper or conduct_upper.startswith(allowed.upper())
            for allowed in ALLOWED_CONDUCTS
        )

    # 1) Build new header label: "DDMMYYYY, CONDUCT_NAME"
    new_col_header = f"{conduct_date}, {conduct_name}"

    # 2) Read ALL data from Progressive sheet
    all_data = sheet_progressive.get_all_values()
    if not all_data:
        raise ValueError("No data found in Progressive sheet")

    # 3) Current number of columns
    num_cols = len(all_data[0])

    # 4) Insert (append) the new header at the end of the header row
    sheet_progressive.update_cell(1, num_cols + 1, new_col_header)

    # 5) Create a {name: is_present} map for easy lookup
    attendance_map = {name: is_present for (name, _, is_present) in attendance_data}

    # 6) Fill the new column (temporary location at the far right)
    updates = []
    for row_idx, row in enumerate(all_data[1:], start=2):
        if len(row) < 3:
            continue
        name_in_sheet = row[2].strip()
        value = "Yes" if attendance_map.get(name_in_sheet, False) else "No"
        cell_a1 = gspread.utils.rowcol_to_a1(row_idx, num_cols + 1)
        updates.append({"range": cell_a1, "values": [[value]]})

    if updates:
        sheet_progressive.batch_update(updates)

    # Read updated data
    all_data_updated = sheet_progressive.get_all_values()
    if not all_data_updated:
        raise ValueError("No data found after adding the new conduct column in Progressive.")

    header_row = all_data_updated[0]
    body_rows = all_data_updated[1:]

    start_of_conduct_cols = 3
    fixed_cols = header_row[:start_of_conduct_cols]
    conduct_cols = header_row[start_of_conduct_cols:]

    def get_conduct_name(hdr: str) -> str:
        parts = hdr.split(",", 1)
        if len(parts) == 2:
            return parts[1].strip()
        return hdr.strip()

    # Filter out conduct columns that don't match allowed conducts
    filtered_conduct_cols = [
        col for col in conduct_cols
        if is_allowed_conduct(get_conduct_name(col))
    ]

    # Sort the filtered conduct columns
    conduct_cols_sorted = sorted(
        filtered_conduct_cols,
        key=lambda h: get_conduct_name(h).upper()
    )

    # Map old indices for the filtered and sorted columns
    old_conduct_index_map = {hdr: i for i, hdr in enumerate(conduct_cols, start=start_of_conduct_cols)}

    # Build new header with only allowed conducts
    new_header = list(fixed_cols) + conduct_cols_sorted

    # Reconstruct matrix with only allowed conducts
    new_matrix = [new_header]

    for row_idx, row_values in enumerate(body_rows, start=1):
        new_row = row_values[:start_of_conduct_cols]

        for sorted_hdr in conduct_cols_sorted:
            old_abs_index = old_conduct_index_map[sorted_hdr]
            cell_val = ""
            if old_abs_index < len(row_values):
                cell_val = row_values[old_abs_index]
            new_row.append(cell_val)

        new_matrix.append(new_row)

    # Update the sheet with filtered data
    sheet_progressive.clear()
    sheet_progressive.update("A1", new_matrix)


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
    ["Add Conduct", "Update Conduct", "Update Parade", "Queries", "Overall View", "Generate WhatsApp Message", "Safety Sharing", "HA"]
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
        "BALANCING, FLEXIBILITY, MOBILITY",
        "BPK",
        "BSK",
        "CBA&CPR-AED",
        "CADENCE RUN",
        "COY FIRE DRILL",
        "CPT",
        "DISTANCE INTERVAL",
        "ELISS FAMILIARISATION",
        "ENDURANCE RUN",
        "ENDURANCE RUN TEMPO",
        "FARTLEK",
        "FOOT DRILLS",
        "GYM ORIENTATION",
        "GYM TRAINING",
        "INFANTRY SMALL ARMS DEMONSTRATION",
        "INTRO TO HEARTRATE",
        "IPPT",
        "LEADERSHIP VALUES",
        "MO TALK",
        "METABOLIC CIRCUIT",
        "NATIONAL EDUCATION",
        "OO ENGAGEMENT",
        "ORIENTATION RUN",
        "PHYSICAL TRAINING LECTURE",
        "RESILIENCE LEARNING",
        "ROUTE MARCH(3KM)",
        "SAFE & INCLUSIVE WORKPLACE",
        "SAFRA TALK",
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
        # Updated to include 'Coy HQ'
        for plt in platoon_options:
            strength = get_company_strength(plt, records_nominal)
            total_strength_platoons[plt] = strength
            print(total_strength_platoons[plt])

        # Initialize pt_plts with 'Coy HQ'
        pt_plts = ['0/0', '0/0', '0/0', '0/0', '0/0']
        participating = 0
        for row in edited_data:
            if not row.get('Is_Outlier', False):
                participating += 1

        if platoon in platoon_options:
            if platoon != "Coy HQ":
                index = int(platoon) - 1  # Platoons 1-4 map to indices 0-3
            else:
                index = 4  # 'Coy HQ' maps to index 4
            pt_plts[index] = f"{participating}/{total_strength_platoons[platoon]}"

        x_total = 0
        for pt in pt_plts:
            x = int(pt.split('/')[0]) if '/' in pt and pt.split('/')[0].isdigit() else 0
            x_total += x
        y_total = sum(total_strength_platoons.values())
        pt_total = f"{x_total}/{y_total}"

        formatted_date_str = ensure_date_str(date_str)
                # Prepare outliers per platoon – order: PLT1, PLT2, PLT3, PLT4, Coy HQ
        outliers_list = ["None"] * 5
        if platoon in platoon_options:
            index = int(platoon) - 1 if platoon != "Coy HQ" else 4
            outliers_list[index] = ", ".join(all_outliers) if all_outliers else "None"

        SHEET_CONDUCTS.append_row([
            formatted_date_str,  # Column 1: Date
            cname,               # Column 2: Conduct_Name
            pt_plts[0],          # Column 3: P/T PLT1
            pt_plts[1],          # Column 4: P/T PLT2
            pt_plts[2],          # Column 5: P/T PLT3
            pt_plts[3],          # Column 6: P/T PLT4
            pt_plts[4],          # Column 7: P/T Coy HQ
            pt_total,            # Column 8: P/T Total
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
        add_conduct_column_progressive(
            sheet_progressive= worksheets["progressive"],  # or however you reference the sheet
            conduct_date=formatted_date_str,
            conduct_name=cname,
            attendance_data=attendance_data
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

    conduct_index = conduct_names.index(selected_conduct) if selected_conduct in conduct_names else -1
    if conduct_index == -1:
        st.error("Selected conduct not found.")
        st.stop()



    conduct_record = records_conducts[conduct_index]

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

        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)

                # 1) Determine the outlier column for this platoon
        if platoon != "Coy HQ":
            outlier_key = f"plt{platoon} outliers"  # all lower
        else:
            outlier_key = "coy hq outliers"

        # 2) Get the existing outlier string from the conduct_record
        existing_outliers_str = conduct_record.get(outlier_key, "")  # or use ensure_str() if you want
        existing_outliers = parse_existing_outliers(existing_outliers_str)
        print("DEBUG outlier_key:", outlier_key)
        print("DEBUG existing_outliers_str:", existing_outliers_str)
        # 3) Make a helper to find rows in the conduct_data
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

        # 4) Merge existing outliers into the table
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
            else:
                # If not in the table, create a brand new row
                new_row = {
                    "4D_Number": identifier_original if "4d" in identifier_original.lower() else "",
                    "Name": identifier_original if "4d" not in identifier_original.lower() else "",
                    "Rank": "",
                    "Platoon": platoon,
                    "Status": "",
                    "StatusDesc": status_desc,
                    "Is_Outlier": True
                }
                conduct_data.append(new_row)



        st.session_state.update_conduct_table = conduct_data
        st.success(
            f"Loaded {len(conduct_data)} personnel for Platoon {platoon} from Conduct '{selected_conduct}'."
        )
        logger.info(
            f"Loaded conduct personnel for Platoon {platoon} from Conduct '{selected_conduct}' "
            f"in company '{selected_company}' by user '{st.session_state.username}'."
        )

    if "update_conduct_table" in st.session_state and st.session_state.update_conduct_table:
        st.subheader(f"Edit Conduct Data for Platoon {st.session_state.conduct_platoon}")
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
        platoon = str(st.session_state.conduct_platoon).strip()
        pt_field = f"P/T PLT{platoon}"
        new_participating = sum([1 for row in edited_data if not row.get('Is_Outlier', False)])
        new_total = len(edited_data)
        new_outliers = []
        pointers_list = []

        SHEET_EVERYTHING = worksheets["everything"]
        SHEET_PROGRESSIVE = worksheets["progressive"]
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
        update_conduct_column_everything(
            SHEET_PROGRESSIVE,
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

        def reconstruct_outliers(outliers_dict):
            """
            Reconstructs the outliers string from the dictionary.
            
            Returns:
                str: Comma-separated outliers.
            """
            outliers_list = []
            for entry in outliers_dict.values():
                if entry['status_desc']:
                    outliers_list.append(f"{entry['original']} ({entry['status_desc']})")
                else:
                    outliers_list.append(f"{entry['original']}")
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
            
            updated_outliers = reconstruct_outliers(existing_outliers)
            return updated_outliers
        updated_outliers = update_outliers(edited_data, conduct_record, platoon)

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
            pt1 = SHEET_CONDUCTS.cell(row_number, 3).value
            pt2 = SHEET_CONDUCTS.cell(row_number, 4).value
            pt3 = SHEET_CONDUCTS.cell(row_number, 5).value
            pt4 = SHEET_CONDUCTS.cell(row_number, 6).value
            pt5 = SHEET_CONDUCTS.cell(row_number, 7).value

            pt1_part = int(pt1.split('/')[0]) if '/' in pt1 and pt1.split('/')[0].isdigit() else 0
            pt2_part = int(pt2.split('/')[0]) if '/' in pt2 and pt2.split('/')[0].isdigit() else 0
            pt3_part = int(pt3.split('/')[0]) if '/' in pt3 and pt3.split('/')[0].isdigit() else 0
            pt4_part = int(pt4.split('/')[0]) if '/' in pt4 and pt4.split('/')[0].isdigit() else 0
            pt5_part = int(pt5.split('/')[0]) if '/' in pt5 and pt5.split('/')[0].isdigit() else 0

            x_total = pt1_part + pt2_part + pt3_part + pt4_part + pt5_part
            y_total = sum([
                int(p.split('/')[1]) if '/' in p and p.split('/')[1].isdigit() else 0 
                for p in [pt1, pt2, pt3, pt4, pt5]
            ])

            pt_total = f"{x_total}/{y_total}"

            SHEET_CONDUCTS.update_cell(row_number, 8, pt_total)
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
        edited_data = st.data_editor(
            st.session_state.parade_table,
            num_rows="dynamic",
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
            valid_status_prefixes = ("ex", "rib", "ld", "mc", "ml")
            filtered_person_rows = [
                row for row in person_rows if row.get("status", "").lower().startswith(valid_status_prefixes)
            ]
            print(filtered_person_rows)
            filtered_person_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("start_date_ddmmyyyy", "")))

            enhanced_rows = []
            for row in filtered_person_rows:
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
        selected_date = st.date_input("Select Parade Date", datetime.now(TIMEZONE).date())
        target_datetime = datetime.combine(selected_date, datetime.min.time())
        # Fetch nominal and parade records for the selected company
        company_nominal = [record for record in records_nominal if record['company'] == selected_company]
        company_parade = [record for record in records_parade if record['company'] == selected_company]

        if not company_nominal:
            st.warning(f"No nominal records found for company '{selected_company}'.")
            st.stop()

        # Generate the company-specific message
        company_message = generate_company_message(selected_company, company_nominal, company_parade, target_date=target_datetime)
        st.code(company_message, language='text')


elif feature == "Safety Sharing":
    st.header("Safety Sharing")
    SHEET_SAFETY = worksheets["safety"]
    safety_sharing_app_form(SHEET_SAFETY, SHEET_PARADE, selected_company)

   

elif feature == "HA":
    display_ha_status()
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
            'submitted_by': 'Submitted By'
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
