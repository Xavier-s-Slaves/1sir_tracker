import streamlit as st  # type: ignore
import gspread  # type: ignore
from oauth2client.service_account import ServiceAccountCredentials  # type: ignore
from datetime import datetime, timedelta
from collections import defaultdict
import difflib
import re
import logging

# ------------------------------------------------------------------------------
# Setup Logging
# ------------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
    "Support": "Support"
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

@st.cache_data(ttl=300)
def get_nominal_records(_sheet_nominal):
    """
    Returns all rows from Nominal_Roll as a list of dicts, cached for 5 min.
    The parameter '_sheet_nominal' is ignored in hashing due to the leading underscore.
    """
    records = _sheet_nominal.get_all_records()
    # Ensure all relevant fields are strings and properly formatted
    for row in records:
        row['4D_Number'] = is_valid_4d(row.get('4D_Number', ''))
        row['Platoon'] = ensure_str(row.get('Platoon', ''))
        row['Name'] = ensure_str(row.get('Name', ''))
    # Remove any records with invalid 4D_Number
    records = [row for row in records if row['4D_Number']]
    return records

@st.cache_data(ttl=300)
def get_parade_records(_sheet_parade):
    """
    Returns all rows from Parade_State as a list of dicts, including row numbers, cached.
    The parameter '_sheet_parade' is ignored in hashing due to the leading underscore.
    """
    all_values = _sheet_parade.get_all_values()  # includes header row at index 0
    records = []
    header = all_values[0]
    for idx, row in enumerate(all_values[1:], start=2):  # Start at row 2 in Google Sheets
        record = dict(zip(header, row))
        # Include row number for updating
        record['_row_num'] = idx
        # Ensure all relevant fields are strings and properly formatted
        record['4D_Number'] = is_valid_4d(record.get('4D_Number', ''))
        record['Platoon'] = ensure_str(record.get('Platoon', ''))
        record['Start_Date_DDMMYYYY'] = ensure_date_str(record.get('Start_Date_DDMMYYYY', ''))
        record['End_Date_DDMMYYYY'] = ensure_date_str(record.get('End_Date_DDMMYYYY', ''))
        record['Status'] = ensure_str(record.get('Status', ''))
        # Remove any records with invalid 4D_Number
        if record['4D_Number']:
            records.append(record)
    return records

@st.cache_data(ttl=300)
def get_conduct_records(_sheet_conducts):
    """
    Returns all rows from Conducts as a list of dicts, cached.
    The parameter '_sheet_conducts' is ignored in hashing due to the leading underscore.
    """
    records = _sheet_conducts.get_all_records()
    # Ensure all relevant fields are strings and properly formatted
    for row in records:
        row['Date'] = ensure_date_str(row.get('Date', ''))
        row['Platoon'] = ensure_str(row.get('Platoon', ''))
        row['Conduct_Name'] = ensure_str(row.get('Conduct_Name', ''))
        row['Outliers'] = ensure_str(row.get('Outliers', ''))
    return records

def get_company_strength(platoon: str, records_nominal):
    """
    Count how many rows in Nominal_Roll belong to that platoon.
    Uses cached data to avoid repeated API calls.
    """
    return sum(
        1 for row in records_nominal
        if normalize_name(row.get('Platoon', '')) == normalize_name(platoon)
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
        four_d = row.get('4D_Number', '').strip().upper()
        parade_map[four_d].append(row)
    
    data_with_status = []
    data_nominal = []
    
    for row in records_nominal:
        p = row.get('Platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        four_d = row.get('4D_Number', '').strip().upper()
        name = row.get('Name', '')
        # Retrieve all parade statuses for the person
        person_parades = parade_map.get(four_d, [])
        for parade in person_parades:
            data_with_status.append({
                'Name': name,
                '4D_Number': four_d,
                'Status': parade.get('Status', ''),
                'Start_Date': parade.get('Start_Date_DDMMYYYY', ''),
                'End_Date': parade.get('End_Date_DDMMYYYY', ''),
                'Number_of_Leaves_Left': row.get('Number of Leaves Left', 14),
                'Dates_Taken': row.get('Dates Taken', ''),
                '_row_num': parade.get('_row_num')  # Track row number for updating
            })
        # Add the nominal entry without status
        data_nominal.append({
            'Name': name,
            '4D_Number': four_d,
            'Status': '',
            'Start_Date': '',
            'End_Date': '',
            'Number_of_Leaves_Left': row.get('Number of Leaves Left', 14),
            'Dates_Taken': row.get('Dates Taken', ''),
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
        if ensure_str(row.get("4D_Number", "")).upper() == four_d:
            return ensure_str(row.get("Name", ""))
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
        four_d = row.get('4D_Number', '').strip().upper()
        parade_map[four_d].append(row)
    
    for row in records_nominal:
        p = row.get('Platoon', '')
        if normalize_name(p) == normalize_name(platoon):
            four_d = row.get('4D_Number', '').strip().upper()
            name = row.get('Name', '')
            # Retrieve all parade statuses for the person
            person_parades = parade_map.get(four_d, [])
            for parade in person_parades:
                try:
                    start_dt = datetime.strptime(parade.get('Start_Date_DDMMYYYY', '01012000'), "%d%m%Y")
                    end_dt = datetime.strptime(parade.get('End_Date_DDMMYYYY', '01012000'), "%d%m%Y")
                except ValueError:
                    logger.warning(f"Invalid date format in Parade_State for {four_d}: {parade.get('Start_Date_DDMMYYYY', '')} - {parade.get('End_Date_DDMMYYYY', '')}")
                    continue

                if start_dt <= date_obj <= end_dt:
                    status = ensure_str(parade.get('Status', '')).lower()
                    if not four_d:
                        continue  # Skip invalid 4D_Number
                    # If multiple statuses, keep the one with higher priority
                    if four_d in out:
                        existing_status = out[four_d]['StatusDesc'].lower()
                        if status_priority.get(status, 0) > status_priority.get(existing_status, 0):
                            out[four_d] = {
                                "Name": find_name_by_4d(four_d, records_nominal),
                                "4D_Number": four_d,
                                "StatusDesc": ensure_str(parade.get('Status', '')),
                                "Is_Outlier": True
                            }
                    else:
                        out[four_d] = {
                            "Name": find_name_by_4d(four_d, records_nominal),
                            "4D_Number": four_d,
                            "StatusDesc": ensure_str(parade.get('Status', '')),
                            "Is_Outlier": True
                        }
    logger.info(f"Built on-status table with {len(out)} entries for platoon {platoon} on {date_obj.strftime('%d%m%Y')}.")
    return list(out.values())

def build_conduct_table(platoon: str, date_obj: datetime, records_nominal, records_parade):
    """
    Return a list of dicts for all personnel in the platoon.
    'Is_Outlier' is True if the person has an active status on the given date, else False.
    Includes 'StatusDesc' for personnel on status.
    """
    parade_map = defaultdict(list)
    for row in records_parade:
        four_d = row.get('4D_Number', '').strip().upper()
        parade_map[four_d].append(row)
    
    data = []
    for person in records_nominal:
        p = person.get('Platoon', '')
        if normalize_name(p) != normalize_name(platoon):
            continue
        four_d = person.get('4D_Number', '').strip().upper()
        name = person.get('Name', '')
        # Check if person has an active parade status on the given date
        active_status = False
        status_desc = ""
        for parade in parade_map.get(four_d, []):
            try:
                start_dt = datetime.strptime(parade.get('Start_Date_DDMMYYYY', ''), "%d%m%Y")
                end_dt = datetime.strptime(parade.get('End_Date_DDMMYYYY', ''), "%d%m%Y")
                if start_dt <= date_obj <= end_dt:
                    active_status = True
                    status_desc = parade.get('Status', '')
                    break
            except ValueError:
                logger.warning(f"Invalid date format for {four_d}: {parade.get('Start_Date_DDMMYYYY', '')} - {parade.get('End_Date_DDMMYYYY', '')}")
                continue
        data.append({
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
        if is_valid_4d(row.get("4D_Number", "")) == four_d:
            start_date = row.get("Start_Date_DDMMYYYY", "")
            end_date = row.get("End_Date_DDMMYYYY", "")
            
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
# Remove Expired Statuses from Parade_State on App Launch
# ------------------------------------------------------------------------------
def remove_expired_statuses(sheet_parade):
    """
    Removes any row in Parade_State whose End_Date (DDMMYYYY) is strictly before today.
    """
    today = datetime.today().date()
    all_values = sheet_parade.get_all_values()  # includes header row at index 0

    # Iterate from bottom to top (skip header row 0)
    for idx in range(len(all_values) - 1, 0, -1):
        row = all_values[idx]
        if len(row) < 5:
            continue  # skip malformed row

        end_date = row[4].strip()
        end_date = ensure_date_str(end_date)
        try:
            end_dt = datetime.strptime(end_date, "%d%m%Y").date()
            if end_dt < today:
                # Google Sheets rows are 1-based; idx is 0-based in the list
                sheet_parade.delete_rows(idx + 1)
                logger.info(f"Deleted expired status for row {idx + 1}.")
        except ValueError:
            # If there's a parsing error, skip it
            logger.warning(f"Invalid date format in row {idx + 1}: {end_date}")
            continue

# ------------------------------------------------------------------------------
# 4) Streamlit Layout
# ------------------------------------------------------------------------------

st.title("Training & Parade Management App")

# ------------------------------------------------------------------------------
# 5) Company Selection
# ------------------------------------------------------------------------------

st.sidebar.header("Configuration")

# Company selection: Dropdown to select one of the four companies
selected_company = st.sidebar.selectbox(
    "Select Company",
    options=list(COMPANY_SPREADSHEETS.keys())
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

# Remove expired statuses upon loading the selected company's Parade_State
remove_expired_statuses(SHEET_PARADE)  # Corrected from SHEET_PARDE to SHEET_PARADE

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
if "conduct_pointers" not in st.session_state:
    st.session_state.conduct_pointers = ""
if "conduct_submitted_by" not in st.session_state:
    st.session_state.conduct_submitted_by = ""  # New Session State for Conduct

# Parade Session State
if "parade_platoon" not in st.session_state:
    st.session_state.parade_platoon = 1  # Initialize as integer
if "parade_table" not in st.session_state:
    st.session_state.parade_table = []
if "parade_submitted_by" not in st.session_state:
    st.session_state.parade_submitted_by = ""  # New Session State for Parade

# ------------------------------------------------------------------------------
# 7) Feature Selection
# ------------------------------------------------------------------------------
feature = st.sidebar.selectbox(
    "Select Feature",
    ["Add Conduct", "Update Parade", "Queries"]
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
        "Platoon",
        options=[1, 2, 3, 4],
        format_func=lambda x: str(x)
    )
    st.session_state.conduct_name = st.text_input(
        "Conduct Name (e.g. IPPT)",
        value=st.session_state.conduct_name
    )

    # Extra text area for "Pointers" or remarks
    st.session_state.conduct_pointers = st.text_area(
        "Pointers (any additional remarks)",
        value=st.session_state.conduct_pointers
    )

    # New: Submitted By field
    st.session_state.conduct_submitted_by = st.text_input(
        "Submitted By",
        value=st.session_state.conduct_submitted_by
    )

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

        # Fetch cached records
        records_nominal = get_nominal_records(SHEET_NOMINAL)
        records_parade = get_parade_records(SHEET_PARADE)

        # Build conduct table with all personnel, marking 'Is_Outlier' based on status
        conduct_data = build_conduct_table(platoon, date_obj, records_nominal, records_parade)

        # Store in session
        st.session_state.conduct_table = conduct_data
        st.success(f"Loaded {len(conduct_data)} personnel for Platoon {platoon} ({date_str}).")
        logger.info(f"Loaded conduct personnel for Platoon {platoon} on {date_str}.")

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
        submitted_by = st.session_state.conduct_submitted_by.strip()  # Get Submitted By

        if not date_str or not platoon or not cname:
            st.error("Please fill all fields (Date, Platoon, Conduct Name) first.")
            st.stop()

        # Validate date
        try:
            datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format.")
            st.stop()

        # Fetch cached records
        records_nominal = get_nominal_records(SHEET_NOMINAL)
        records_parade = get_parade_records(SHEET_PARADE)

        # We'll figure out who is outlier + who is new to Nominal_Roll
        existing_4ds = {row.get("4D_Number", "").strip().upper() for row in records_nominal}

        new_people = []
        all_outliers = []

        for row in edited_data:
            four_d = is_valid_4d(row.get("4D_Number", ""))
            name_ = ensure_str(row.get("Name", ""))
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
                    new_people.append((name_, four_d, platoon))
                    logger.info(f"Adding new person: {name_}, {four_d}, Platoon {platoon}.")

                # If is_outlier, we'll add to outliers list with StatusDesc
                if is_outlier:
                    if status_desc:
                        all_outliers.append(f"{four_d} ({status_desc})")
                    else:
                        all_outliers.append(f"{four_d}")

        # Insert new people into Nominal_Roll
        for (nm, fd, p_) in new_people:
            formatted_fd = ensure_date_str(fd)
            SHEET_NOMINAL.append_row([nm, formatted_fd, p_, 14, ""])  # Initialize leaves
            logger.info(f"Added new person to Nominal_Roll: {nm}, {formatted_fd}, Platoon {p_}.")

        # Now recalc total strength
        records_nominal = get_nominal_records(SHEET_NOMINAL)  # Refresh after adding new people
        total_strength = get_company_strength(platoon, records_nominal)
        outliers_num = len(all_outliers)
        participating = total_strength - outliers_num
        outliers_str = ",".join(all_outliers)

        # Append row to Conducts with Submitted By
        pointers = st.session_state.conduct_pointers.strip()
        formatted_date_str = ensure_date_str(date_str)
        SHEET_CONDUCTS.append_row([
            formatted_date_str,
            platoon,
            cname,
            total_strength,
            participating,
            outliers_str,
            pointers,
            submitted_by  # Added Submitted By
        ])
        logger.info(f"Appended Conduct: {formatted_date_str}, Platoon {platoon}, {cname}, Total: {total_strength}, Participating: {participating}, Outliers: {outliers_str}, Submitted_By: {submitted_by}")

        st.success(
            f"Conduct Finalized!\n\n"
            f"Date: {formatted_date_str}\n"
            f"Platoon: {platoon}\n"
            f"Conduct Name: {cname}\n"
            f"Total Strength: {total_strength}\n"
            f"Participating: {participating}\n"
            f"Outliers: {outliers_str if outliers_str else 'None'}\n"
            f"Pointers: {pointers if pointers else 'None'}\n"
            f"Submitted By: {submitted_by if submitted_by else 'N/A'}"
        )

        # Clear session state variables
        st.session_state.conduct_date = ""
        st.session_state.conduct_platoon = 1
        st.session_state.conduct_name = ""
        st.session_state.conduct_table = []
        st.session_state.conduct_pointers = ""
        st.session_state.conduct_submitted_by = ""  # Clear Submitted By

        # **Clear Cached Data to Reflect Updates**
        get_nominal_records.clear()
        get_conduct_records.clear()
        get_parade_records.clear()

# ------------------------------------------------------------------------------
# 9) Feature B: Update Parade
# ------------------------------------------------------------------------------
elif feature == "Update Parade":
    st.header("Update Parade State")

    # (a) Input for platoon
    st.session_state.parade_platoon = st.selectbox(
        "Platoon for Parade Update:",
        options=[1, 2, 3, 4],
        format_func=lambda x: str(x)
    )

    # New: Submitted By field
    st.session_state.parade_submitted_by = st.text_input(
        "Submitted By",
        value=st.session_state.parade_submitted_by
    )

    if st.button("Load Personnel"):
        platoon = str(st.session_state.parade_platoon).strip()
        if not platoon:
            st.error("Please select a valid platoon.")
            st.stop()

        # Fetch cached records
        records_nominal = get_nominal_records(SHEET_NOMINAL)
        records_parade = get_parade_records(SHEET_PARADE)

        data = get_company_personnel(platoon, records_nominal, records_parade)
        st.session_state.parade_table = data
        st.info(f"Loaded {len(data)} personnel for Platoon {platoon}.")
        logger.info(f"Loaded personnel for Platoon {platoon}.")

        # ------------------------------------------------------------------------------
        # Display Current Parade Statuses for the Platoon
        # ------------------------------------------------------------------------------
        current_statuses = [
            row for row in records_parade
            if normalize_name(row.get('Platoon', '')) == normalize_name(platoon)
        ]

        if current_statuses:
            st.subheader("Current Parade Status")
            # Format the current statuses for better readability
            formatted_statuses = []
            for status in current_statuses:
                formatted_statuses.append({
                    "4D_Number": status.get("4D_Number", ""),
                    "Name": find_name_by_4d(status.get("4D_Number", ""), records_nominal),
                    "Status": status.get("Status", ""),
                    "Start_Date": status.get("Start_Date_DDMMYYYY", ""),
                    "End_Date": status.get("End_Date_DDMMYYYY", "")
                })
            #st.table(formatted_statuses)
            logger.info(f"Displayed current parade statuses for platoon {platoon}.")

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

        if st.button("Update Parade State"):
            rows_updated = 0
            platoon = str(st.session_state.parade_platoon).strip()
            submitted_by = st.session_state.parade_submitted_by.strip()  # Get Submitted By

            # Fetch cached records
            records_nominal = get_nominal_records(SHEET_NOMINAL)
            records_parade = get_parade_records(SHEET_PARADE)

            for idx, row in enumerate(edited_data):
                four_d = is_valid_4d(row.get('4D_Number', ''))
                status_val = ensure_str(row.get('Status', '')).strip()
                start_val = ensure_str(row.get('Start_Date', '')).strip()
                end_val = ensure_str(row.get('End_Date', '')).strip()

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
                        logger.info(f"Deleted Parade_State row {row_num} for {four_d}.")
                        rows_updated += 1
                        continue
                    except Exception as e:
                        st.error(f"Error deleting row for {four_d}: {e}. Skipping.")
                        logger.error(f"Exception while deleting row for {four_d}: {e}.")
                        continue

                if not status_val or not start_val or not end_val:
                    #st.error(f"Missing fields for {four_d}. Skipping.")
                    logger.error(f"Missing fields for {four_d}.")
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
                        logger.error(f"End date before start date for {four_d}.")
                        continue
                except ValueError:
                    st.error(f"Invalid date(s) for {four_d}, skipping.")
                    logger.error(f"Invalid date format for {four_d}: Start={formatted_start_val}, End={formatted_end_val}.")
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
                        logger.error(f"Invalid leave duration for {four_d}: {dates_str}.")
                        continue

                    # Check for overlapping statuses
                    if has_overlapping_status(four_d, start_dt, end_dt, records_parade):
                        #st.error(f"Leave dates overlap with existing status for {four_d}, skipping.")
                        logger.error(f"Leave dates overlap for {four_d}: {dates_str}.")
                        continue

                    # Fetch current leaves and dates taken
                    try:
                        nominal_record = SHEET_NOMINAL.find(four_d, in_column=2)  # Assuming 4D_Number is column B
                        if nominal_record:
                            current_leaves_left = SHEET_NOMINAL.cell(nominal_record.row, 4).value
                            try:
                                current_leaves_left = int(current_leaves_left)
                            except ValueError:
                                current_leaves_left = 14  # Default if invalid
                                logger.warning(f"Invalid 'Number of Leaves Left' for {four_d}. Resetting to 14.")

                            if leaves_used > current_leaves_left:
                                st.error(f"{four_d} does not have enough leaves left. Available: {current_leaves_left}, Requested: {leaves_used}. Skipping.")
                                logger.error(f"{four_d} insufficient leaves. Available: {current_leaves_left}, Requested: {leaves_used}.")
                                continue

                            # Update leaves left
                            new_leaves_left = current_leaves_left - leaves_used
                            SHEET_NOMINAL.update_cell(nominal_record.row, 4, new_leaves_left)
                            logger.info(f"Updated 'Number of Leaves Left' for {four_d}: {new_leaves_left}.")

                            # Update Dates Taken
                            existing_dates = SHEET_NOMINAL.cell(nominal_record.row, 5).value
                            new_dates_entry = dates_str
                            if existing_dates:
                                updated_dates = existing_dates + f",{new_dates_entry}"
                            else:
                                updated_dates = new_dates_entry
                            SHEET_NOMINAL.update_cell(nominal_record.row, 5, updated_dates)
                            logger.info(f"Updated 'Dates Taken' for {four_d}: {updated_dates}.")
                        else:
                            st.error(f"{four_d} not found in Nominal_Roll. Skipping.")
                            logger.error(f"{four_d} not found in Nominal_Roll.")
                            continue
                    except Exception as e:
                        st.error(f"Error updating leaves for {four_d}: {e}. Skipping.")
                        logger.error(f"Exception while updating leaves for {four_d}: {e}.")
                        continue

                # Update the existing Parade_State row instead of appending
                if row_num:
                    # Find the column numbers based on header
                    header = SHEET_PARADE.row_values(1)
                    try:
                        status_col = header.index("Status") + 1
                        start_date_col = header.index("Start_Date_DDMMYYYY") + 1
                        end_date_col = header.index("End_Date_DDMMYYYY") + 1
                        submitted_by_col = header.index("Submitted_By") + 1 if "Submitted_By" in header else None
                    except ValueError as ve:
                        st.error(f"Required column missing in Parade_State: {ve}.")
                        logger.error(f"Required column missing in Parade_State: {ve}.")
                        continue

                    SHEET_PARADE.update_cell(row_num, status_col, status_val)  # Corrected SHEET_PARDE to SHEET_PARADE
                    SHEET_PARADE.update_cell(row_num, start_date_col, formatted_start_val)  # Corrected SHEET_PARDE to SHEET_PARADE
                    SHEET_PARADE.update_cell(row_num, end_date_col, formatted_end_val)  # Corrected SHEET_PARDE to SHEET_PARADE

                    # Update 'Submitted_By' only if the row was changed
                    if is_changed and submitted_by_col:
                        SHEET_PARADE.update_cell(row_num, submitted_by_col, submitted_by)
                        logger.info(f"Updated Parade_State for {four_d}: Status={status_val}, Start={formatted_start_val}, End={formatted_end_val}, Submitted_By={submitted_by}")
                    elif is_changed and not submitted_by_col:
                        # If 'Submitted_By' column doesn't exist, log a warning
                        logger.warning(f"'Submitted_By' column not found in Parade_State. Cannot update for {four_d}.")
                    
                    rows_updated += 1
                else:
                    # If no existing row, append as a new entry
                    SHEET_PARADE.append_row([platoon, four_d, status_val, formatted_start_val, formatted_end_val, submitted_by])  # Corrected SHEET_PARDE to SHEET_PARADE
                    logger.info(f"Appended Parade_State for {four_d}: Status={status_val}, Start={formatted_start_val}, End={formatted_end_val}, Submitted_By={submitted_by}")
                    rows_updated += 1

            st.success(f"Parade State updated.")
            logger.info(f"Parade State updated for {rows_updated} row(s) for platoon {platoon}.")

            # **Reset session_state variables**
            st.session_state.parade_platoon = 1
            st.session_state.parade_table = []
            st.session_state.parade_submitted_by = ""  # Clear Submitted By

            # **Clear Cached Data to Reflect Updates**
            get_parade_records.clear()
            get_nominal_records.clear()
            get_conduct_records.clear()

# ------------------------------------------------------------------------------
# 10) Feature C: Queries (Combined Query Person & Query Outliers)
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

            # Fetch cached records
            records_nominal = get_nominal_records(SHEET_NOMINAL)
            records_parade = get_parade_records(SHEET_PARADE)

            parade_data = records_parade
            # Filter rows for this 4D
            person_rows = [
                row for row in parade_data
                if row.get("4D_Number", "").strip().upper() == four_d_input_clean
            ]

            if not person_rows:
                st.warning(f"No Parade_State records found for {four_d_input_clean}")
                logger.info(f"No Parade_State records found for {four_d_input_clean}.")
            else:
                # Sort by start date
                def parse_ddmmyyyy(d):
                    try:
                        return datetime.strptime(str(d), "%d%m%Y")
                    except ValueError:
                        return datetime.min

                person_rows.sort(key=lambda r: parse_ddmmyyyy(r.get("Start_Date_DDMMYYYY", "")))

                # Enhance display by adding Name
                enhanced_rows = []
                for row in person_rows:
                    enhanced_rows.append({
                        "4D_Number": row.get("4D_Number", ""),
                        "Name": find_name_by_4d(row.get("4D_Number", ""), records_nominal),
                        "Status": row.get("Status", ""),
                        "Start_Date": row.get("Start_Date_DDMMYYYY", ""),
                        "End_Date": row.get("End_Date_DDMMYYYY", "")
                    })

                st.subheader(f"Statuses for {four_d_input_clean}")
                # Show as a table
                st.table(enhanced_rows)
                logger.info(f"Displayed statuses for {four_d_input_clean}.")

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

            conducts_data = get_conduct_records(SHEET_CONDUCTS)

            # Filter records matching both platoon and conduct name
            matched_records = [
                row for row in conducts_data
                if normalize_name(row.get('Platoon', '')) == normalize_name(platoon_query) and
                   normalize_name(row.get('Conduct_Name', '')) == conduct_norm
            ]

            if not matched_records:
                # Attempt fuzzy matching if no exact match found
                conduct_pairs = [
                    (normalize_name(row.get('Platoon', '')), normalize_name(row.get('Conduct_Name', '')))
                    for row in conducts_data
                    if row.get('Platoon', '').strip() and row.get('Conduct_Name', '').strip()
                ]
                query_pair = (normalize_name(platoon_query), conduct_norm)
                closest_matches = difflib.get_close_matches(query_pair, conduct_pairs, n=1, cutoff=0.6)
                if not closest_matches:
                    st.error("❌ **No similar platoon and conduct combination found.**\n\nPlease check your input and try again.")
                    logger.error(f"No similar platoon and conduct combination found for: {query_pair}.")
                    st.stop()
                matched_norm = closest_matches[0]
                # Retrieve the original names
                matched_records = [
                    row for row in conducts_data
                    if normalize_name(row.get('Platoon', '')) == matched_norm[0] and
                       normalize_name(row.get('Conduct_Name', '')) == matched_norm[1]
                ]
                if not matched_records:
                    st.error("❌ **No data found for the matched platoon and conduct.**")
                    logger.error(f"No data found for the matched platoon and conduct: {matched_norm}.")
                    st.stop()

            # Collect outliers from matched records
            all_outliers = []
            for row in matched_records:
                outliers_value = row.get('Outliers', '')
                if isinstance(outliers_value, (int, float)):
                    outliers_str = str(outliers_value)
                elif isinstance(outliers_value, str):
                    outliers_str = outliers_value
                else:
                    outliers_str = ''

                if outliers_str.lower() != 'none' and outliers_str.strip():
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
                st.markdown(f"📈 **Outliers for '{conduct_query}' at Platoon {platoon_query}:**")
                st.table(outlier_table)
                logger.info(f"Displayed outliers for '{conduct_query}' at Platoon {platoon_query}.")
            else:
                st.info(f"✅ **No outliers recorded for '{conduct_query}' at Platoon {platoon_query}'.**")
                logger.info(f"No outliers recorded for '{conduct_query}' at Platoon {platoon_query}.")
