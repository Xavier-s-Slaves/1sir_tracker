import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
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
SPREADSHEET_NAME = "tracker"  # <-- Change if your spreadsheet has a different title

@st.cache_resource
def get_sheets():
    """
    Open the spreadsheet once and return references to worksheets.
    This is cached so we don't re-open on each script run.
    """
    gc = gspread.authorize(creds)
    sh = gc.open(SPREADSHEET_NAME)
    return {
        "nominal": sh.worksheet("Nominal_Roll"),
        "parade": sh.worksheet("Parade_State"),
        "conducts": sh.worksheet("Conducts")
    }

worksheets = get_sheets()
SHEET_NOMINAL = worksheets["nominal"]
SHEET_PARADE = worksheets["parade"]
SHEET_CONDUCTS = worksheets["conducts"]

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
def get_nominal_records():
    """
    Returns all rows from Nominal_Roll as a list of dicts, cached for 5 min.
    """
    records = SHEET_NOMINAL.get_all_records()
    # Ensure all relevant fields are strings and properly formatted
    for row in records:
        row['4D_Number'] = is_valid_4d(row.get('4D_Number', ''))
        row['Company'] = ensure_str(row.get('Company', ''))
        row['Name'] = ensure_str(row.get('Name', ''))
    # Remove any records with invalid 4D_Number
    records = [row for row in records if row['4D_Number']]
    return records

@st.cache_data(ttl=300)
def get_parade_records():
    """
    Returns all rows from Parade_State as a list of dicts, cached.
    """
    records = SHEET_PARADE.get_all_records()
    # Ensure all relevant fields are strings and properly formatted
    for row in records:
        row['4D_Number'] = is_valid_4d(row.get('4D_Number', ''))
        row['Company'] = ensure_str(row.get('Company', ''))
        row['Start_Date_DDMMYYYY'] = ensure_date_str(row.get('Start_Date_DDMMYYYY', ''))
        row['End_Date_DDMMYYYY'] = ensure_date_str(row.get('End_Date_DDMMYYYY', ''))
        row['Status'] = ensure_str(row.get('Status', ''))
    # Remove any records with invalid 4D_Number
    records = [row for row in records if row['4D_Number']]
    return records

@st.cache_data(ttl=300)
def get_conduct_records():
    """
    Returns all rows from Conducts as a list of dicts, cached.
    """
    records = SHEET_CONDUCTS.get_all_records()
    # Ensure all relevant fields are strings and properly formatted
    for row in records:
        row['Date'] = ensure_date_str(row.get('Date', ''))
        row['Company'] = ensure_str(row.get('Company', ''))
        row['Conduct_Name'] = ensure_str(row.get('Conduct_Name', ''))
        row['Outliers'] = ensure_str(row.get('Outliers', ''))
    return records

def get_company_strength(company: str) -> int:
    """
    Count how many rows in Nominal_Roll belong to that company.
    Uses cached data to avoid repeated API calls.
    """
    records = get_nominal_records()
    return sum(
        1 for row in records
        if normalize_name(row.get('Company', '')) == normalize_name(company)
    )

def get_company_personnel(company: str):
    """
    Returns a list of dicts for 'Update Parade' (or similar)
    with placeholders for Status, Start_Date, End_Date.
    Only from Nominal_Roll, filtered by the given company.
    """
    records = get_nominal_records()
    data = []
    for row in records:
        c = row.get('Company', '')
        if normalize_name(c) == normalize_name(company):
            data.append({
                'Name': row.get('Name', 'Unknown'),
                '4D_Number': row.get('4D_Number', 'Unknown'),
                'Status': '',         # to be filled
                'Start_Date': '',     # to be filled
                'End_Date': '',       # to be filled
                'Number_of_Leaves_Left': row.get('Number of Leaves Left', 14),
                'Dates_Taken': row.get('Dates Taken', '')
            })
    return data

def find_name_by_4d(four_d: str) -> str:
    """
    Optional helper: If you want to look up person's Name from Nominal_Roll
    given a 4D_Number.
    """
    four_d = ensure_str(four_d).upper()
    for row in get_nominal_records():
        if ensure_str(row.get("4D_Number", "")).upper() == four_d:
            return ensure_str(row.get("Name", ""))
    return ""

def build_onstatus_table(company: str, date_obj: datetime):
    """
    Return a list of dicts for everyone on status for that date + company.
    If multiple statuses exist for the same person, prioritize based on a hierarchy.
    For example: 'Leave' > 'Fever' > 'MC'
    """
    parade_data = get_parade_records()
    status_priority = {'leave': 3, 'fever': 2, 'mc': 1}  # Define priority
    out = {}
    for row in parade_data:
        c = row.get('Company', '')
        if normalize_name(c) == normalize_name(company):
            start_str = row.get('Start_Date_DDMMYYYY', "")
            end_str = row.get('End_Date_DDMMYYYY', "")
            
            # Ensure dates are strings with leading zeros
            start_str = ensure_date_str(start_str)
            end_str = ensure_date_str(end_str)
            
            try:
                start_d = datetime.strptime(start_str, "%d%m%Y")
                end_d = datetime.strptime(end_str, "%d%m%Y")
            except ValueError:
                logger.warning(f"Invalid date format in Parade_State: {start_str} - {end_str}")
                continue

            if start_d <= date_obj <= end_d:
                four_d = is_valid_4d(row.get('4D_Number', ''))
                status = ensure_str(row.get('Status', '')).lower()
                if not four_d:
                    continue  # Skip invalid 4D_Number
                # If multiple statuses, keep the one with higher priority
                if four_d in out:
                    existing_status = out[four_d]['StatusDesc'].lower()
                    if status_priority.get(status, 0) > status_priority.get(existing_status, 0):
                        out[four_d] = {
                            "Name": find_name_by_4d(four_d),
                            "4D_Number": four_d,
                            "StatusDesc": ensure_str(row.get('Status', '')),
                            "Is_Outlier": True
                        }
                else:
                    out[four_d] = {
                        "Name": find_name_by_4d(four_d),
                        "4D_Number": four_d,
                        "StatusDesc": ensure_str(row.get('Status', '')),
                        "Is_Outlier": True
                    }
    logger.info(f"Built on-status table with {len(out)} entries for company {company} on {date_obj.strftime('%d%m%Y')}.")
    return list(out.values())

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

def has_overlapping_status(four_d: str, new_start: datetime, new_end: datetime) -> bool:
    """
    Check if the new status dates overlap with existing statuses for the given 4D_Number.
    """
    parade_data = get_parade_records()
    four_d = is_valid_4d(four_d)
    if not four_d:
        return False  # Invalid 4D_Number, cannot have overlapping status
    
    for row in parade_data:
        if is_valid_4d(row.get("4D_Number", "")) == four_d:
            start_date = ensure_date_str(row.get("Start_Date_DDMMYYYY", ""))
            end_date = ensure_date_str(row.get("End_Date_DDMMYYYY", ""))
            
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
def remove_expired_statuses():
    """
    Removes any row in Parade_State whose End_Date (DDMMYYYY) is strictly before 'now'.
    """
    now = datetime.now()
    all_values = SHEET_PARADE.get_all_values()  # includes header row at index 0

    # Iterate from bottom to top (skip header row 0)
    for idx in range(len(all_values) - 1, 0, -1):
        row = all_values[idx]
        if len(row) < 5:
            continue  # skip malformed row

        end_date = row[4].strip()
        end_date = ensure_date_str(end_date)
        try:
            end_dt = datetime.strptime(end_date, "%d%m%Y")
            if end_dt < now:
                # Google Sheets rows are 1-based; idx is 0-based in the list
                SHEET_PARADE.delete_rows(idx + 1)
                logger.info(f"Deleted expired status for row {idx + 1}.")
        except ValueError:
            # If there's a parsing error, skip it
            logger.warning(f"Invalid date format in row {idx + 1}: {end_date}")
            continue

# Call the removal function once, right after we load the sheet references
remove_expired_statuses()

# ------------------------------------------------------------------------------
# 4) Streamlit Layout
# ------------------------------------------------------------------------------
st.title("Training & Parade Management App")

feature = st.sidebar.selectbox(
    "Select Feature",
    ["Add Conduct", "Update Parade", "Queries"]
)

# ------------------------------------------------------------------------------
# 5) Session State: We store data so it's not lost on each run
# ------------------------------------------------------------------------------
if "conduct_date" not in st.session_state:
    st.session_state.conduct_date = ""
if "conduct_company" not in st.session_state:
    st.session_state.conduct_company = ""
if "conduct_name" not in st.session_state:
    st.session_state.conduct_name = ""
if "conduct_table" not in st.session_state:
    st.session_state.conduct_table = []
if "conduct_pointers" not in st.session_state:
    st.session_state.conduct_pointers = ""
if "conduct_submitted_by" not in st.session_state:
    st.session_state.conduct_submitted_by = ""  # New Session State for Conduct

if "parade_company" not in st.session_state:
    st.session_state.parade_company = ""
if "parade_table" not in st.session_state:
    st.session_state.parade_table = []
if "parade_submitted_by" not in st.session_state:
    st.session_state.parade_submitted_by = ""  # New Session State for Parade

# ------------------------------------------------------------------------------
# 6) Feature A: Add Conduct (table-based On-Status approach)
# ------------------------------------------------------------------------------
if feature == "Add Conduct":
    st.header("Add Conduct - Table-Based On-Status")

    # (a) Basic Inputs
    st.session_state.conduct_date = st.text_input(
        "Date (DDMMYYYY)",
        value=st.session_state.conduct_date
    )
    st.session_state.conduct_company = st.text_input(
        "Company",
        value=st.session_state.conduct_company
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
        comp = st.session_state.conduct_company.strip()

        if not date_str or not comp:
            st.error("Please enter both Date and Company.")
            st.stop()

        # Validate date format
        try:
            date_obj = datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format (use DDMMYYYY).")
            st.stop()

        # Build table from Parade_State for that date + company
        onstatus_data = build_onstatus_table(comp, date_obj)

        # Store in session
        st.session_state.conduct_table = onstatus_data
        st.success(f"Loaded {len(onstatus_data)} on-status personnel for {comp} ({date_str}).")
        logger.info(f"Loaded on-status personnel for {comp} on {date_str}.")

    # (c) Data Editor (allow new rows) - ALWAYS show, so you can finalize even with zero outliers
    st.write("Toggle 'Is_Outlier' if not participating, or add new rows for extra people.")
    edited_data = st.data_editor(
        st.session_state.conduct_table,
        num_rows="dynamic",
        use_container_width=True
    )

    # (d) Finalize Conduct
    if st.button("Finalize Conduct"):
        date_str = st.session_state.conduct_date.strip()
        comp = st.session_state.conduct_company.strip()
        cname = st.session_state.conduct_name.strip()
        submitted_by = st.session_state.conduct_submitted_by.strip()  # Get Submitted By

        if not date_str or not comp or not cname:
            st.error("Please fill all fields (Date, Company, Conduct Name) first.")
            st.stop()

        # Validate date
        try:
            datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            st.error("Invalid date format.")
            st.stop()

        # We'll figure out who is outlier + who is new to Nominal_Roll
        existing_nominal = get_nominal_records()
        existing_4ds = {row.get("4D_Number", "").strip().upper() for row in existing_nominal}

        new_people = []
        all_outliers = []

        for row in edited_data:
            four_d = is_valid_4d(row.get("4D_Number", ""))
            name_ = ensure_str(row.get("Name", ""))
            is_outlier = row.get("Is_Outlier", False)
            status_desc = ensure_str(row.get("StatusDesc", ""))

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
                    new_people.append((name_, four_d, comp))
                    logger.info(f"Adding new person: {name_}, {four_d}, {comp}.")

                # If is_outlier, we'll add to outliers list
                if is_outlier:
                    if status_desc:
                        all_outliers.append(f"{four_d} ({status_desc})")
                    else:
                        all_outliers.append(f"{four_d}")

        # Insert new people into Nominal_Roll
        for (nm, fd, c_) in new_people:
            formatted_fd = ensure_date_str(fd)
            SHEET_NOMINAL.append_row([nm, formatted_fd, c_, 14, ""])  # Initialize leaves
            logger.info(f"Added new person to Nominal_Roll: {nm}, {formatted_fd}, {c_}.")

        # Now recalc total strength
        total_strength = get_company_strength(comp)
        outliers_num = len(all_outliers)
        participating = total_strength - outliers_num
        outliers_str = ",".join(all_outliers)

        # Append row to Conducts with Submitted By
        pointers = st.session_state.conduct_pointers.strip()
        formatted_date_str = ensure_date_str(date_str)
        SHEET_CONDUCTS.append_row([
            formatted_date_str,
            comp,
            cname,
            total_strength,
            participating,
            outliers_str,
            pointers,
            submitted_by  # Added Submitted By
        ])
        logger.info(f"Appended Conduct: {formatted_date_str}, {comp}, {cname}, Total: {total_strength}, Participating: {participating}, Outliers: {outliers_str}, Submitted_By: {submitted_by}")

        st.success(
            f"Conduct Finalized!\n\n"
            f"Date: {formatted_date_str}\n"
            f"Company: {comp}\n"
            f"Conduct Name: {cname}\n"
            f"Total Strength: {total_strength}\n"
            f"Participating: {participating}\n"
            f"Outliers: {outliers_str if outliers_str else 'None'}\n"
            f"Pointers: {pointers if pointers else 'None'}\n"
            f"Submitted By: {submitted_by if submitted_by else 'N/A'}"
        )

        # Clear session state variables
        st.session_state.conduct_date = ""
        st.session_state.conduct_company = ""
        st.session_state.conduct_name = ""
        st.session_state.conduct_table = []
        st.session_state.conduct_pointers = ""
        st.session_state.conduct_submitted_by = ""  # Clear Submitted By

        # **Clear Cached Data to Reflect Updates**
        get_nominal_records.clear()
        get_conduct_records.clear()
        get_parade_records.clear()

# ------------------------------------------------------------------------------
# 7) Feature B: Update Parade
# ------------------------------------------------------------------------------
elif feature == "Update Parade":
    st.header("Update Parade State")

    # (a) Input for company
    st.session_state.parade_company = st.text_input(
        "Company for Parade Update:",
        value=st.session_state.parade_company
    )

    # New: Submitted By field
    st.session_state.parade_submitted_by = st.text_input(
        "Submitted By",
        value=st.session_state.parade_submitted_by
    )

    if st.button("Load Personnel"):
        c = st.session_state.parade_company.strip()
        if not c:
            st.error("Please enter a valid company.")
            st.stop()

        data = get_company_personnel(c)
        st.session_state.parade_table = data
        st.info(f"Loaded {len(data)} personnel for {c}.")
        logger.info(f"Loaded {len(data)} personnel for {c}.")

    # (b) Show data editor if we have data
    if st.session_state.parade_table:
        st.subheader("Edit Parade Data, Then Click 'Update'")
        st.write("Fill in 'Status', 'Start_Date (DDMMYYYY)', 'End_Date (DDMMYYYY)'")

        edited_data = st.data_editor(
            st.session_state.parade_table,
            num_rows="dynamic",
            use_container_width=True
        )

        if st.button("Update Parade State"):
            rows_updated = 0
            c = st.session_state.parade_company.strip()
            submitted_by = st.session_state.parade_submitted_by.strip()  # Get Submitted By

            for row in edited_data:
                four_d = is_valid_4d(row.get('4D_Number', ''))
                status_val = ensure_str(row.get('Status', '')).strip()
                start_val = ensure_str(row.get('Start_Date', '')).strip()
                end_val = ensure_str(row.get('End_Date', '')).strip()

                # Validate 4D_Number format
                if not four_d:
                    st.error(f"Invalid 4D_Number format: {row.get('4D_Number', '')}. Skipping.")
                    logger.error(f"Invalid 4D_Number format: {row.get('4D_Number', '')}.")
                    continue

                if not status_val or not start_val or not end_val:
                    st.error(f"Missing fields for {four_d}. Skipping.")
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
                    if has_overlapping_status(four_d, start_dt, end_dt):
                        st.error(f"Leave dates overlap with existing status for {four_d}, skipping.")
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

                # Append to Parade_State with Submitted By
                SHEET_PARADE.append_row([c, four_d, status_val, formatted_start_val, formatted_end_val, submitted_by])
                rows_updated += 1
                logger.info(f"Appended Parade_State for {four_d}: Status={status_val}, Start={formatted_start_val}, End={formatted_end_val}, Submitted_By={submitted_by}")

            st.success(f"Parade State updated for {rows_updated} row(s).")
            logger.info(f"Parade State updated for {rows_updated} row(s) for company {c}.")

            # **Reset session_state variables**
            st.session_state.parade_company = ""
            st.session_state.parade_table = []
            st.session_state.parade_submitted_by = ""  # Clear Submitted By

            # **Clear Cached Data to Reflect Updates**
            get_parade_records.clear()
            get_nominal_records.clear()
            get_conduct_records.clear()

# ------------------------------------------------------------------------------
# 8) Feature C: Queries (Combined Query Person & Query Outliers)
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

            parade_data = get_parade_records()
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

                st.subheader(f"Statuses for {four_d_input_clean}")
                # Show as a table
                st.table(person_rows)
                logger.info(f"Displayed statuses for {four_d_input_clean}.")

    # ---------------------------
    # Tab 2: Query Outliers
    # ---------------------------
    with query_tabs[1]:
        st.subheader("Query Outliers for a Specific Company & Conduct")

        # Input fields
        comp_q = st.text_input("Company", key="query_outliers_company")
        cond_q = st.text_input("Conduct Name", key="query_outliers_conduct")

        if st.button("Get Outliers", key="btn_query_outliers"):
            company_query = ensure_str(comp_q)
            conduct_query = ensure_str(cond_q)

            if not company_query or not conduct_query:
                st.error("Please enter both Company and Conduct Name.")
                st.stop()

            company_norm = normalize_name(company_query)
            conduct_norm = normalize_name(conduct_query)

            conducts_data = get_conduct_records()

            # Filter records matching both company and conduct name
            matched_records = [
                row for row in conducts_data
                if normalize_name(row.get('Company', '')) == company_norm and
                   normalize_name(row.get('Conduct_Name', '')) == conduct_norm
            ]

            if not matched_records:
                # Attempt fuzzy matching if no exact match found
                company_conduct_pairs = [
                    (normalize_name(row.get('Company', '')), normalize_name(row.get('Conduct_Name', '')))
                    for row in conducts_data
                    if row.get('Company', '').strip() and row.get('Conduct_Name', '').strip()
                ]
                query_pair = (company_norm, conduct_norm)
                closest_matches = difflib.get_close_matches(query_pair, company_conduct_pairs, n=1, cutoff=0.6)
                if not closest_matches:
                    st.error("❌ **No similar company and conduct combination found.**\n\nPlease check your input and try again.")
                    logger.error(f"No similar company and conduct combination found for: {query_pair}.")
                    st.stop()
                matched_norm = closest_matches[0]
                # Retrieve the original names
                matched_records = [
                    row for row in conducts_data
                    if normalize_name(row.get('Company', '')) == matched_norm[0] and
                       normalize_name(row.get('Conduct_Name', '')) == matched_norm[1]
                ]
                if not matched_records:
                    st.error("❌ **No data found for the matched company and conduct.**")
                    logger.error(f"No data found for the matched company and conduct: {matched_norm}.")
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
                st.markdown(f"📈 **Outliers for '{conduct_query}' at '{company_query}':**")
                st.table(outlier_table)
                logger.info(f"Displayed outliers for '{conduct_query}' at '{company_query}'.")
            else:
                st.info(f"✅ **No outliers recorded for '{conduct_query}' at '{company_query}'.**")
                logger.info(f"No outliers recorded for '{conduct_query}' at '{company_query}'.")
