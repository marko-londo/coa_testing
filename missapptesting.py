import os
import json
import streamlit as st
import streamlit_authenticator as stauth
import gspread
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import datetime
import pytz
import re
import dropbox
from googleapiclient.errors import HttpError
import uuid

jpm_logo = "https://github.com/marko-londo/coa_testing/blob/main/1752457645003.png?raw=true"

coa_logo = "https://raw.githubusercontent.com/marko-londo/coa_testing/0ef57ff891efc1b7258d99368cd47b487c4284a7/Allentown_logo.svg"

credentials_json = st.secrets["auth_users"]["usernames"]

credentials = json.loads(credentials_json)

authenticator = stauth.Authenticate(
    credentials, 'missed_stops_app', 'some_secret_key', cookie_expiry_days=3)

app_key = st.secrets["dropbox"]["app_key"]

app_secret = st.secrets["dropbox"]["app_secret"]

refresh_token = st.secrets["dropbox"]["refresh_token"]

dbx = dropbox.Dropbox(
    oauth2_refresh_token=refresh_token,
    app_key=app_key,
    app_secret=app_secret
)

SERVICE_ACCOUNT_INFO = st.secrets["google_service_account"]

FOLDER_ID = '18f3aW-ZI5-tNKBCfHwToQ7MXQ3DS1MFj'

ADDRESS_LIST_SHEET_URL = "https://docs.google.com/spreadsheets/d/1JJeufDkoQ6p_LMe5F-Nrf_t0r_dHrAHu8P8WXi96V9A/edit#gid=0"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

credentials_gs = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)

gs_client = gspread.authorize(credentials_gs)

st.set_page_config(
    page_title="MPU Portal | JP Mascaro & Sons",
    page_icon="https://raw.githubusercontent.com/marko-londo/coa_testing/refs/heads/main/favicon.ico",
    layout="centered",  # or "wide"
    initial_sidebar_state="collapsed",
    )

st.logo(image=coa_logo)

def get_prior_legit_miss_count(master_records, address, this_row_date, this_row_called_in_time):
    """
    Returns the number of legit missed stops for this address
    *before* this row, based on date and time called in (unique per row).
    """
    prior_rows = []
    for row in master_records:
        # Skip this row itself by comparing date + time called in
        if (
            row.get("Address") == address and
            str(row.get("Date")) < str(this_row_date)
        ):
            if str(row.get("Collection Status", "").strip().upper()) in LEGIT_MISS_STATUSES:
                prior_rows.append(row)
        # If same day, check time
        elif (
            row.get("Address") == address and
            str(row.get("Date")) == str(this_row_date) and
            str(row.get("Time Called In")) < str(this_row_called_in_time)
        ):
            if str(row.get("Collection Status", "").strip().upper()) in LEGIT_MISS_STATUSES:
                prior_rows.append(row)
    return len(prior_rows)


def user_login(authenticator, credentials):
    name, authentication_status, username = authenticator.login('main')

    if authentication_status is False:
        st.error("Incorrect username or password. Please try again.")
        st.stop()
    elif authentication_status is None:
        st.info("Please enter your username and password.")
        st.stop()

    user_obj = credentials["usernames"].get(username, {})
    user_role = user_obj.get("role", "city")
    st.info(f"Welcome, {name}!")
    authenticator.logout("Logout", "sidebar")
    return name, username, user_role

def generate_all_minutes():
    times = []
    for hour in range(0, 24):
        for minute in range(0, 60):
            t = datetime.time(hour, minute)
            times.append(t.strftime("%I:%M %p"))
    return times

def updates():
    APP_VERSION = "v2.3"
    CHANGELOG = """
    - **v2.3** (2025-07-18):  
        - Added “Submit Completion Times” section for JPM
        - Stops submitted before completion time will be flagged as potentially premature


    """

    
    st.markdown("<br>", unsafe_allow_html=True)  # One blank line
    
    # --- Centered Logo ---
    st.markdown(
        f"""
        <div style='display: flex; justify-content: center; align-items: center; margin-bottom: 12px;'>
            <img src='{jpm_logo}' width='320'>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # --- H1 Style ---
    st.markdown("""
        <style>
        h1 {
            font-family: 'Poppins', sans-serif !important;
            font-weight: 700 !important;
            font-size: 3em !important;
            letter-spacing: 1.5px !important;
            text-shadow:
                -1px -1px 0 #181b20,
                 1px -1px 0 #181b20,
                -1px  1px 0 #181b20,
                 1px  1px 0 #181b20,
                 0  3px 12px #6CA0DC55;
        }
        </style>
        """, unsafe_allow_html=True)
    
    # --- Centered Header, Subtitle, and Divider ---
    st.markdown(
        """
        <div style='text-align:center;'>
            <h1 style='color:#6CA0DC; margin-bottom:0;'>Missed Pickup Portal</h1>
            <div style='font-size:1.1em; font-style:italic; margin-bottom:12px;'>
                <span style='color:#FF8C8C;'>City of Allentown</span>
                <span style='color:#fff; padding:0 10px;'>|</span>
                <span style='color:#FF8C8C;'>JP Mascaro & Sons</span>
            </div>
            <hr style='border:1px solid #ececec; margin-top:0;'>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # --- App Version (left-aligned) ---
    st.markdown(f"<div style='color:gray;margin-bottom:8px;'>{APP_VERSION}</div>", unsafe_allow_html=True)

    with st.expander("What's New?", expanded=False):
            st.markdown(CHANGELOG)
        
    doc_col, sht_col, fold_col = st.columns(3)
    
    with doc_col:
            
        DOC_LINK = "https://docs.google.com/document/d/1UkKj56Qn-25gMWheC-G2rC6YRJzeGsfxk9k2XNLpeTw"
        st.link_button("📄 View Full Docs", DOC_LINK)

    with sht_col:
        st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")

    with fold_col:
        st.link_button("Open Folder", f"https://drive.google.com/drive/u/0/folders/1ogx3zPeIdTKp7C5EJ5jKavFv21mDmySj")

COLUMNS = [
    "Date",
    "Submitted By",
    "Time Called In",
    "Zone",
    "YW Zone Color",
    "Time Sent to JPM",
    "Address",
    "Service Type",
    "Route",
    "Whole Block",
    "Placement Exception",
    "PE Address",
    "City Notes",
    "Time Dispatched",
    "Driver Check-in Time",
    "Collection Status",
    "JPM Notes",
    "Image",
    "Times Missed",
    "Last Missed",
    "MissID"
]

DAY_TABS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

LEGIT_MISS_STATUSES = ("PENDING", "DISPATCHED", "NOT OUT", "PICKED UP")

def calculate_times_missed(master_records, address):
    return sum(
        1
        for row in master_records
        if (
            row.get("Address") == address and
            str(row.get("Collection Status", "")).strip().upper() in LEGIT_MISS_STATUSES
        )
    )

def upload_image_to_drive(file, folder_id, credentials):
    import io
    from googleapiclient.http import MediaIoBaseUpload

    drive_service = build("drive", "v3", credentials=credentials)

    filename = getattr(file, "name", "upload.jpg")

    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }
    media = MediaIoBaseUpload(io.BytesIO(file.read()), mimetype=file.type)
    uploaded_file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    file_id = uploaded_file.get("id")
    return f"https://drive.google.com/uc?id={file_id}"

def get_completion_times_sheet_title(today):
    next_saturday = get_next_saturday(today)
    return f"completion Times Week Ending {next_saturday.strftime('%Y-%m-%d')}"

def ensure_completion_times_gsheet_exists(drive, folder_id, title):
    results = drive.files().list(
        q=f"'{folder_id}' in parents and name='{title}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    else:
        # If you want to create it automatically, implement creation logic here.
        st.error(
            f"Completion Times sheet '{title}' does not exist in the specified folder.\n"
            "Please contact your admin to create this week's completion log sheet."
        )
        st.stop()

def submit_completion_time_section():
    st.subheader("Submit Completion Time")

    today = datetime.datetime.now(pytz.timezone("America/New_York")).date()
    if today.weekday() == 6:
        st.info("Completion times cannot be submitted on Sundays. Please return on a service day (Monday–Saturday).")
        return

    completion_sheet_title = get_completion_times_sheet_title(today)
    drive = build('drive', 'v3', credentials=credentials_gs)
    completion_sheet_id = ensure_completion_times_gsheet_exists(drive, FOLDER_ID, completion_sheet_title)
    completion_times_ws = gs_client.open_by_key(completion_sheet_id).worksheet(get_today_tab_name(today))


    # Fetch all rows; assume 1 header + 3 rows (MSW, SS, YW)
    sheet_data = completion_times_ws.get_all_records()
    user = name  # from user_login()

    # Find incomplete services
    incomplete_services = [
        (idx, row) for idx, row in enumerate(sheet_data, start=2)
        if row.get("Service Type") and row.get("Completion Status", "").strip().upper() != "COMPLETE"
    ]

    if not incomplete_services:
        st.info("All services completed for today.")
    else:
        for row_idx, row in incomplete_services:
            service_type = row.get("Service Type")
            st.write(f"**{service_type}** not yet completed.")
            time_key = f"completion_time_{service_type}"
            if time_key not in st.session_state:
                st.session_state[time_key] = now_str if now_str in time_options else time_options[0]
            selected_time = st.selectbox(
                f"Select completion time for {service_type}",
                time_options,
                key=time_key
            )
                    
            if st.button(f"Submit {service_type}", key=f"submit_{service_type}"):
                now_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")
                completion_times_ws.update(
                    f"B{row_idx}:E{row_idx}",
                    [["COMPLETE", st.session_state[time_key], now_time, user]]
                )
                st.info(f"Completion time for {service_type} recorded at {st.session_state[time_key]} by {user}.")
                del st.session_state[time_key]  # Clear it after submission
                st.rerun()


    # --- DIALOG DEFINITION ---
    @st.dialog("WARNING: This will clear all existing submissions in the sheet for today. Continue?")
    def clear_all_dialog():
        if st.button("Yes, Clear All"):
            for i in range(2, 5):  # Rows 2,3,4 (Google Sheets 1-indexed)
                completion_times_ws.update(f"B{i}:E{i}", [["NOT COMPLETE", "", "", ""]])
            st.info("All submissions cleared.")
            st.rerun()
        if st.button("Cancel"):
            st.rerun()

    # --- TRIGGER THE DIALOG ---
    if st.button("Clear All Submissions", type="primary"):
        clear_all_dialog()



def get_next_saturday(today):
    # If today is Sunday, treat as start of next week, so return *next* Saturday
    if today.weekday() == 6:  # Sunday
        # Sunday: add 6 days to get to next Saturday
        return today + datetime.timedelta(days=6)
    else:
        # For Mon-Sat: get this week's Saturday
        days_until_sat = 5 - today.weekday()
        return today + datetime.timedelta(days=days_until_sat)

def upload_to_dropbox(file, row_index, service_type):
    import dropbox
    app_key = st.secrets["dropbox"]["app_key"]
    app_secret = st.secrets["dropbox"]["app_secret"]
    refresh_token = st.secrets["dropbox"]["refresh_token"]
    dbx = dropbox.Dropbox(
        oauth2_refresh_token=refresh_token,
        app_key=app_key,
        app_secret=app_secret
    )
    filename = f"{row_index}-{service_type}-{today_str}"
    
    ext = ""
    if hasattr(file, "name") and "." in file.name:
        ext = file.name[file.name.rfind("."):]
    elif hasattr(file, "type"):
        mime_map = {"image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp", "image/heic": ".heic"}
        ext = mime_map.get(getattr(file, "type", ""), "")

    filename += ext

    dropbox_path = f"/missed_stops/{filename}"
    file.seek(0)
    dbx.files_upload(file.read(), dropbox_path, mode=dropbox.files.WriteMode.overwrite)

    try:
        link_metadata = dbx.sharing_create_shared_link_with_settings(dropbox_path)
        url = link_metadata.url
    except dropbox.exceptions.ApiError as e:
        if (isinstance(e.error, dropbox.sharing.CreateSharedLinkWithSettingsError) and
            e.error.is_shared_link_already_exists()):
            links = dbx.sharing_list_shared_links(path=dropbox_path, direct_only=True).links
            if links:
                url = links[0].url
            else:
                raise RuntimeError("Could not get existing Dropbox shared link.")
        else:
            raise
    return url.replace("?dl=0", "?raw=1")

def get_sheet_title(today):
    next_saturday = get_next_saturday(today)
    return f"Misses Week Ending {next_saturday.strftime('%Y-%m-%d')}"

def get_monday_of_week(saturday_date):
    return saturday_date - datetime.timedelta(days=5)

def get_today_tab_name(today):
    # If Sunday, tab is *next* Monday of next week (for the next sheet)
    if today.weekday() == 6:  # Sunday
        next_monday = today + datetime.timedelta(days=1)
        tab_date = next_monday
        day_label = "Monday"
    else:
        # As before
        next_saturday = get_next_saturday(today)
        monday_of_week = next_saturday - datetime.timedelta(days=5)
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        tab_date = monday_of_week + datetime.timedelta(days=today.weekday())
        day_label = weekdays[today.weekday()]
    return f"{day_label} {tab_date.month}/{tab_date.day}/{str(tab_date.year)[-2:]}"


def ensure_gsheet_exists(drive, folder_id, title):
    results = drive.files().list(
        q=f"'{folder_id}' in parents and name='{title}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    else:
        st.error(
            f"Sheet '{title}' does not exist in the specified folder.\n"
            "Please contact your admin to create this week's log sheet."
        )
        st.stop()

def find_row_by_missid(ws, missid):
    col_idx = len(COLUMNS)  # last column
    missids = ws.col_values(col_idx)
    for i, v in enumerate(missids):
        if v == missid:
            return i + 1  # Google Sheets rows are 1-based
    return None
    
def get_master_log_id(drive, folder_id):
    results = drive.files().list(
        q=f"'{folder_id}' in parents and name = 'Master Misses Log' and mimeType = 'application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    else:
        st.error(
            "The 'Master Misses Log' sheet does not exist in the specified folder.\n"
            "Please contact your admin to create the log sheet."
        )
        st.stop()

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def update_rows(ws, indices, updates, columns=COLUMNS):
    last_col = colnum_string(len(columns))
    for idx in indices:
        row_values = ws.row_values(idx)
        row_dict = dict(zip(columns, row_values + [""]*(len(columns)-len(row_values))))
        row_dict.update(updates)
        ws.update(
            f"A{idx}:{last_col}{idx}",
            [[row_dict.get(col, "") for col in columns]],
            value_input_option="USER_ENTERED"
        )

@st.cache_data(ttl=3600)
def load_address_df(_service_account_info, address_sheet_url):
    creds = Credentials.from_service_account_info(_service_account_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    ws = client.open_by_url(address_sheet_url).get_worksheet(0)
    return ws.get_all_records()
address_df = load_address_df(SERVICE_ACCOUNT_INFO, ADDRESS_LIST_SHEET_URL)

def help_page(name, user_role):
    st.subheader("Help & Support")
    st.write(
        "Welcome to the Missed Pickup Portal Help page. "
        "For detailed documentation, click the “View Full Docs” button above. "
        "If you would like to submit feedback, request additional features, or report a bug, "
        "please use the 'Submit Feedback' button below. "
        "If you are in need of immediate assistance, please contact us via email or phone. "
        "Thank you for using our service!"
    )

    st.markdown("---")

    st.write("#### Rate your overall experience:")
    feedback = st.feedback("thumbs", key="overall_exp")

    FEEDBACK_SHEET_ID = "1fUrJymiIfC5GS_ofz9x4czUG6e3b8W63mMwLUyxHvFM"
    FEEDBACK_SHEET_NAME = "Feedback"

    if feedback is not None:
        # Map thumbs to text
        rating_map = {0: "Thumbs Down", 1: "Thumbs Up"}
        rating_text = rating_map.get(feedback, str(feedback))

        row = [
            name,
            user_role,
            datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S"),
            "Quick Feedback",
            "",  # No details for thumbs, just the rating
            rating_text
        ]
        try:
            feedback_ws = gs_client.open_by_key(FEEDBACK_SHEET_ID).worksheet(FEEDBACK_SHEET_NAME)
            feedback_ws.append_row(row)
            if feedback == 1:
                st.info("Thanks for the thumbs up!")
            else:
                st.info("Sorry to hear that. For more detailed feedback or to report an issue, please use the button below.")
        except Exception as e:
            st.error(f"Failed to write to feedback sheet: {e}")

    st.markdown("---")

    @st.dialog("Submit Feedback / Report Bug / Request Feature")
    def feedback_dialog():
        feedback_type = st.selectbox("Type", ["Bug Report", "Feature Request", "General Feedback"])
        details = st.text_area("Describe the issue or idea")
        if st.button("Submit"):
            row = [
                name,
                user_role,
                datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S"),
                feedback_type,
                details,
                ""  # Leave rating blank for detailed feedback
            ]
            try:
                feedback_ws = gs_client.open_by_key(FEEDBACK_SHEET_ID).worksheet(FEEDBACK_SHEET_NAME)
                feedback_ws.append_row(row)
                st.info("Thank you for your feedback! It has been recorded.")
            except Exception as e:
                st.error(f"Failed to write to feedback sheet: {e}")
            st.rerun()

    if st.button("Submit Feedback / Report Bug / Request Feature"):
        feedback_dialog()




def city_ops(name, user_role):
    st.sidebar.subheader("City of Allentown")
    if "city_mode" not in st.session_state:
        st.session_state.city_mode = "Submit a Missed Pickup"

    city_mode = st.sidebar.radio("Select Action:", ["Submit a Missed Pickup", "Help"])

    if city_mode == "Submit a Missed Pickup":
        today = datetime.datetime.now(pytz.timezone("America/New_York")).date()
    
        drive = build('drive', 'v3', credentials=credentials_gs)
        sheet_title = get_sheet_title(today)
        weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
        weekly_ss = gs_client.open_by_key(weekly_id)
        today_tab = get_today_tab_name(today)
        
        service_type = st.selectbox("Service Type", ["MSW", "SS", "YW"])
        zone_field = f"{service_type} Zone"
        day_field = f"{service_type} Zone"
        zone_to_day = {}
        for row in address_df:
            zone = row.get(zone_field)
            day = row.get(day_field) or row.get("Day", "")
            if zone:  
                if zone not in zone_to_day:
                    zone_to_day[zone] = row.get(f"{service_type} Zone") or row.get(f"{service_type} Day", "")
    
    
        week_order = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        
        def get_weekday_idx(zone):
            for i, day in enumerate(week_order):
                if day.lower() in str(zone_to_day[zone]).lower():
                    return i
            return 99
        
        zones = sorted({row[zone_field] for row in address_df if row[zone_field]}, key=get_weekday_idx)
        
        def weekday_to_week_order_idx(py_weekday):
            return (py_weekday + 1) % 7
        
        today_py_idx = datetime.date.today().weekday()  
        today_idx = weekday_to_week_order_idx(today_py_idx)  
        
        yesterday_idx = (today_idx - 1) % 7
        yesterday_day = week_order[yesterday_idx]
        
    
        default_zone = None
        for z in zones:
            if yesterday_day.lower() in str(zone_to_day[z]).lower():
                default_zone = z
                break
        if not default_zone:
            default_zone = zones[0] if zones else ""
        
    
        zone = st.selectbox("Zone", zones, index=zones.index(default_zone) if default_zone in zones else 0)
    
        zone_color = None
        if service_type == "YW":
            # Only addresses in this zone
            zone_addresses = [row for row in address_df if row[zone_field] == zone]
            # Pull unique YW Zone Colors
            zone_colors = sorted({row["YW Zone Color"] for row in zone_addresses if row.get("YW Zone Color")})
            if zone_colors:
                zone_color = st.selectbox("YW Zone Color", zone_colors)
            else:
                zone_color = ""
        
        if service_type == "YW":
            address = st.selectbox(
                "Address",
                sorted({
                    row["Address"]
                    for row in address_df
                    if row[zone_field] == zone and row.get("YW Zone Color") == zone_color
                })
            )
        else:
            address = st.selectbox(
                "Address",
                sorted({
                    row["Address"]
                    for row in address_df
                    if row[zone_field] == zone
                })
            )
        selected_row = next((row for row in address_df if row["Address"] == address), None)
        if selected_row and "Latitude" in selected_row and "Longitude" in selected_row:
            import pandas as pd
            map_df = pd.DataFrame([{
                "lat": float(selected_row["Latitude"]),
                "lon": float(selected_row["Longitude"])
            }])
            st.map(map_df, latitude="lat", longitude="lon", zoom=16, size=10)       
        if service_type == "MSW":
            route = next(
                (
                    row["MSW Route"]
                    for row in address_df
                    if row["Address"] == address and row["MSW Zone"] == zone
                ),
                ""
            )
        elif service_type == "SS":
            route = next(
                (
                    row["SS Route"]
                    for row in address_df
                    if row["Address"] == address and row["SS Zone"] == zone
                ),
                ""
            )
        elif service_type == "YW":
            route = next(
                (
                    row["YW Route"]
                    for row in address_df
                    if row["Address"] == address
                    and row["YW Zone"] == zone
                    and row.get("YW Zone Color", "") == zone_color
                ),
                ""
            )
        else:
            route = ""

        placement_exception = st.selectbox("Placement Exception?", ["NO", "YES"])
        pe_address = st.text_input("PE Address") if placement_exception == "YES" else "N/A"
        fields_to_reset = [
            "whole_block", "called_in_time", "city_notes", 
            "placement_exception", "pe_address"
        ]
        
        # --- Whole Block ---
        whole_block = st.selectbox("Whole Block", ["NO", "YES"], key="whole_block")
        
        # --- Time Called In ---
        if "called_in_time" not in st.session_state:
            now = datetime.datetime.now(pytz.timezone("America/New_York"))
            current_time_str = now.strftime("%I:%M %p")
            st.session_state.called_in_time = (
                current_time_str if current_time_str in time_options else time_options[0]
            )
        called_in_time = st.selectbox(
            "Time Called In",
            time_options,
            index=time_options.index(st.session_state.called_in_time),
            key="called_in_time"
        )
        
        city_notes = st.text_input("City Notes (optional)", key="city_notes")
        submit_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")
        form_data = {
            "Date": str(today), "Submitted By": name, "Time Called In": called_in_time, "Zone": zone,
            "Time Sent to JPM": submit_time, "Address": address, "Service Type": service_type, "Route": route,
            "Whole Block": whole_block, "Placement Exception": placement_exception, "PE Address": pe_address,
            "City Notes": city_notes, "Collection Status": "Pending", "YW Zone Color": zone_color if service_type == "YW" else "N/A", "MissID": str(uuid.uuid4())
        }

        completion_sheet_title = get_completion_times_sheet_title(today)
        try:
            completion_sheet_id = ensure_completion_times_gsheet_exists(drive, FOLDER_ID, completion_sheet_title)
            completion_times_ws = gs_client.open_by_key(completion_sheet_id).worksheet(get_today_tab_name(today))
            # Find the correct row for this service type ("MSW", "SS", or "YW")
            ct_records = completion_times_ws.get_all_records()
            completion_row = next((row for row in ct_records if row.get("Service Type", "").strip().upper() == service_type), None)
            if completion_row and completion_row.get("Completion Status", "").strip().upper() == "NOT COMPLETE":
                form_data["Collection Status"] = "Premature"
                st.info(f"FYI: The {service_type} service has not been marked completed yet for today. "
                        f"This stop will be flagged as **Premature**. ")
        except Exception as e:
            st.error(f"Could not check completion status for today: {e}")        
        
        missing_fields = []
        
        if placement_exception == "YES" and not pe_address.strip():
            missing_fields.append("PE Address")
        
        if missing_fields:
            st.error(f"🚫 Please complete the following required fields: {', '.join(missing_fields)}")
            st.stop()
        
        if st.button("Submit Missed Stop"):
    
            master_id = get_master_log_id(drive, FOLDER_ID)
            master_ws = gs_client.open_by_key(master_id).sheet1
            master_records = master_ws.get_all_records()
    
            duplicate_pending_or_premature = any(
                row.get("Address") == address and
                str(row.get("Collection Status", "")).strip().upper() in ("PENDING", "PREMATURE")
                and not row.get("Time Dispatched")
                for row in master_records
            )
            if duplicate_pending_or_premature:
                st.error("🚫 This address already has a pending or premature missed stop not yet dispatched. Please close or dispatch it before submitting a new one.")
                st.stop()

            # Only count legitimate missed stops (exclude Premature/Rejected/other non-miss statuses)
            matching_entries = [
                row for row in master_records
                if (
                    row.get("Address") == address and
                    str(row.get("Collection Status", "")).strip().upper() in LEGIT_MISS_STATUSES
                )
            ]
            form_data["Times Missed"] = str(len(matching_entries) + 1)
            form_data["Last Missed"] = matching_entries[-1]["Date"] if matching_entries else "First Time"

    
            ws = weekly_ss.worksheet(today_tab)
            ws.append_row([form_data.get(col, "") for col in COLUMNS], value_input_option="USER_ENTERED")
            master_ws.append_row([form_data.get(col, "") for col in COLUMNS], value_input_option="USER_ENTERED")
        
            st.info("Miss submitted successfully!")         
            for k in fields_to_reset:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()  # Ensures UI is reset instantly
        
        # Manual "Start Over" button for user control
        if st.button("Start Over"):
            for k in fields_to_reset:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()
    else:
        help_page(name, user_role)

def jpm_ops(name, user_role):

    st.sidebar.subheader("JPM Operations")
    jpm_mode = st.sidebar.radio("Select Action:", ["Dispatch Misses", "Complete a Missed Stop", "Submit Completion Times", "Help"])

    def update_rows(ws, indices, updates, columns=COLUMNS):
        last_col = colnum_string(len(columns))
        for idx in indices:
            try:
                row_values = ws.row_values(idx)
                row_dict = dict(zip(columns, row_values + [""]*(len(columns)-len(row_values))))
                row_dict.update(updates)
                ws.update(
                    f"A{idx}:{last_col}{idx}",
                    [[row_dict.get(col, "") for col in columns]],
                    value_input_option="USER_ENTERED"
                )
            except HttpError as e:
                if e.resp.status == 429 or "Rate Limit" in str(e):
                    st.error(
                        "⚠️ Too many updates at once! Google Sheets is rate-limiting you. "
                        "Please wait a minute and try again, or select fewer items at a time."
                    )
                    # Optionally: break or return to prevent further updates
                    break
                else:
                    st.error(f"Error updating row {idx}: {e}")

    if jpm_mode == "Dispatch Misses":
        # Always work from Master Misses Log
        master_id = get_master_log_id(drive, FOLDER_ID)
        master_ws = gs_client.open_by_key(master_id).sheet1
        master_records = master_ws.get_all_records()
    
        open_misses = []
        for i, row in enumerate(master_records):
            if (
                str(row.get("Collection Status", "")).strip().upper() in ("", "PREMATURE", "PENDING")
                and not row.get("Time Dispatched")
            ):
                label = (
                    f"{row.get('Address','')} | {row.get('Zone','')} | Date: {row.get('Date','')} | Called: {row.get('Time Called In','')}"
                )
                open_misses.append({"row_idx": i+2, "row": row, "label": label})
    
        if not open_misses:
            st.info("🎉 No pending missed stops to dispatch!")

        else:
            chosen = st.multiselect(
                "Select missed stops to dispatch:", open_misses, format_func=lambda x: x["label"]
            )
            if chosen and st.button("Dispatch Selected Misses"):
                now_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")
                
                # --- Master Log: update by MissID ---
                for c in chosen:
                    missid = c["row"].get("MissID")
                    row_idx_master = find_row_by_missid(master_ws, missid)
                    current_status = c["row"].get("Collection Status", "").strip().upper()
                    updates = {"Time Dispatched": now_time}
                    if current_status != "PREMATURE":
                        updates["Collection Status"] = "Dispatched"
                    # Else, leave status as "Premature"
                    if row_idx_master:
                        update_rows(master_ws, [row_idx_master], updates)
                    else:
                        st.error(f"Could not find MissID {missid} in Master Misses Log. It may have been deleted.")

                
                # --- Weekly log: update by MissID ---
                for c in chosen:
                    r = c["row"]
                    missid = r.get("MissID")
                    miss_date = r.get("Date")
                    if miss_date:
                        try:
                            miss_date_dt = datetime.datetime.strptime(miss_date, "%Y-%m-%d").date()
                            sheet_title = get_sheet_title(miss_date_dt)
                            weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
                            weekly_ss = gs_client.open_by_key(weekly_id)
                            tab_name = get_today_tab_name(miss_date_dt)
                            ws = weekly_ss.worksheet(tab_name)
                            row_idx_weekly = find_row_by_missid(ws, missid)
                            current_status = r.get("Collection Status", "").strip().upper()
                            updates = {"Time Dispatched": now_time}
                            if current_status != "PREMATURE":
                                updates["Collection Status"] = "Dispatched"
                            # Else, leave status as "Premature"
                            if row_idx_weekly:
                                update_rows(ws, [row_idx_weekly], updates)
                            else:
                                st.error(f"Could not find MissID {missid} in weekly sheet '{tab_name}'.")
                        except Exception as e:
                            pass  # If the weekly sheet/tab doesn't exist, just skip

            
                st.info(f"Dispatched {len(chosen)} missed stop(s)!")
                if chosen:
                    last_dispatched = chosen[-1]
                    miss_date = last_dispatched["row"].get("Date")
                    if miss_date:
                        miss_date_dt = datetime.datetime.strptime(miss_date, "%Y-%m-%d").date()
                        dispatched_weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, get_sheet_title(miss_date_dt))
                    else:
                        st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")


    elif jpm_mode == "Complete a Missed Stop":
        fields_to_reset = ["driver_checkin", "collection_status", "jpm_notes", "uploaded_image"]
        master_id = get_master_log_id(drive, FOLDER_ID)
        master_ws = gs_client.open_by_key(master_id).sheet1
        master_records = master_ws.get_all_records()
    
        # Use session state for caching/filtering if desired (optional)
        if "to_complete_data" not in st.session_state or st.session_state.get("reload_to_complete", False):
            st.session_state.to_complete_data = master_records
            st.session_state.reload_to_complete = False
    
        to_complete = []
        for i, row in enumerate(st.session_state.to_complete_data):
            status = row.get("Collection Status", "").strip().upper()
            has_dispatched = bool(row.get("Time Dispatched"))

            if has_dispatched and status in ("DISPATCHED", "DELAYED", "PREMATURE"):
                label = (
                    f"{row.get('Address','')} | {row.get('Zone','')} | Date: {row.get('Date','')} | Dispatched: {row.get('Time Dispatched','')}"
                )
                to_complete.append({"row_idx": i+2, "row": row, "label": label})

    
        if not to_complete:
            st.info("✅ No dispatched, incomplete misses for today!")
        else:
            st.caption(
                "Only 'Premature' stops that have been dispatched will be listed here for completion."
            )

            chosen = st.selectbox("Select a dispatched miss to complete:", to_complete, format_func=lambda x: x["label"])
            sel = chosen["row"]
    
            if "driver_checkin" not in st.session_state:
                now = datetime.datetime.now(pytz.timezone("America/New_York"))
                current_time_str = now.strftime("%I:%M %p")
                st.session_state.driver_checkin = (
                    current_time_str if current_time_str in time_options else time_options[0]
                )
            driver_checkin = st.selectbox(
                "Driver Check In Time",
                time_options,
                index=time_options.index(st.session_state.driver_checkin),
                key="driver_checkin"
            )
            
            # --- The rest, using session state for sticky fields if you want ---
            collection_status = st.selectbox("Collection Status", ["Picked Up", "Not Out", "Rejected", "Delayed", "Confirmed Premature", "One Time Exception"], key="collection_status")
            jpm_notes = st.text_area("JPM Notes", key="jpm_notes")
            uploaded_image = st.file_uploader("Upload Image (optional)", type=["jpg","jpeg","png","heic","webp"])
            
            image_link = "N/A"
            
            if uploaded_image:
                uploaded_image.seek(0)
                st.image(uploaded_image, caption="Preview", use_container_width=True)
            
            can_complete = driver_checkin and collection_status
            
            if st.button("Complete Missed Stop", disabled=not can_complete):
                now_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")
                check_in_time = driver_checkin
            
                if uploaded_image:
                    try:
                        uploaded_image.seek(0)
                        # Find matching row index in the weekly tab
                        row_index_weekly = None
                        r = sel
                        miss_date = r.get("Date")
                        if miss_date:
                            try:
                                miss_date_dt = datetime.datetime.strptime(miss_date, "%Y-%m-%d").date()
                                sheet_title = get_sheet_title(miss_date_dt)
                                weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
                                weekly_ss = gs_client.open_by_key(weekly_id)
                                tab_name = get_today_tab_name(miss_date_dt)
                                ws = weekly_ss.worksheet(tab_name)
                                tab_records = ws.get_all_records()
                                for j, tr in enumerate(tab_records):
                                    if (tr.get("Address") == r.get("Address")
                                        and tr.get("Date") == r.get("Date")
                                        and tr.get("Time Called In") == r.get("Time Called In")):
                                        row_index_weekly = j + 2  # +2 because get_all_records skips header
                                        break
                            except Exception as e:
                                pass  # If not found, will fallback to master row
                
                        if not row_index_weekly:
                            row_index_weekly = chosen["row_idx"]  # fallback to master row
                
                        service_type = sel.get("Service Type", "Unknown")
                        dropbox_url = upload_to_dropbox(uploaded_image, row_index_weekly, service_type)
                        image_link = f'=HYPERLINK("{dropbox_url}", "Image Link")'
                    except Exception as e:
                        st.error(f"Dropbox upload failed: {e}")
                        image_link = "UPLOAD FAILED"
                else:
                    image_link = "N/A"
    
                updates = {
                    "Driver Check-in Time": check_in_time,
                    "Collection Status": collection_status,
                    "JPM Notes": jpm_notes,
                    "Image": image_link,
                }
                
                address = sel.get("Address")
                row_date = sel.get("Date")
                called_in_time = sel.get("Time Called In")
                prior_legit_misses = get_prior_legit_miss_count(master_records, address, row_date, called_in_time)
                
                if collection_status.upper() in ("PREMATURE", "CONFIRMED PREMATURE", "REJECTED", "ONE TIME EXCEPTION"):
                    updates["Times Missed"] = str(prior_legit_misses)
                    # Find last legit prior miss date, else "Never"
                    prior_misses = [
                        row for row in master_records
                        if (
                            row.get("Address") == address and
                            str(row.get("Collection Status", "")).strip().upper() in LEGIT_MISS_STATUSES and
                            (str(row.get("Date")) < str(row_date) or (
                                str(row.get("Date")) == str(row_date) and str(row.get("Time Called In")) < str(called_in_time)
                            ))
                        )
                    ]
                    if prior_misses:
                        updates["Last Missed"] = prior_misses[-1]["Date"]
                    else:
                        updates["Last Missed"] = "Never"

                else:
                    updates["Times Missed"] = str(prior_legit_misses + 1)
                    updates["Last Missed"] = row_date

                missid = sel.get("MissID") if isinstance(sel, dict) else chosen["row"].get("MissID")
                
                # --- Update in Master Misses Log ---
                row_idx_master = find_row_by_missid(master_ws, missid)
                if row_idx_master:
                    update_rows(master_ws, [row_idx_master], updates)
                else:
                    st.error("Could not find this record in the Master Misses Log. It may have been deleted.")
                
                # --- Also update in the correct weekly sheet/tab for recordkeeping ---
                r = sel
                miss_date = r.get("Date")
                if miss_date:
                    try:
                        miss_date_dt = datetime.datetime.strptime(miss_date, "%Y-%m-%d").date()
                        sheet_title = get_sheet_title(miss_date_dt)
                        weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
                        weekly_ss = gs_client.open_by_key(weekly_id)
                        tab_name = get_today_tab_name(miss_date_dt)
                        ws = weekly_ss.worksheet(tab_name)
                        # NEW: find row by MissID in this worksheet!
                        row_idx_weekly = find_row_by_missid(ws, missid)
                        if row_idx_weekly:
                            update_rows(ws, [row_idx_weekly], updates)
                        else:
                            # Optional: log/warn
                            st.error("Could not find this record in the weekly sheet. It may have been deleted or is missing a MissID.")
                    except Exception as e:
                        pass  # skip if the weekly sheet/tab doesn't exist

    
                st.session_state.reload_to_complete = True
                st.info("Miss completed and logged!")
                if 'miss_date_dt' in locals():
                    completed_weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, get_sheet_title(miss_date_dt))
                for k in fields_to_reset:
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()  # Immediately resets the UI
            
            if st.button("Start Over"):
                for k in fields_to_reset:
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()                

    elif jpm_mode == "Submit Completion Times":

        submit_completion_time_section()        

    else:
        help_page(name, user_role)

NY_TZ = pytz.timezone("America/New_York")
now = datetime.datetime.now(NY_TZ)
now_str = now.strftime("%I:%M %p")
time_options = generate_all_minutes()
today = datetime.datetime.now(pytz.timezone("America/New_York")).date()
today_str = today.strftime("%-m.%-d.%Y")
drive = build('drive', 'v3', credentials=credentials_gs)
sheet_title = get_sheet_title(today)
weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
weekly_ss = gs_client.open_by_key(weekly_id)
today_tab = get_today_tab_name(today)
name, username, user_role = user_login(authenticator, credentials)

updates()
if user_role == "city":
    city_ops(name, user_role)
elif user_role == "jpm":
    jpm_ops(name, user_role)
else:
    st.error("Role not recognized. Please contact your admin.")

