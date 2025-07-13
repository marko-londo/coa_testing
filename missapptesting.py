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


# ----------- 1. AUTH -----------
credentials_json = st.secrets["auth_users"]["usernames"]
credentials = json.loads(credentials_json)

authenticator = stauth.Authenticate(
    credentials, 'missed_stops_app', 'some_secret_key', cookie_expiry_days=3
)


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
TEMPLATE_ID = "Do not remove this yet"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
credentials_gs = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
gs_client = gspread.authorize(credentials_gs)

name, authentication_status, username = authenticator.login('main')

if authentication_status is False:
    st.error("Incorrect username or password. Please try again.")
    st.stop()
elif authentication_status is None:
    st.warning("Please enter your username and password.")
    st.stop()

st.success(f"Welcome, {name}!")
authenticator.logout("Logout", "sidebar")

# ----------- VERSION & CHANGELOG -----------
APP_VERSION = "v1.2"
CHANGELOG = """
- **v1.2** (2025-07-12):  
    - Fixed Dropbox temporary token issue.
    - Fixed Image upload issue.
    - Uploaded images are now automatically renamed based on relevant data (row, date, service type, etc.) for easier identification and organization.
"""

st.title("Missed Stops Manager")
st.markdown(f"<div style='color:gray;margin-bottom:8px;'>{APP_VERSION}</div>", unsafe_allow_html=True)

cl_col, doc_col = st.columns([3,1])

with cl_col:
    with st.expander("What's New?", expanded=False):
        st.markdown(CHANGELOG)

# ---- Documentation Link ----

with doc_col:
        
    DOC_LINK = "https://docs.google.com/document/d/1UkKj56Qn-25gMWheC-G2rC6YRJzeGsfxk9k2XNLpeTw"
    
    st.link_button("📄 View Full Docs", DOC_LINK)

COLUMNS = [
    "Date", "Submitted By", "Time Called In", "Zone", "Time Sent to JPM",
    "Address", "Service Type", "Route", "Whole Block", "Placement Exception", "PE Address",
    "City Notes", "Time Dispatched", "Driver Check-in Time", "Collection Status",
    "JPM Notes", "Image", "Times Missed", "Last Missed"
]

DAY_TABS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

def upload_image_to_drive(file, folder_id, credentials):
    import io
    from googleapiclient.http import MediaIoBaseUpload

    drive_service = build("drive", "v3", credentials=credentials)

    # Use original filename or fallback
    filename = getattr(file, "name", "upload.jpg")

    # Upload the file
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

    # Return shareable link
    file_id = uploaded_file.get("id")
    return f"https://drive.google.com/uc?id={file_id}"


def get_next_saturday(today):
    days_until_sun = (6 - today.weekday()) % 7
    return today + datetime.timedelta(days=days_until_sun)

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
    today = datetime.date.today().strftime("%-m.%-d.%Y")  # For 7.12.2025
    filename = f"{row_index}-{service_type}-{today}"

    # Use original file extension if possible
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
    next_sunday = get_next_sunday(today)
    monday_of_week = next_sunday - datetime.timedelta(days=6)
    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    tab_dates = { (monday_of_week + datetime.timedelta(days=i)): weekdays[i] for i in range(7) }
    day_label = tab_dates.get(today, today.strftime('%A'))
    return f"{day_label} {today.month}/{today.day}/{str(today.year)[-2:]}"


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

def load_address_df(service_account_info, address_sheet_url):
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    ws = client.open_by_url(address_sheet_url).get_worksheet(0)
    return ws.get_all_records()

address_df = load_address_df(SERVICE_ACCOUNT_INFO, ADDRESS_LIST_SHEET_URL)

# ----------- MAIN PAGE SELECTOR -----------
main_mode = st.sidebar.radio(
    "Choose your mode:",
    ["Submit a Missed Stop (City Side)", "JPM Operations (Dispatch/Complete)"]
)

today = datetime.date.today()
drive = build('drive', 'v3', credentials=credentials_gs)
sheet_title = get_sheet_title(today)
weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
weekly_ss = gs_client.open_by_key(weekly_id)
today_tab = get_today_tab_name(today)

# ----------- PAGE 1: SUBMISSION -----------
if main_mode == "Submit a Missed Stop (City Side)":
    # ----- Build a {zone: collection_day} mapping -----
    service_type = st.selectbox("Service Type", ["MSW", "SS", "YW"])
    zone_field = f"{service_type} Zone"
    day_field = f"{service_type} Zone"
    # Build {zone: day} from your address_df
    zone_to_day = {}
    for row in address_df:
        zone = row.get(zone_field)
        day = row.get(day_field) or row.get("Day", "")
        if zone:  # Avoid blanks/None
            # Find first day encountered for zone
            if zone not in zone_to_day:
                zone_to_day[zone] = row.get(f"{service_type} Zone") or row.get(f"{service_type} Day", "")

    # Use your standard order (edit as needed)
    week_order = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    
    def get_weekday_idx(zone):
        for i, day in enumerate(week_order):
            if day.lower() in str(zone_to_day[zone]).lower():
                return i
        return 99
    
    zones = sorted({row[zone_field] for row in address_df if row[zone_field]}, key=get_weekday_idx)
    
    def weekday_to_week_order_idx(py_weekday):
        return (py_weekday + 1) % 7
    
    today_py_idx = datetime.date.today().weekday()  # 0=Monday, ..., 6=Sunday
    today_idx = weekday_to_week_order_idx(today_py_idx)  # 0=Sunday, ..., 6=Saturday
    
    yesterday_idx = (today_idx - 1) % 7
    yesterday_day = week_order[yesterday_idx]
    
    # Find the first zone whose assigned day matches yesterday's day
    default_zone = None
    for z in zones:
        if yesterday_day.lower() in str(zone_to_day[z]).lower():
            default_zone = z
            break
    if not default_zone:
        default_zone = zones[0] if zones else ""
    
    # Now build dropdown with correct order and default
    zone = st.selectbox("Zone", zones, index=zones.index(default_zone) if default_zone in zones else 0)
    
    # The rest remains the same!

    address = st.selectbox(
        "Address",
        sorted({row["Address"] for row in address_df if row[zone_field] == zone})
    )
    route = next((row[f"{service_type} Route"] for row in address_df if row["Address"] == address), "")
    whole_block = st.selectbox("Whole Block", ["NO", "YES"])
    placement_exception = st.selectbox("Placement Exception?", ["NO", "YES"])
    pe_address = st.text_input("PE Address") if placement_exception == "YES" else "N/A"
    col1, col2 = st.columns([2, 1])
    now = datetime.datetime.now(pytz.timezone("America/New_York"))
    current_time_str = now.strftime("%I:%M")
    default_ampm = "AM" if now.hour < 12 else "PM"

    with col1:
        called_in_time = st.text_input("Time Called In (HH:MM)", placeholder=current_time_str)
    with col2:
        ampm = st.selectbox("AM/PM", ["AM", "PM"], index=0 if default_ampm == "AM" else 1)

    city_notes = st.text_area("City Notes (optional)")
    submit_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")

    form_data = {
        "Date": str(today), "Submitted By": name, "Time Called In": f"{called_in_time.strip()} {ampm}", "Zone": zone,
        "Time Sent to JPM": submit_time, "Address": address, "Service Type": service_type, "Route": route,
        "Whole Block": whole_block, "Placement Exception": placement_exception, "PE Address": pe_address,
        "City Notes": city_notes, "Collection Status": "Pending"
    }
    
    # (Continue with your existing validation and submit logic as before)

    
    missing_fields = []
    
    # Validate time format (e.g., 9:30 or 09:30)
    time_format_valid = bool(re.match(r"^([1-9]|1[0-2]):[0-5][0-9]$", called_in_time.strip()))
    if not called_in_time.strip():
        missing_fields.append("Time Called In")
    elif not time_format_valid:
        st.warning("⏰ Enter time as HH:MM in 12-hour format (e.g., 9:30 or 10:45)")
        missing_fields.append("Time Called In (invalid format)")
    
    if placement_exception == "YES" and not pe_address.strip():
        missing_fields.append("PE Address")
    
    if missing_fields:
        st.warning(f"🚫 Please complete the following required fields: {', '.join(missing_fields)}")
        st.stop()

    if st.button("Submit Missed Stop"):
        # --- Check for past misses in Master Log ---
        master_id = get_master_log_id(drive, gs_client, FOLDER_ID)
        master_ws = gs_client.open_by_key(master_id).sheet1
        master_records = master_ws.get_all_records()
    
        # --- Prevent duplicate same-day submissions ---
        duplicate_today = any(
            row.get("Address") == address and row.get("Date") == str(today)
            for row in master_records
        )
        if duplicate_today:
            st.error("🚫 This address has already been submitted today.")
            st.stop()
    
        # --- Track past misses ---
        matching_entries = [
            row for row in master_records
            if row.get("Address") == address
        ]
        form_data["Times Missed"] = str(len(matching_entries) + 1)
        form_data["Last Missed"] = matching_entries[-1]["Date"] if matching_entries else "First Time"
    
        # --- Write to sheets ---
        ws = weekly_ss.worksheet(today_tab)
        ws.append_row([form_data.get(col, "") for col in COLUMNS], value_input_option="USER_ENTERED")
        master_ws.append_row([form_data.get(col, "") for col in COLUMNS], value_input_option="USER_ENTERED")
    
        st.success("Miss submitted successfully!")
        st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")




# ========== PAGE 2: JPM OPERATIONS ==========
else:
    st.subheader("JPM Side: Dispatch or Complete Misses")
    jpm_mode = st.radio("JPM Action:", ["Dispatch Misses", "Complete a Missed Stop"])

    today = datetime.date.today()
    sheet_title = get_sheet_title(today)
    drive = build('drive', 'v3', credentials=credentials_gs)
    weekly_id = ensure_gsheet_exists(drive, FOLDER_ID, sheet_title)
    weekly_ss = gs_client.open_by_key(weekly_id)
    today_tab = get_today_tab_name(today)

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

    if jpm_mode == "Dispatch Misses":
        try:
            ws = weekly_ss.worksheet(today_tab)
        except gspread.exceptions.WorksheetNotFound:
            st.info(f"No '{today_tab}' tab found in this week's log.")
            ws = None

        open_misses = []
        if ws:
            records = ws.get_all_records()
            for i, row in enumerate(records):
                if (
                    str(row.get("Collection Status", "")).strip().upper() in ("", "MISSED", "PENDING")
                    and not row.get("Time Dispatched")
                ):
                    label = (
                        f"{row['Address']} | {row['Zone']} | Called: {row.get('Time Called In', '')}"
                    )
                    open_misses.append({"row_idx": i+2, "row": row, "label": label})

        if not open_misses:
            st.info("🎉 No pending missed stops to dispatch today!")
            st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")
        else:
            chosen = st.multiselect(
                "Select missed stops to dispatch:", open_misses, format_func=lambda x: x["label"]
            )
            if chosen and st.button("Dispatch Selected Misses"):
                now_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")
                indices = [c["row_idx"] for c in chosen]
                update_rows(ws, indices, {"Time Dispatched": now_time, "Collection Status": "Dispatched"})

                master_id = get_master_log_id(drive, gs_client, FOLDER_ID)
                master_ws = gs_client.open_by_key(master_id).sheet1
                master_records = master_ws.get_all_records()
                for c in chosen:
                    r = c["row"]
                    match_idx = next(
                        (i+2 for i, mr in enumerate(master_records)
                         if mr.get("Address") == r.get("Address")
                         and mr.get("Date") == r.get("Date")
                         and mr.get("Time Called In") == r.get("Time Called In")),
                        None
                    )
                    if match_idx:
                        update_rows(master_ws, [match_idx], {"Time Dispatched": now_time, "Collection Status": "Dispatched"})
                st.success(f"Dispatched {len(indices)} missed stop(s)!")
                st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")


    else:  # Complete Misses
        try:
            ws = weekly_ss.worksheet(today_tab)
        except gspread.exceptions.WorksheetNotFound:
            st.info(f"No '{today_tab}' tab found in this week's log.")
            ws = None

        if "to_complete_data" not in st.session_state or st.session_state.get("reload_to_complete", False):
            st.session_state.to_complete_data = ws.get_all_records() if ws else []
            st.session_state.reload_to_complete = False

        to_complete = []
        for i, row in enumerate(st.session_state.to_complete_data):
            if row.get("Time Dispatched") and row.get("Collection Status", "").strip().upper() == "DISPATCHED":
                label = f"{row['Address']} | {row['Zone']} | Dispatched: {row.get('Time Dispatched')}"
                to_complete.append({"row_idx": i+2, "row": row, "label": label})

        if not to_complete:
            st.info("✅ No dispatched, incomplete misses for today!")
            st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")
        else:
            chosen = st.selectbox("Select a dispatched miss to complete:", to_complete, format_func=lambda x: x["label"])
            sel = chosen["row"]

            col1, col2 = st.columns([2, 1])
            now = datetime.datetime.now(pytz.timezone("America/New_York"))
            current_time_str = now.strftime("%I:%M")
            default_ampm = "AM" if now.hour < 12 else "PM"
            with col1:
                driver_checkin = st.text_input("Driver Check-in Time (HH:MM)", placeholder=current_time_str)
            with col2:
                ampm2 = st.selectbox("AM/PM (Check-in)", ["AM", "PM"], index=0 if default_ampm == "AM" else 1)
            # Accepts 1 or 2 digit hour, leading zero optional
            valid_ci = bool(re.match(r"^([1-9]|1[0-2]):[0-5][0-9]$", driver_checkin.strip()))
            if not valid_ci and driver_checkin:
                st.warning("⏰ Enter check-in time in 12-hour format, e.g., 1:30 or 09:45")
            
            # Normalize format for consistency (pad hour if needed)
            if valid_ci:
                parts = driver_checkin.strip().split(":")
                hour = parts[0].zfill(2)
                minute = parts[1]
                formatted_checkin = f"{hour}:{minute}"
                check_in_time = f"{formatted_checkin} {ampm2}"
            else:
                check_in_time = ""


            collection_status = st.selectbox("Collection Status", ["Picked Up", "Not Out"])
            jpm_notes = st.text_area("JPM Notes")
            uploaded_image = st.file_uploader("Upload Image (optional)", type=["jpg","jpeg","png","heic","webp"])
            image_link = "N/A"
            
            if uploaded_image:
                uploaded_image.seek(0)
                st.image(uploaded_image, caption="Preview", use_container_width=True)

            can_complete = valid_ci and collection_status
            
            if st.button("Complete Missed Stop", disabled=not can_complete):
                now_time = datetime.datetime.now(pytz.timezone("America/New_York")).strftime("%Y-%m-%d %H:%M:%S")
                check_in_time = f"{driver_checkin} {ampm2}"
            
                # Upload image now (not earlier)
                if uploaded_image:
                    try:
                        uploaded_image.seek(0)
                        row_index = chosen["row_idx"]  # Chosen is your current row dict, as in your code
                        service_type = sel.get("Service Type", "Unknown")  # sel is the row being completed
                        dropbox_url = upload_to_dropbox(uploaded_image, row_index, service_type)
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
            
                # update_rows, etc...

                update_rows(ws, [chosen["row_idx"]], updates)

                master_id = get_master_log_id(drive, gs_client, FOLDER_ID)
                master_ws = gs_client.open_by_key(master_id).sheet1
                master_records = master_ws.get_all_records()
                match_idx = next(
                    (i+2 for i, mr in enumerate(master_records)
                     if mr.get("Address") == sel.get("Address") and
                        mr.get("Date") == sel.get("Date") and
                        mr.get("Time Called In") == sel.get("Time Called In")),
                    None
                )
                if match_idx:
                    update_rows(master_ws, [match_idx], updates)

                st.session_state.reload_to_complete = True
                st.success("Miss completed and logged!")
                st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")
