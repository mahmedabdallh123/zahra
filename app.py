import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import requests
import shutil
import re
from datetime import datetime, timedelta
import io
import uuid


from PIL import Image
from github import Github, GithubException

# ------------------------------- App Config -------------------------------
APP_CONFIG = {
    "APP_TITLE": "vivo 1 - CMMS",
    "APP_ICON": "🏭",
    "REPO_NAME": "mahmedabdallh123/zahra",
    "BRANCH": "main",
    "FILE_PATH": "l9.xlsx",
    "LOCAL_FILE": "l9.xlsx",
    "MAX_ACTIVE_USERS": 5,
    "SESSION_DURATION_MINUTES": 60,
    "IMAGES_FOLDER": "event_images",
    "ALLOWED_IMAGE_TYPES": ["jpg", "jpeg", "png", "gif", "bmp", "webp"],
    "MAX_IMAGE_SIZE_MB": 10,
    "DEFAULT_SHEET_COLUMNS": [
        "Repair_Duration", "Date", "Equipment", "Event_Fault", "Corrective_Action",
        "Performed_By", "Spare_Parts_Used", "Fault_Type", "Technician_Skill",
        "Safety_Compliance", "Image_URL"
    ],
    "SPARE_PARTS_SHEET": "Spare_Parts",
    "SPARE_PARTS_COLUMNS": ["Part_Name", "Size", "Tension", "Available_Quantity", "Lead_Time", "Is_Critical", "Section", "Image_URL"],
    "MAINTENANCE_SHEET": "Preventive_Maintenance",
    "MAINTENANCE_COLUMNS": [
        "Equipment", "Maintenance_Type", "Task_Name", "Period_Days",
        "Last_Execution", "Next_Date", "Notes", "Default_Spare_Part", "Image_URL"
    ],
    "GENERAL_SECTION": "General"
}

st.set_page_config(page_title=APP_CONFIG["APP_TITLE"], layout="wide")

# ------------------------------- Optional Imports -------------------------------
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False
    try:
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
        plt.rcParams['font.family'] = 'Arial'
        MATPLOTLIB_AVAILABLE = True
    except ImportError:
        MATPLOTLIB_AVAILABLE = False

# ------------------------------- Constants -------------------------------
USERS_FILE = "users.json"
STATE_FILE = "state.json"
SESSION_DURATION = timedelta(minutes=APP_CONFIG["SESSION_DURATION_MINUTES"])
MAX_ACTIVE_USERS = APP_CONFIG["MAX_ACTIVE_USERS"]
IMAGES_FOLDER = APP_CONFIG["IMAGES_FOLDER"]
EQUIPMENT_CONFIG_FILE = "equipment_config.json"
SUPPORT_CONFIG_FILE = "support_config.json"

GITHUB_EXCEL_URL = f"https://github.com/{APP_CONFIG['REPO_NAME'].split('/')[0]}/{APP_CONFIG['REPO_NAME'].split('/')[1]}/raw/{APP_CONFIG['BRANCH']}/{APP_CONFIG['FILE_PATH']}"
GITHUB_USERS_URL = "https://raw.githubusercontent.com/mahmedabdallh123/zahra/refs/heads/main/users.json"
GITHUB_REPO_USERS = "mahmedabdallh123/zahra"
GITHUB_TOKEN = st.secrets.get("github", {}).get("token", None)
GITHUB_AVAILABLE = GITHUB_TOKEN is not None
ACTIVITY_LOG_FILE = "activity_log.json"

# ------------------------------- Email Functions -------------------------------
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Optional

def get_email_config():
    """Read email configuration from secrets."""
    try:
        host = st.secrets["email"]["host"]
        port = int(st.secrets["email"]["port"])
        username = st.secrets["email"]["username"]
        password = st.secrets["email"]["password"].replace(" ", "")
        recipients_str = st.secrets["email"]["recipients"]
        recipients = [r.strip() for r in recipients_str.split(",") if r.strip()]
        return {
            "host": host,
            "port": port,
            "username": username,
            "password": password,
            "recipients": recipients
        }
    except Exception:
        return None

def send_email(subject: str, body: str, recipients: Optional[List[str]] = None) -> bool:
    config = get_email_config()
    if not config:
        st.warning("Email configuration is incomplete. Notifications will not be sent.")
        return False
    if recipients is None:
        recipients = config["recipients"]
    if not recipients:
        st.warning("No email recipients configured.")
        return False
    try:
        msg = MIMEMultipart()
        msg["From"] = config["username"]
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))
        server = smtplib.SMTP(config["host"], config["port"])
        server.starttls()
        server.login(config["username"], config["password"])
        server.sendmail(config["username"], recipients, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Failed to send email: {e}")
        return False

def get_current_notifications_text() -> str:
    all_sheets = load_all_sheets()
    if not all_sheets:
        return "No data available currently."

    existing_sections = [name for name in all_sheets.keys()
                        if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]

    existing_equipment = []
    for sheet_name in existing_sections:
        df = all_sheets.get(sheet_name)
        if df is not None and "Equipment" in df.columns:
            existing_equipment.extend(df["Equipment"].dropna().unique())
    existing_equipment = [str(eq).strip() for eq in existing_equipment if str(eq).strip() != ""]

    username = st.session_state.get("username")
    allowed_sections = get_allowed_sections(all_sheets, username, "view")
    allowed_sections = [sec for sec in allowed_sections if sec in existing_sections]

    overdue, upcoming = get_upcoming_maintenance(3)

    if not overdue.empty and "Equipment" in overdue.columns:
        overdue = overdue[overdue["Equipment"].isin(existing_equipment)]
    if not upcoming.empty and "Equipment" in upcoming.columns:
        upcoming = upcoming[upcoming["Equipment"].isin(existing_equipment)]

    if username != "admin":
        allowed_equipment = []
        for sheet_name in allowed_sections:
            df = all_sheets.get(sheet_name)
            if df is not None and "Equipment" in df.columns:
                allowed_equipment.extend(df["Equipment"].dropna().unique())
        allowed_equipment = [str(eq).strip() for eq in allowed_equipment if str(eq).strip() != ""]
        overdue = overdue[overdue["Equipment"].isin(allowed_equipment)] if not overdue.empty else overdue
        upcoming = upcoming[upcoming["Equipment"].isin(allowed_equipment)] if not upcoming.empty else upcoming

    parts = []
    for _, row in overdue.iterrows():
        eq = row['Equipment']
        task = row['Task_Name']
        due_date = row['Next_Date'].strftime('%Y-%m-%d') if pd.notna(row['Next_Date']) else "Not Set"
        parts.append(f"OVERDUE: {eq} - {task} (Due: {due_date})")

    for _, row in upcoming.iterrows():
        eq = row['Equipment']
        task = row['Task_Name']
        days = (row['Next_Date'].date() - datetime.now().date()).days
        due_date = row['Next_Date'].strftime('%Y-%m-%d') if pd.notna(row['Next_Date']) else "Not Set"
        parts.append(f"UPCOMING: {eq} - {task} (in {days} days - {due_date})")

    critical = get_critical_spare_parts()
    if username != "admin":
        critical = [p for p in critical if p.get("Section", "") in allowed_sections]

    for part in critical:
        parts.append(f"CRITICAL PART: {part['Part_Name']} (Qty: {part['Available_Quantity']} < Threshold: {part['Alert_Threshold']}) [Section: {part['Section']}]")

    if not parts:
        return "No critical notifications currently."
    return "\n".join(parts)

# ------------------------------- Image Upload Functions -------------------------------
def upload_image_to_github(image_file, entity_type, entity_id, custom_filename=None):
    if not GITHUB_AVAILABLE:
        st.error("GitHub token not available, cannot upload images.")
        return None
    try:
        img = Image.open(image_file)
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGB')
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=85, optimize=True)
        buffer.seek(0)
        if custom_filename:
            filename = custom_filename
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{entity_type}_{entity_id}_{timestamp}.jpg"
        repo_path = f"{IMAGES_FOLDER}/{entity_type}/{filename}"
        g = Github(GITHUB_TOKEN)
        repo = g.get_repo(APP_CONFIG["REPO_NAME"])
        try:
            repo.get_contents(f"{IMAGES_FOLDER}/{entity_type}/", ref=APP_CONFIG["BRANCH"])
        except GithubException:
            repo.create_file(f"{IMAGES_FOLDER}/{entity_type}/.gitkeep", f"Create folder for {entity_type} images", "", branch=APP_CONFIG["BRANCH"])
        content = buffer.getvalue()
        repo.create_file(path=repo_path, message=f"Add image for {entity_type} {entity_id}", content=content, branch=APP_CONFIG["BRANCH"])
        return f"https://raw.githubusercontent.com/{APP_CONFIG['REPO_NAME']}/{APP_CONFIG['BRANCH']}/{repo_path}"
    except Exception as e:
        st.error(f"Error processing image: {e}")
        return None

def get_image_component(image_url, caption=""):
    if not image_url or not isinstance(image_url, str):
        return None
    try:
        return st.image(image_url, caption=caption, use_container_width=True)
    except:
        st.warning(f"Could not display image: {image_url}")
        return None

# ------------------------------- Support Config -------------------------------
def load_support_config():
    default_config = {"image_url": "", "youtube_link": ""}
    if GITHUB_AVAILABLE:
        try:
            g = Github(GITHUB_TOKEN)
            repo = g.get_repo(APP_CONFIG["REPO_NAME"])
            contents = repo.get_contents(SUPPORT_CONFIG_FILE, ref=APP_CONFIG["BRANCH"])
            import base64
            content = base64.b64decode(contents.content).decode('utf-8')
            return json.loads(content)
        except:
            pass
    if os.path.exists(SUPPORT_CONFIG_FILE):
        try:
            with open(SUPPORT_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return default_config
    return default_config

def save_support_config(config):
    config_str = json.dumps(config, indent=2, ensure_ascii=False)
    if GITHUB_AVAILABLE:
        try:
            g = Github(GITHUB_TOKEN)
            repo = g.get_repo(APP_CONFIG["REPO_NAME"])
            try:
                contents = repo.get_contents(SUPPORT_CONFIG_FILE, ref=APP_CONFIG["BRANCH"])
                repo.update_file(SUPPORT_CONFIG_FILE, "Update support config", config_str, contents.sha, branch=APP_CONFIG["BRANCH"])
            except:
                repo.create_file(SUPPORT_CONFIG_FILE, "Create support config", config_str, branch=APP_CONFIG["BRANCH"])
        except:
            pass
    with open(SUPPORT_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

# ------------------------------- Spare Parts Functions -------------------------------
def load_spare_parts():
    if not os.path.exists(APP_CONFIG["LOCAL_FILE"]):
        return pd.DataFrame(columns=APP_CONFIG["SPARE_PARTS_COLUMNS"])
    try:
        df = pd.read_excel(APP_CONFIG["LOCAL_FILE"], sheet_name=APP_CONFIG["SPARE_PARTS_SHEET"])
        df.columns = df.columns.astype(str).str.strip()
        for col in APP_CONFIG["SPARE_PARTS_COLUMNS"]:
            if col not in df.columns:
                df[col] = ""
        df = df.fillna("")
        df["Available_Quantity"] = pd.to_numeric(df["Available_Quantity"], errors='coerce').fillna(0)
        if "Alert_Threshold" not in df.columns:
            df["Alert_Threshold"] = 1
        else:
            df["Alert_Threshold"] = pd.to_numeric(df["Alert_Threshold"], errors='coerce').fillna(1)
        return df
    except Exception:
        return pd.DataFrame(columns=APP_CONFIG["SPARE_PARTS_COLUMNS"])

def get_spare_parts_for_section(section_name):
    df = load_spare_parts()
    if df.empty:
        return []
    filtered = df[(df["Section"] == section_name) | (df["Section"] == APP_CONFIG["GENERAL_SECTION"])]
    return list(zip(filtered["Part_Name"], filtered["Available_Quantity"]))

def consume_spare_part(part_name, quantity=1):
    df = load_spare_parts()
    if df.empty:
        return False, "No spare parts registered.", None
    mask = df["Part_Name"] == part_name
    if not mask.any():
        return False, f"Part '{part_name}' not found.", None
    current_qty = df.loc[mask, "Available_Quantity"].values[0]
    if current_qty < quantity:
        return False, f"Insufficient stock (Available: {current_qty}, Requested: {quantity})", current_qty
    new_qty = current_qty - quantity
    df.loc[mask, "Available_Quantity"] = new_qty
    st.session_state.temp_spare_parts_df = df
    return True, f"Deducted {quantity} of '{part_name}'. New balance: {new_qty}", new_qty

def get_critical_spare_parts():
    df = load_spare_parts()
    if df.empty:
        return []
    df["Available_Quantity"] = pd.to_numeric(df["Available_Quantity"], errors='coerce').fillna(0)
    if "Alert_Threshold" not in df.columns:
        df["Alert_Threshold"] = 1
    else:
        df["Alert_Threshold"] = pd.to_numeric(df["Alert_Threshold"], errors='coerce').fillna(1)
    if "Section" not in df.columns:
        return []
    df["Section"] = df["Section"].fillna("").astype(str)
    df = df[df["Section"].str.strip() != ""]
    df["Is_Critical"] = df["Is_Critical"].astype(str).str.strip()
    critical = df[(df["Is_Critical"] == "Yes") & (df["Available_Quantity"] < df["Alert_Threshold"])]
    result = critical[["Part_Name", "Section", "Available_Quantity", "Alert_Threshold"]].to_dict('records')
    return result

# ------------------------------- Duplicate Event Check -------------------------------
def is_duplicate_event(df, new_row, compare_columns=None, ignore_columns=None, time_window_days=0):
    if df.empty:
        return False
    default_compare = ["Date", "Equipment", "Event_Fault", "Corrective_Action", "Performed_By"]
    if compare_columns is None:
        compare_columns = default_compare
    else:
        compare_columns = list(set(default_compare + compare_columns))
    if ignore_columns:
        compare_columns = [col for col in compare_columns if col not in ignore_columns]
    available_cols = [col for col in compare_columns if col in df.columns and col in new_row]
    if not available_cols:
        return False
    mask = pd.Series([True] * len(df))
    for col in available_cols:
        mask &= (df[col].astype(str).str.strip() == str(new_row.get(col, "")).strip())
    if mask.any() and time_window_days > 0 and "Date" in df.columns and "Date" in new_row:
        new_date = pd.to_datetime(new_row["Date"]).date()
        matching_rows = df[mask].copy()
        if not matching_rows.empty:
            matching_rows["date_only"] = pd.to_datetime(matching_rows["Date"]).dt.date
            date_diff = (new_date - matching_rows["date_only"]).abs()
            if (date_diff <= timedelta(days=time_window_days)).any():
                return True
            else:
                return False
    else:
        return mask.any()
    return False

# ------------------------------- User Management Functions -------------------------------
def load_users_from_github():
    try:
        response = requests.get(GITHUB_USERS_URL, timeout=10)
        response.raise_for_status()
        users_data = response.json()
        for username, info in users_data.items():
            if "permissions" in info and isinstance(info["permissions"], list):
                if "all" in info["permissions"]:
                    info["permissions"] = {"all_sections": True}
                else:
                    info["permissions"] = {"all_sections": False}
            elif "permissions" not in info:
                info["permissions"] = {"all_sections": False}
            if "sections_permissions" not in info:
                info["sections_permissions"] = {}
        return users_data
    except Exception as e:
        st.error(f"Failed to load users from GitHub: {e}")
        if os.path.exists(USERS_FILE):
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return {"admin": {"password": "1234", "role": "admin", "permissions": {"all_sections": True}, "sections_permissions": {}}}

def save_users_to_github(users_data):
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users_data, f, indent=4, ensure_ascii=False)
        token = st.secrets.get("github", {}).get("token", None)
        if not token:
            st.warning("GitHub token not available, saved locally only.")
            return True
        g = Github(token)
        repo = g.get_repo(GITHUB_REPO_USERS)
        users_json = json.dumps(users_data, indent=4, ensure_ascii=False, sort_keys=True)
        try:
            contents = repo.get_contents("users.json", ref="main")
            repo.update_file(path="users.json", message="Update user permissions", content=users_json, sha=contents.sha, branch="main")
        except GithubException as e:
            if e.status == 404:
                repo.create_file(path="users.json", message="Create users file", content=users_json, branch="main")
            else:
                raise
        return True
    except Exception as e:
        st.error(f"Failed to upload users to GitHub: {e}")
        st.warning("Saved locally only, not pushed to GitHub.")
        return True

def get_all_sections_from_excel():
    sheets = load_all_sheets()
    if not sheets:
        return []
    return [name for name in sheets.keys() if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]

def admin_users_management_tab():
    st.header("👥 User & Permission Management")
    st.info("Here you can add, edit, or delete users and assign permissions per section.")
    users = load_users_from_github()
    sections_list = get_all_sections_from_excel()
    if not sections_list:
        st.warning("No sections available. Please add a section first from 'Add New Section' tab.")

    st.subheader("📋 User List")
    for username, info in users.items():
        with st.expander(f"👤 {username} (Role: {info.get('role', 'viewer')})"):
            col1, col2 = st.columns(2)
            with col1:
                new_password = st.text_input(f"New Password", type="password", key=f"pass_{username}")
                if new_password:
                    if st.button(f"🔐 Change Password", key=f"change_pass_{username}"):
                        users[username]["password"] = new_password
                        if save_users_to_github(users):
                            st.success(f"Password changed for {username}")
                            st.rerun()
                        else:
                            st.error("Failed to save changes")
            with col2:
                current_role = info.get("role", "viewer")
                role_options = ["admin", "editor", "viewer"]
                if current_role not in role_options:
                    current_role = "viewer"
                new_role = st.selectbox(f"Role", role_options, index=role_options.index(current_role), key=f"role_{username}")
                if new_role != info.get("role"):
                    users[username]["role"] = new_role
                    if save_users_to_github(users):
                        st.success(f"Role changed for {username} to {new_role}")
                        st.rerun()

            # Email field
            current_email = info.get("email", "")
            new_email = st.text_input(
                "📧 Email (for notifications):",
                value=current_email,
                key=f"email_{username}",
                placeholder="example@domain.com"
            )
            if new_email != current_email:
                if st.button(f"💾 Save Email", key=f"save_email_{username}"):
                    users[username]["email"] = new_email.strip()
                    if save_users_to_github(users):
                        st.success(f"Email saved for {username}")
                        st.rerun()
                    else:
                        st.error("Failed to save email")

            st.markdown("#### Section Permissions")
            all_sections_access = st.checkbox(
                "Grant access to all sections (no per-section detail)",
                value=info.get("permissions", {}).get("all_sections", False),
                key=f"all_sections_{username}"
            )
            if all_sections_access:
                users[username]["permissions"] = {"all_sections": True}
                users[username]["sections_permissions"] = {}
                st.info("This user has access to all current and future sections.")
            else:
                users[username]["permissions"] = {"all_sections": False}
                if "sections_permissions" not in users[username]:
                    users[username]["sections_permissions"] = {}
                if sections_list:
                    for section in sections_list:
                        current_perms = users[username]["sections_permissions"].get(section, [])
                        st.markdown(f"**{section}**")
                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            view_perm = st.checkbox("View", value=("view" in current_perms), key=f"view_{username}_{section}")
                        with c2:
                            edit_perm = st.checkbox("Edit", value=("edit" in current_perms), key=f"edit_{username}_{section}")
                        with c3:
                            add_event_perm = st.checkbox("Add Event", value=("add_event" in current_perms), key=f"add_{username}_{section}")
                        with c4:
                            manage_perm = st.checkbox("Manage Machines", value=("manage_machines" in current_perms), key=f"manage_{username}_{section}")
                        new_perms = []
                        if view_perm: new_perms.append("view")
                        if edit_perm: new_perms.append("edit")
                        if add_event_perm: new_perms.append("add_event")
                        if manage_perm: new_perms.append("manage_machines")
                        users[username]["sections_permissions"][section] = new_perms
                else:
                    st.info("No sections available currently.")

            if st.button(f"💾 Save Permissions for {username}", key=f"save_perms_{username}"):
                if save_users_to_github(users):
                    st.success(f"Permissions saved for {username}")
                    st.rerun()
                else:
                    st.error("Failed to save")

            if username != "admin":
                st.markdown("---")
                col_del1, col_del2 = st.columns([3, 1])
                with col_del1:
                    st.warning(f"Deleting user **{username}** permanently. This cannot be undone.")
                with col_del2:
                    if st.button(f"🗑️ Delete", key=f"delete_btn_{username}", type="primary"):
                        del users[username]
                        if save_users_to_github(users):
                            st.success(f"User {username} deleted successfully.")
                            st.rerun()
                        else:
                            st.error("Failed to delete user. Try again.")

    st.markdown("---")
    st.subheader("➕ Add New User")
    with st.form("add_user_form"):
        col1, col2 = st.columns(2)
        with col1:
            new_username = st.text_input("Username (letters and numbers only)")
            new_password = st.text_input("Password", type="password")
            new_email = st.text_input("📧 Email (optional):", placeholder="example@domain.com")
        with col2:
            new_role = st.selectbox("Default Role", ["viewer", "editor", "admin"])
            st.caption("You can adjust per-section permissions later")
        submitted = st.form_submit_button("➕ Add User", type="primary")
        if submitted:
            if not new_username or not new_password:
                st.error("Username and password are required")
            elif new_username in users:
                st.error(f"User '{new_username}' already exists")
            elif not new_username.replace("_", "").isalnum():
                st.error("Username must contain only letters and numbers")
            else:
                users[new_username] = {
                    "password": new_password,
                    "role": new_role,
                    "permissions": {"all_sections": False},
                    "sections_permissions": {},
                    "email": new_email.strip()
                }
                if save_users_to_github(users):
                    st.success(f"User {new_username} added")
                    st.balloons()
                    st.rerun()
                else:
                    st.error("Failed to save new user")

# ------------------------------- Activity Log Functions -------------------------------
def log_activity(action_type, details, username=None, section=None):
    if username is None:
        username = st.session_state.get("username", "unknown")
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "username": username,
        "action_type": action_type,
        "details": details,
        "section": section
    }
    log = []
    if os.path.exists(ACTIVITY_LOG_FILE):
        try:
            with open(ACTIVITY_LOG_FILE, "r", encoding="utf-8") as f:
                log = json.load(f)
        except:
            log = []
    log.append(log_entry)
    if len(log) > 200:
        log = log[-200:]
    with open(ACTIVITY_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(log, f, indent=2, ensure_ascii=False)
    if GITHUB_AVAILABLE:
        try:
            g = Github(GITHUB_TOKEN)
            repo = g.get_repo(APP_CONFIG["REPO_NAME"])
            content = json.dumps(log, indent=2, ensure_ascii=False)
            try:
                contents = repo.get_contents(ACTIVITY_LOG_FILE, ref=APP_CONFIG["BRANCH"])
                repo.update_file(ACTIVITY_LOG_FILE, "Update activity log", content, contents.sha, branch=APP_CONFIG["BRANCH"])
            except:
                repo.create_file(ACTIVITY_LOG_FILE, "Create activity log", content, branch=APP_CONFIG["BRANCH"])
        except:
            pass

def clean_old_activity_log(days_to_keep=1):
    log = load_activity_log()
    if not log:
        return
    cutoff = datetime.now() - timedelta(days=days_to_keep)
    new_log = []
    for entry in log:
        try:
            entry_time = datetime.fromisoformat(entry["timestamp"])
            if entry_time >= cutoff:
                new_log.append(entry)
        except:
            new_log.append(entry)
    if len(new_log) != len(log):
        with open(ACTIVITY_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(new_log, f, indent=2, ensure_ascii=False)
        if GITHUB_AVAILABLE:
            try:
                g = Github(GITHUB_TOKEN)
                repo = g.get_repo(APP_CONFIG["REPO_NAME"])
                content = json.dumps(new_log, indent=2, ensure_ascii=False)
                try:
                    contents = repo.get_contents(ACTIVITY_LOG_FILE, ref=APP_CONFIG["BRANCH"])
                    repo.update_file(ACTIVITY_LOG_FILE, "Clean old log", content, contents.sha, branch=APP_CONFIG["BRANCH"])
                except:
                    repo.create_file(ACTIVITY_LOG_FILE, "Create activity log", content, branch=APP_CONFIG["BRANCH"])
            except:
                pass

def load_activity_log():
    if GITHUB_AVAILABLE:
        try:
            g = Github(GITHUB_TOKEN)
            repo = g.get_repo(APP_CONFIG["REPO_NAME"])
            contents = repo.get_contents(ACTIVITY_LOG_FILE, ref=APP_CONFIG["BRANCH"])
            import base64
            content = base64.b64decode(contents.content).decode('utf-8')
            log = json.loads(content)
            log.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
            return log
        except:
            pass
    if os.path.exists(ACTIVITY_LOG_FILE):
        with open(ACTIVITY_LOG_FILE, "r", encoding="utf-8") as f:
            log = json.load(f)
            log.sort(key=lambda x: x.get("timestamp", ""), reverse=True)
            return log
    return []

# ------------------------------- Preventive Maintenance Functions -------------------------------
def load_maintenance_tasks():
    if not os.path.exists(APP_CONFIG["LOCAL_FILE"]):
        return pd.DataFrame(columns=APP_CONFIG["MAINTENANCE_COLUMNS"])
    try:
        df = pd.read_excel(APP_CONFIG["LOCAL_FILE"], sheet_name=APP_CONFIG["MAINTENANCE_SHEET"])
        df.columns = df.columns.astype(str).str.strip()
        for col in APP_CONFIG["MAINTENANCE_COLUMNS"]:
            if col not in df.columns:
                df[col] = ""
        df = df.fillna("")
        if "Last_Execution" in df.columns:
            df["Last_Execution"] = pd.to_datetime(df["Last_Execution"], errors='coerce')
        if "Next_Date" in df.columns:
            df["Next_Date"] = pd.to_datetime(df["Next_Date"], errors='coerce')
        if "Period_Days" in df.columns:
            df["Period_Days"] = pd.to_numeric(df["Period_Days"], errors='coerce').fillna(0)
        return df
    except Exception:
        return pd.DataFrame(columns=APP_CONFIG["MAINTENANCE_COLUMNS"])

def get_tasks_for_equipment(equipment_name):
    df = load_maintenance_tasks()
    if df.empty:
        return df
    return df[df["Equipment"] == equipment_name]

def add_maintenance_task(sheets_edit, equipment, task_name, period_hours, start_date=None, notes="", default_spare="", image_url=None, section=None):
    if APP_CONFIG["MAINTENANCE_SHEET"] not in sheets_edit:
        sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = pd.DataFrame(columns=APP_CONFIG["MAINTENANCE_COLUMNS"])
    df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
    if start_date is None:
        start_date = datetime.now().date()
    period_days = period_hours / 24.0
    next_date = start_date + timedelta(days=period_days)
    new_row = pd.DataFrame([{
        "Equipment": equipment, "Maintenance_Type": f"{period_hours} hours", "Task_Name": task_name,
        "Period_Days": period_days, "Last_Execution": pd.NaT, "Next_Date": next_date,
        "Notes": notes, "Default_Spare_Part": default_spare, "Image_URL": image_url or ""
    }])
    new_df = pd.concat([df, new_row], ignore_index=True)
    sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = new_df
    if section is None:
        for sheet_name, sh_df in sheets_edit.items():
            if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
                continue
            if equipment in get_equipment_list_from_sheet(sh_df):
                section = sheet_name
                break
        if section is None:
            section = "Unknown"
    log_activity("add_maintenance_task", f"Added maintenance task '{task_name}' for equipment {equipment} (period {period_hours} hours)", section=section)
    return sheets_edit

def get_upcoming_maintenance(days_ahead=3):
    df = load_maintenance_tasks()
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    today = datetime.now().date()
    overdue = df[df["Next_Date"] < pd.Timestamp(today)]
    upcoming = df[(df["Next_Date"] >= pd.Timestamp(today)) & (df["Next_Date"] <= pd.Timestamp(today + timedelta(days=days_ahead)))]
    return overdue, upcoming

def clean_orphan_maintenance_tasks(sheets_edit):
    """Remove maintenance tasks for equipment that does not belong to any existing section."""
    if APP_CONFIG["MAINTENANCE_SHEET"] not in sheets_edit:
        return sheets_edit

    df_maintenance = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
    if df_maintenance.empty or "Equipment" not in df_maintenance.columns:
        return sheets_edit

    existing_equipment = []
    for sheet_name, df in sheets_edit.items():
        if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
            continue
        if "Equipment" in df.columns:
            existing_equipment.extend(df["Equipment"].dropna().unique())
    existing_equipment = [str(eq).strip() for eq in existing_equipment if str(eq).strip() != ""]

    if existing_equipment:
        df_maintenance = df_maintenance[df_maintenance["Equipment"].isin(existing_equipment)]
    else:
        df_maintenance = pd.DataFrame(columns=APP_CONFIG["MAINTENANCE_COLUMNS"])

    sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = df_maintenance
    return sheets_edit

# ------------------------------- Failure Analysis Functions -------------------------------
def flexible_date_parser(date_series):
    def parse_single(val):
        if pd.isna(val) or val == "":
            return pd.NaT
        if isinstance(val, (pd.Timestamp, datetime)):
            return val
        val_str = str(val).strip().replace('\\', '/')
        for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y', '%Y/%m/%d']:
            try:
                return pd.to_datetime(val_str, format=fmt, errors='raise')
            except:
                pass
        return pd.to_datetime(val_str, errors='coerce')
    return date_series.apply(parse_single)

def analyze_time_between_corrections(df, filter_text=None):
    if df is None or df.empty:
        return pd.DataFrame()
    data = df.copy()
    if "Date" not in data.columns or "Equipment" not in data.columns or "Corrective_Action" not in data.columns:
        return pd.DataFrame()
    data["Date"] = flexible_date_parser(data["Date"])
    data = data.dropna(subset=["Date"]).sort_values(["Equipment", "Date"])
    if filter_text:
        data["Corrective_Action"] = data["Corrective_Action"].fillna("").astype(str)
        data = data[data["Corrective_Action"].str.contains(filter_text, case=False, na=False)]
    results = []
    for equipment in data["Equipment"].unique():
        eq_data = data[data["Equipment"] == equipment].copy()
        if len(eq_data) < 2:
            continue
        for i in range(len(eq_data)-1):
            current = eq_data.iloc[i]
            next_row = eq_data.iloc[i+1]
            gap_days = (next_row["Date"] - current["Date"]).total_seconds() / (24 * 3600)
            prev_correction = eq_data.iloc[i-1]["Corrective_Action"] if i > 0 else None
            prev_date = eq_data.iloc[i-1]["Date"] if i > 0 else None
            results.append({
                "Equipment": equipment,
                "Previous_Action": prev_correction if prev_correction else "---",
                "Previous_Action_Date": prev_date.strftime("%Y-%m-%d") if prev_date else "---",
                "Next_Action": next_row["Corrective_Action"],
                "Next_Action_Date": next_row["Date"].strftime("%Y-%m-%d"),
                "Gap_Days": round(gap_days, 1)
            })
    result_df = pd.DataFrame(results)
    if result_df.empty:
        return pd.DataFrame()
    result_df.reset_index(drop=True, inplace=True)
    return result_df

def failures_analysis_tab(all_sheets):
    st.header("📊 Corrective Actions Analysis")
    if not all_sheets:
        st.warning("No data for analysis")
        return
    username = st.session_state.get("username")
    allowed_sections = get_allowed_sections(all_sheets, username, "view")
    if not allowed_sections:
        st.warning("No sections available for analysis")
        return
    selected_section = st.selectbox("🏭 Select Section:", allowed_sections, key="analysis_section")
    df = all_sheets[selected_section].copy()
    if "Equipment" not in df.columns:
        st.error(f"Section '{selected_section}' does not have an 'Equipment' column")
        return
    df["Equipment"] = df["Equipment"].astype(str).str.strip()
    equipment_list = get_equipment_list_from_sheet(df)
    if not equipment_list:
        st.warning(f"No machines registered in section '{selected_section}'")
        return
    equipment_options = ["All Machines"] + equipment_list
    selected_equipment = st.selectbox("🔧 Select Machine:", equipment_options, key="analysis_equipment")
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("📅 From Date (optional):", value=None, key="start_date_filter")
    with col2:
        end_date = st.date_input("📅 To Date (optional):", value=None, key="end_date_filter")
    search_text = st.text_input("🔍 Search text in corrective action (optional):", placeholder="e.g. belt, cooler, 1270", key="search_text_analysis")
    if st.button("🔄 Run Analysis", key="run_analysis", type="primary"):
        filtered_df = df.copy()
        if selected_equipment != "All Machines":
            filtered_df = filtered_df[filtered_df["Equipment"] == selected_equipment]
        if "Date" in filtered_df.columns:
            filtered_df["Date"] = flexible_date_parser(filtered_df["Date"])
            filtered_df = filtered_df.dropna(subset=["Date"])
            if start_date:
                filtered_df = filtered_df[filtered_df["Date"] >= pd.to_datetime(start_date)]
            if end_date:
                filtered_df = filtered_df[filtered_df["Date"] <= pd.to_datetime(end_date) + timedelta(days=1)]
        if filtered_df.empty:
            st.warning("No data matches the filter criteria")
            return
        details_gaps = analyze_time_between_corrections(filtered_df, search_text if search_text else None)
        if "Corrective_Action" in filtered_df.columns:
            top_corrections = filtered_df["Corrective_Action"].value_counts().reset_index().head(10)
            top_corrections.columns = ["Corrective_Action", "Count"]
        else:
            top_corrections = pd.DataFrame()
        if selected_equipment == "All Machines" and "Equipment" in filtered_df.columns:
            top_equipment = filtered_df["Equipment"].value_counts().reset_index().head(10)
            top_equipment.columns = ["Equipment", "Faults_Count"]
        else:
            top_equipment = pd.DataFrame()
        st.success(f"Found {len(filtered_df)} corrective actions")
        if not top_corrections.empty:
            st.subheader("🔝 Most Frequent Corrective Actions")
            st.dataframe(top_corrections, use_container_width=True)
        if not top_equipment.empty:
            st.subheader("🏭 Machines Requiring Most Corrective Actions")
            st.dataframe(top_equipment, use_container_width=True)
        st.subheader("📋 Detailed Time Gaps Between Repeated Corrective Actions")
        if search_text:
            st.info(f"Gaps are calculated only between actions containing: **'{search_text}'**")
        if details_gaps.empty:
            st.info("Not enough data to compute gaps (need at least 2 actions per machine matching the search text)")
        else:
            st.dataframe(details_gaps, use_container_width=True, height=500)
            csv = details_gaps.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download Detailed Gaps CSV", csv, "detailed_corrections_gaps.csv", "text/csv")
        st.markdown("---")
        st.subheader("📥 Export Full Report (Excel)")
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, sheet_name="Original_Data", index=False)
            if not top_corrections.empty:
                top_corrections.to_excel(writer, sheet_name="Top_Corrections", index=False)
            if not top_equipment.empty:
                top_equipment.to_excel(writer, sheet_name="Top_Equipment", index=False)
            if not details_gaps.empty:
                details_gaps.to_excel(writer, sheet_name="Detailed_Gaps", index=False)
        excel_buffer.seek(0)
        st.download_button(
            "📥 Download Report (Excel)",
            excel_buffer,
            f"corrections_analysis_{selected_section}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------------------- User / Permission Loaders -------------------------------
def download_users_from_github():
    try:
        response = requests.get(GITHUB_USERS_URL, timeout=10)
        response.raise_for_status()
        users_data = response.json()
        for username, user_info in users_data.items():
            if "permissions" in user_info and isinstance(user_info["permissions"], list):
                if "all" in user_info["permissions"]:
                    user_info["permissions"] = {"all_sections": True}
                else:
                    user_info["permissions"] = {"all_sections": False}
                user_info["sections_permissions"] = {}
            elif "permissions" not in user_info:
                user_info["permissions"] = {"all_sections": False}
                user_info["sections_permissions"] = {}
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users_data, f, indent=4, ensure_ascii=False)
        return users_data
    except Exception as e:
        st.warning(f"Could not load users.json from GitHub: {e}")
        if os.path.exists(USERS_FILE):
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

def upload_users_to_github(users_data):
    try:
        token = st.secrets.get("github", {}).get("token", None)
        if not token:
            st.error("GitHub token not found")
            return False
        g = Github(token)
        repo = g.get_repo(GITHUB_REPO_USERS)
        users_json = json.dumps(users_data, indent=4, ensure_ascii=False, sort_keys=True)
        try:
            contents = repo.get_contents("users.json", ref="main")
            repo.update_file(path="users.json", message="Update users file", content=users_json, sha=contents.sha, branch="main")
            return True
        except:
            repo.create_file(path="users.json", message="Create users file", content=users_json, branch="main")
            return True
    except Exception as e:
        st.error(f"Failed to upload users: {e}")
        return False

def load_users():
    try:
        users_data = download_users_from_github()
        if not users_data or "admin" not in users_data:
            if os.path.exists(USERS_FILE):
                with open(USERS_FILE, "r", encoding="utf-8") as f:
                    local_users = json.load(f)
                    if "admin" in local_users:
                        return local_users
            default_users = {
                "admin": {"password": "1234", "role": "admin", "permissions": {"all_sections": True}, "sections_permissions": {}, "email": ""},
                "maintenance_manager": {"password": "12345", "role": "admin", "permissions": {"all_sections": True}, "sections_permissions": {}, "email": ""}
            }
            return default_users
        return users_data
    except Exception as e:
        st.error(f"Error loading users: {e}")
        return {"admin": {"password": "1234", "role": "admin", "permissions": {"all_sections": True}, "sections_permissions": {}, "email": ""}}

def load_state():
    if not os.path.exists(STATE_FILE):
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f)
        return {}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=4, ensure_ascii=False)

def cleanup_sessions(state):
    now = datetime.now()
    changed = False
    for user, info in list(state.items()):
        if info.get("active") and "login_time" in info:
            try:
                login_time = datetime.fromisoformat(info["login_time"])
                if now - login_time > SESSION_DURATION:
                    info["active"] = False
                    info.pop("login_time", None)
                    changed = True
            except:
                info["active"] = False
                changed = True
    if changed:
        save_state(state)
    return state

def remaining_time(state, username):
    if not username or username not in state:
        return None
    info = state.get(username)
    if not info or not info.get("active"):
        return None
    try:
        lt = datetime.fromisoformat(info["login_time"])
        remaining = SESSION_DURATION - (datetime.now() - lt)
        if remaining.total_seconds() <= 0:
            return None
        return remaining
    except:
        return None

def logout_action():
    state = load_state()
    username = st.session_state.get("username")
    if username and username in state:
        state[username]["active"] = False
        state[username].pop("login_time", None)
        save_state(state)
    for k in list(st.session_state.keys()):
        st.session_state.pop(k, None)
    st.rerun()

def login_ui():
    users = load_users()
    state = cleanup_sessions(load_state())
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username = None
        st.session_state.user_role = None
        st.session_state.user_permissions = []
    st.title(f"{APP_CONFIG['APP_ICON']} Login - {APP_CONFIG['APP_TITLE']}")
    username_input = st.selectbox("Select User", list(users.keys()))
    password = st.text_input("Password", type="password")
    active_users = [u for u, v in state.items() if v.get("active")]
    active_count = len(active_users)
    st.caption(f"Active users: {active_count} / {MAX_ACTIVE_USERS}")
    if not st.session_state.logged_in:
        if st.button("Login"):
            current_users = load_users()
            if username_input in current_users and current_users[username_input]["password"] == password:
                if username_input != "admin" and username_input in active_users:
                    st.warning("This user is already logged in.")
                    return False
                elif active_count >= MAX_ACTIVE_USERS and username_input != "admin":
                    st.error("Maximum number of active users reached.")
                    return False
                state[username_input] = {"active": True, "login_time": datetime.now().isoformat()}
                save_state(state)
                st.session_state.logged_in = True
                st.session_state.username = username_input
                st.session_state.user_role = current_users[username_input].get("role", "viewer")
                st.session_state.user_permissions = current_users[username_input].get("permissions", ["view"])
                st.success(f"Logged in: {username_input}")
                st.rerun()
            else:
                st.error("Invalid password.")
        return False
    else:
        st.success(f"Logged in as: {st.session_state.username}")
        rem = remaining_time(state, st.session_state.username)
        if rem:
            mins, secs = divmod(int(rem.total_seconds()), 60)
            st.info(f"Time remaining: {mins:02d}:{secs:02d}")
        if st.button("Logout"):
            logout_action()
        return True

# ------------------------------- Permissions Functions -------------------------------
def get_user_permissions(username):
    users = load_users()
    if username not in users:
        return {"all_sections": False, "sections_permissions": {}}
    user_data = users[username]
    if "permissions" in user_data and isinstance(user_data["permissions"], dict):
        perms = user_data["permissions"]
    elif "permissions" in user_data and isinstance(user_data["permissions"], list):
        perms = {"all_sections": "all" in user_data["permissions"]}
    else:
        perms = {"all_sections": False}
    if "sections_permissions" not in user_data:
        user_data["sections_permissions"] = {}
    return {"all_sections": perms.get("all_sections", False), "sections_permissions": user_data.get("sections_permissions", {})}

def has_section_permission(username, section_name, required_permission="view"):
    if username == "admin":
        return True
    permissions = get_user_permissions(username)
    if not permissions:
        return False
    if permissions.get("all_sections", False):
        return True
    section_perms = permissions.get("sections_permissions", {}).get(section_name, [])
    return required_permission in section_perms

def get_allowed_sections(all_sheets, username, required_permission="view"):
    allowed = []
    for sheet_name in all_sheets.keys():
        if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
            continue
        if has_section_permission(username, sheet_name, required_permission):
            allowed.append(sheet_name)
    return allowed

# ------------------------------- File Operations -------------------------------
def fetch_from_github_requests():
    try:
        response = requests.get(GITHUB_EXCEL_URL, stream=True, timeout=15)
        response.raise_for_status()
        with open(APP_CONFIG["LOCAL_FILE"], "wb") as f:
            shutil.copyfileobj(response.raw, f)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Update failed: {e}")
        return False

@st.cache_data(show_spinner=False)
def load_all_sheets():
    if not os.path.exists(APP_CONFIG["LOCAL_FILE"]):
        return None
    try:
        sheets = pd.read_excel(APP_CONFIG["LOCAL_FILE"], sheet_name=None)
        if not sheets:
            return None
        for name, df in sheets.items():
            if df.empty:
                continue
            df.columns = df.columns.astype(str).str.strip()
            df = df.fillna('')
            sheets[name] = df
        return sheets
    except Exception as e:
        st.error(f"Error loading sheets: {e}")
        return None

@st.cache_data(show_spinner=False)
def load_sheets_for_edit():
    if not os.path.exists(APP_CONFIG["LOCAL_FILE"]):
        return None
    try:
        sheets = pd.read_excel(APP_CONFIG["LOCAL_FILE"], sheet_name=None, dtype=object)
        if not sheets:
            return None
        for name, df in sheets.items():
            df.columns = df.columns.astype(str).str.strip()
            df = df.fillna('')
            sheets[name] = df
        return sheets
    except Exception as e:
        st.error(f"Error loading sheets: {e}")
        return None

def save_excel_locally(sheets_dict):
    try:
        if "temp_spare_parts_df" in st.session_state:
            sheets_dict[APP_CONFIG["SPARE_PARTS_SHEET"]] = st.session_state.temp_spare_parts_df
            del st.session_state.temp_spare_parts_df
        if APP_CONFIG["MAINTENANCE_SHEET"] not in sheets_dict:
            sheets_dict[APP_CONFIG["MAINTENANCE_SHEET"]] = load_maintenance_tasks()
        with pd.ExcelWriter(APP_CONFIG["LOCAL_FILE"], engine="openpyxl") as writer:
            for name, sh in sheets_dict.items():
                try:
                    sh.to_excel(writer, sheet_name=name, index=False)
                except Exception:
                    sh.astype(object).to_excel(writer, sheet_name=name, index=False)
        return True
    except Exception as e:
        st.error(f"Local save error: {e}")
        return False

def push_to_github():
    try:
        token = st.secrets.get("github", {}).get("token", None)
        if not token:
            st.error("GitHub token not found in secrets")
            return False
        if not GITHUB_AVAILABLE:
            st.error("PyGithub not available")
            return False
        g = Github(token)
        repo = g.get_repo(APP_CONFIG["REPO_NAME"])
        with open(APP_CONFIG["LOCAL_FILE"], "rb") as f:
            content = f.read()
        try:
            contents = repo.get_contents(APP_CONFIG["FILE_PATH"], ref=APP_CONFIG["BRANCH"])
            repo.update_file(path=APP_CONFIG["FILE_PATH"], message=f"Update data - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", content=content, sha=contents.sha, branch=APP_CONFIG["BRANCH"])
            st.success("Changes pushed to GitHub")
            return True
        except GithubException as e:
            if e.status == 404:
                repo.create_file(path=APP_CONFIG["FILE_PATH"], message=f"Create new file - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", content=content, branch=APP_CONFIG["BRANCH"])
                st.success("File created on GitHub")
                return True
            else:
                st.error(f"GitHub error: {e}")
                return False
    except Exception as e:
        st.error(f"Push failed: {e}")
        return False

def save_and_push_to_github(sheets_dict, operation_name):
    st.info(f"Saving {operation_name}...")
    if save_excel_locally(sheets_dict):
        st.success("Saved locally")
        if push_to_github():
            st.success("Pushed to GitHub")
            st.cache_data.clear()
            return True
        else:
            st.warning("Saved locally only")
            return True
    else:
        st.error("Local save failed")
        return False

# ------------------------------- Export & Display Functions -------------------------------
def export_sheet_to_excel(sheets_dict, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheets_dict[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

def export_all_sheets_to_excel(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

def export_filtered_results_to_excel(results_df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

def display_sheet_data(sheet_name, df, unique_id, sheets_edit):
    st.markdown(f"### 🏭 {sheet_name}")
    st.info(f"Machines: {len(df)} | Columns: {len(df.columns)}")
    equipment_list = get_equipment_list_from_sheet(df)
    if equipment_list and "Equipment" in df.columns:
        st.markdown("#### 🔍 Filter by Machine:")
        selected_filter = st.selectbox("Select machine:", ["All Machines"] + equipment_list, key=f"filter_{unique_id}")
        if selected_filter != "All Machines":
            df = df[df["Equipment"] == selected_filter]
            st.info(f"Showing: {selected_filter} - Records: {len(df)}")
    display_df = df.copy()
    for col in display_df.columns:
        if display_df[col].dtype == 'object':
            display_df[col] = display_df[col].astype(str).apply(lambda x: x[:100] + "..." if len(x) > 100 else x)
    if "Image_URL" in display_df.columns:
        display_df = display_df.drop(columns=["Image_URL"])
    st.dataframe(display_df, use_container_width=True, height=400)
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        excel_file = export_sheet_to_excel({sheet_name: df}, sheet_name)
        st.download_button("📥 Download section (Excel)", excel_file, f"{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"export_sheet_{unique_id}")
    with col_btn2:
        all_sheets_excel = export_all_sheets_to_excel({sheet_name: df})
        st.download_button("📥 Download all data (Excel)", all_sheets_excel, f"all_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"export_all_{unique_id}")

# ------------------------------- Advanced Search -------------------------------
def search_across_sheets(all_sheets):
    st.subheader("Advanced Search")
    if not all_sheets:
        st.warning("No data for search")
        return
    username = st.session_state.get("username")
    search_type = st.selectbox("Data type:", ["Sections (Faults)", "Spare Parts", "Preventive Maintenance"], key="search_type")

    existing_sections = [name for name in all_sheets.keys()
                        if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]
    allowed_sections = get_allowed_sections(all_sheets, username, "view")
    allowed_sections = [sec for sec in allowed_sections if sec in existing_sections]

    selected_section_filter = "All Sections"
    if search_type in ["Spare Parts", "Preventive Maintenance"] and allowed_sections:
        section_options = ["All Sections"] + allowed_sections
        selected_section_filter = st.selectbox("🏭 Section:", section_options, key="section_filter")

    if search_type == "Sections (Faults)":
        sheet_options = ["All Sections"] + allowed_sections
        selected_sheet = st.selectbox("Select Section:", sheet_options, key="search_sheet")
        if selected_sheet != "All Sections":
            equipment_list = get_equipment_list_from_sheet(all_sheets[selected_sheet])
        else:
            all_eq = set()
            for sh_name in allowed_sections:
                all_eq.update(get_equipment_list_from_sheet(all_sheets[sh_name]))
            equipment_list = sorted(all_eq)
        filter_equipment = st.selectbox("Machine filter:", ["All"] + equipment_list, key="search_eq")
        col1, col2 = st.columns(2)
        with col1:
            general_search = st.text_input("🔍 General search (Event/Action):", placeholder="e.g. oil leak...")
        with col2:
            technician_search = st.text_input("👨‍🔧 Technician search (Performed_By):", placeholder="Enter technician name...")
        st.markdown("#### 🏷️ Fault Type Filter")
        all_fault_types = set()
        if selected_sheet != "All Sections":
            df_temp = all_sheets[selected_sheet]
            if "Fault_Type" in df_temp.columns:
                all_fault_types.update(df_temp["Fault_Type"].dropna().unique())
        else:
            for sh_name in allowed_sections:
                df_temp = all_sheets[sh_name]
                if "Fault_Type" in df_temp.columns:
                    all_fault_types.update(df_temp["Fault_Type"].dropna().unique())
        fault_type_options = sorted([str(t).strip() for t in all_fault_types if str(t).strip() != ""])
        use_multiselect = st.checkbox("Choose from list", value=True, key="fault_type_multiselect_check")
        selected_fault_types = []
        custom_fault_type = ""
        if use_multiselect:
            if fault_type_options:
                selected_fault_types = st.multiselect("Select fault type:", fault_type_options, key="fault_type_multiselect")
            else:
                st.info("No fault types registered.")
        else:
            custom_fault_type = st.text_input("Or type fault type:", key="custom_fault_type_input")

        st.markdown("#### Date Range")
        use_date_filter = st.checkbox("Enable date search", key="use_date_filter_failures")
        if use_date_filter:
            col_date1, col_date2 = st.columns(2)
            with col_date1:
                start_date = st.date_input("From:", value=None, key="start_date_failures")
            with col_date2:
                end_date = st.date_input("To:", value=None, key="end_date_failures")
        else:
            start_date = None
            end_date = None
        view_mode = st.radio("View mode:", ["Table", "Cards with images"], horizontal=True, key="search_view_mode_failures")
        if st.button("Search", key="search_btn_failures", type="primary"):
            results = []
            sheets_to_search = []
            if selected_sheet != "All Sections":
                sheets_to_search = [(selected_sheet, all_sheets[selected_sheet])]
            else:
                for sheet_name in allowed_sections:
                    sheets_to_search.append((sheet_name, all_sheets[sheet_name]))
            for sheet_name, df in sheets_to_search:
                df_filtered = df.copy()
                for col in df_filtered.columns:
                    if df_filtered[col].dtype == 'object':
                        df_filtered[col] = df_filtered[col].astype(str).str.strip()
                if filter_equipment != "All" and "Equipment" in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered["Equipment"] == filter_equipment]
                if "Date" in df_filtered.columns:
                    df_filtered["Date"] = flexible_date_parser(df_filtered["Date"])
                    df_filtered = df_filtered.dropna(subset=["Date"])
                    if use_date_filter and start_date and end_date:
                        mask = (df_filtered["Date"] >= pd.to_datetime(start_date)) & (df_filtered["Date"] <= pd.to_datetime(end_date) + timedelta(days=1))
                        df_filtered = df_filtered[mask]
                if general_search:
                    search_clean = general_search.strip()
                    mask_general = pd.Series([False] * len(df_filtered))
                    if "Event_Fault" in df_filtered.columns:
                        mask_general = mask_general | df_filtered["Event_Fault"].astype(str).str.contains(search_clean, case=False, na=False)
                    if "Corrective_Action" in df_filtered.columns:
                        mask_general = mask_general | df_filtered["Corrective_Action"].astype(str).str.contains(search_clean, case=False, na=False)
                    df_filtered = df_filtered[mask_general]
                if technician_search and "Performed_By" in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered["Performed_By"].astype(str).str.contains(technician_search.strip(), case=False, na=False)]
                if use_multiselect and selected_fault_types and "Fault_Type" in df_filtered.columns:
                    pattern = '|'.join([t.strip() for t in selected_fault_types])
                    df_filtered = df_filtered[df_filtered["Fault_Type"].astype(str).str.contains(pattern, case=False, na=False)]
                elif custom_fault_type and "Fault_Type" in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered["Fault_Type"].astype(str).str.contains(custom_fault_type.strip(), case=False, na=False)]
                if not df_filtered.empty:
                    df_filtered["Section"] = sheet_name
                    results.append(df_filtered)
            if results:
                combined_results = pd.concat(results, ignore_index=True)
                if "Image_URL" in combined_results.columns:
                    combined_results["Unified_Image_URL"] = combined_results["Image_URL"]
                    combined_results = combined_results.drop(columns=["Image_URL"])
                else:
                    combined_results["Unified_Image_URL"] = ""
                st.success(f"Found {len(combined_results)} results")
                if "Date" in combined_results.columns:
                    combined_results["Date"] = pd.to_datetime(combined_results["Date"], errors='coerce')
                    combined_results = combined_results.dropna(subset=["Date"])
                    combined_results = combined_results.sort_values(by=["Equipment", "Date"], ascending=[True, False])
                if view_mode == "Table":
                    display_cols = [c for c in combined_results.columns if c not in ["Unified_Image_URL"]]
                    st.dataframe(combined_results[display_cols], use_container_width=True, height=500)
                else:
                    for idx, row in combined_results.iterrows():
                        with st.container(border=True):
                            col_img, col_info = st.columns([1, 3])
                            img_url = row.get("Unified_Image_URL", "")
                            with col_img:
                                if img_url and isinstance(img_url, str) and img_url.strip():
                                    try:
                                        st.image(img_url, use_container_width=True)
                                    except:
                                        st.write("Image unavailable")
                                else:
                                    st.write("No image")
                            with col_info:
                                st.markdown(f"**Section:** {row.get('Section', '')}")
                                st.markdown(f"**Date:** {row.get('Date', '')}")
                                st.markdown(f"**Equipment:** {row.get('Equipment', '')}")
                                st.markdown(f"**Fault:** {str(row.get('Event_Fault', ''))[:150]}")
                                st.markdown(f"**Action:** {str(row.get('Corrective_Action', ''))[:150]}")
                                st.markdown(f"**Performed By:** {row.get('Performed_By', '')}")
                                st.markdown(f"**Fault Type:** {row.get('Fault_Type', '')}")
                export_df = combined_results.drop(columns=["Unified_Image_URL"])
                excel_file = export_filtered_results_to_excel(export_df, "Search_Results")
                st.download_button("📥 Download search results (Excel)", excel_file, f"search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key='download-excel')
            else:
                st.warning("No results found")

    elif search_type == "Spare Parts":
        spare_df = load_spare_parts()
        if spare_df.empty:
            st.warning("No spare parts data")
            return
        for col in spare_df.columns:
            if spare_df[col].dtype == 'object':
                spare_df[col] = spare_df[col].astype(str).str.strip()
        df_filtered = spare_df.copy()
        if selected_section_filter != "All Sections":
            df_filtered = df_filtered[df_filtered["Section"] == selected_section_filter]
        search_term = st.text_input("🔍 Search term (part name, size, section...):", key="search_term_spare")
        if search_term:
            mask = pd.Series([False] * len(df_filtered))
            for col in ["Part_Name", "Size", "Tension", "Lead_Time", "Section"]:
                if col in df_filtered.columns:
                    mask = mask | df_filtered[col].astype(str).str.contains(search_term.strip(), case=False, na=False)
            df_filtered = df_filtered[mask]
        if not df_filtered.empty:
            st.success(f"Found {len(df_filtered)} parts")
            st.dataframe(df_filtered, use_container_width=True)
            excel_file = export_filtered_results_to_excel(df_filtered, "Spare_Parts")
            st.download_button("📥 Download results", excel_file, f"spare_parts_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("No results")

    else:  # Preventive Maintenance
        maint_df = load_maintenance_tasks()
        if maint_df.empty:
            st.warning("No preventive maintenance data")
            return
        for col in maint_df.columns:
            if maint_df[col].dtype == 'object':
                maint_df[col] = maint_df[col].astype(str).str.strip()
        equipment_to_section = {}
        for sheet_name in allowed_sections:
            df_sheet = all_sheets[sheet_name]
            if "Equipment" in df_sheet.columns:
                for eq in df_sheet["Equipment"].dropna().unique():
                    equipment_to_section[str(eq).strip()] = sheet_name
        df_filtered = maint_df.copy()
        if selected_section_filter != "All Sections":
            allowed_equipment = [eq for eq, sec in equipment_to_section.items() if sec == selected_section_filter]
            df_filtered = df_filtered[df_filtered["Equipment"].isin(allowed_equipment)]
        search_term = st.text_input("🔍 Search term (equipment, task, notes...):", key="search_term_maintenance")
        if search_term:
            mask = pd.Series([False] * len(df_filtered))
            for col in ["Equipment", "Task_Name", "Notes"]:
                if col in df_filtered.columns:
                    mask = mask | df_filtered[col].astype(str).str.contains(search_term.strip(), case=False, na=False)
            df_filtered = df_filtered[mask]
        st.markdown("#### Date Range")
        use_date_filter_maint = st.checkbox("Enable date search", key="use_date_filter_maintenance")
        if use_date_filter_maint:
            date_col_options = ["Last_Execution", "Next_Date"]
            available_date_cols = [col for col in date_col_options if col in df_filtered.columns]
            if not available_date_cols:
                st.warning("No date columns in preventive maintenance data")
                date_col_maint = None
            else:
                date_col_maint = st.selectbox("Select date column:", available_date_cols, key="date_col_maintenance")
            if date_col_maint:
                col_date1, col_date2 = st.columns(2)
                with col_date1:
                    start_date_maint = st.date_input("From:", value=None, key="start_date_maintenance")
                with col_date2:
                    end_date_maint = st.date_input("To:", value=None, key="end_date_maintenance")
            else:
                start_date_maint = None
                end_date_maint = None
        else:
            date_col_maint = None
            start_date_maint = None
            end_date_maint = None
        if use_date_filter_maint and date_col_maint and start_date_maint and end_date_maint:
            try:
                df_filtered[date_col_maint] = pd.to_datetime(df_filtered[date_col_maint], errors='coerce')
                df_filtered = df_filtered.dropna(subset=[date_col_maint])
                mask_date = (df_filtered[date_col_maint] >= pd.to_datetime(start_date_maint)) & (df_filtered[date_col_maint] <= pd.to_datetime(end_date_maint) + timedelta(days=1))
                df_filtered = df_filtered[mask_date]
                st.success(f"Date filter applied: {start_date_maint} to {end_date_maint}")
            except Exception as e:
                st.warning(f"Date filter error: {e}")
        if not df_filtered.empty:
            df_filtered["Section"] = df_filtered["Equipment"].map(equipment_to_section).fillna("Unknown")
            if date_col_maint and date_col_maint in df_filtered.columns:
                try:
                    df_filtered = df_filtered.sort_values(by=date_col_maint, ascending=True)
                except:
                    pass
            st.success(f"Found {len(df_filtered)} maintenance tasks")
            st.dataframe(df_filtered, use_container_width=True)
            excel_file = export_filtered_results_to_excel(df_filtered, "Preventive_Maintenance")
            st.download_button("📥 Download results", excel_file, f"maintenance_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("No results")

# ===================== Delete Linked Maintenance Helpers =====================
def delete_maintenance_tasks_for_equipment(equipment_name, sheets_edit):
    if APP_CONFIG["MAINTENANCE_SHEET"] in sheets_edit:
        df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
        if not df.empty and "Equipment" in df.columns:
            df = df[df["Equipment"] != equipment_name]
            sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = df
    return sheets_edit

def delete_maintenance_tasks_for_section(section_name, sheets_edit):
    if section_name not in sheets_edit:
        return sheets_edit
    df_section = sheets_edit[section_name]
    if "Equipment" not in df_section.columns:
        return sheets_edit
    equipment_list = get_equipment_list_from_sheet(df_section)
    if not equipment_list:
        return sheets_edit
    if APP_CONFIG["MAINTENANCE_SHEET"] in sheets_edit:
        df_maintenance = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
        if not df_maintenance.empty and "Equipment" in df_maintenance.columns:
            df_maintenance = df_maintenance[~df_maintenance["Equipment"].isin(equipment_list)]
            sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = df_maintenance
    return sheets_edit
# =====================================================================

# ------------------------------- Equipment & Sections Management -------------------------------
def load_equipment_config():
    if not os.path.exists(EQUIPMENT_CONFIG_FILE):
        with open(EQUIPMENT_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, indent=4, ensure_ascii=False)
        return {}
    try:
        with open(EQUIPMENT_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def save_equipment_config(config):
    try:
        with open(EQUIPMENT_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        st.error(f"Error saving equipment config: {e}")
        return False

def get_equipment_list_from_sheet(df):
    if df is None or df.empty or "Equipment" not in df.columns:
        return []
    equipment = df["Equipment"].dropna().unique()
    equipment = [str(e).strip() for e in equipment if str(e).strip() != ""]
    return sorted(equipment)

def get_available_sections(sheets_edit):
    sections = []
    for sheet_name, df in sheets_edit.items():
        if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
            continue
        if "Equipment" in df.columns and not df["Equipment"].dropna().empty:
            sections.append(sheet_name)
    return sections

def add_equipment_to_sheet_data(sheets_edit, sheet_name, new_equipment):
    if sheet_name not in sheets_edit:
        return False, "Section not found"
    df = sheets_edit[sheet_name]
    if "Equipment" not in df.columns:
        return False, "'Equipment' column not found"
    existing = get_equipment_list_from_sheet(df)
    if new_equipment in existing:
        return False, f"Machine '{new_equipment}' already exists in this section"
    new_row = {col: "" for col in df.columns}
    new_row["Equipment"] = new_equipment
    sheets_edit[sheet_name] = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    return True, f"Machine '{new_equipment}' added to section {sheet_name}"

def remove_equipment_from_sheet_data(sheets_edit, sheet_name, equipment_name):
    if sheet_name not in sheets_edit:
        return False, "Section not found"
    df = sheets_edit[sheet_name]
    if "Equipment" not in df.columns:
        return False, "'Equipment' column not found"
    if equipment_name not in get_equipment_list_from_sheet(df):
        return False, "Machine not found"
    sheets_edit[sheet_name] = df[df["Equipment"] != equipment_name]
    sheets_edit = delete_maintenance_tasks_for_equipment(equipment_name, sheets_edit)
    return True, f"All records of machine '{equipment_name}' and its maintenance tasks were deleted"

def add_new_department(sheets_edit):
    if st.session_state.get("username") == "admin":
        st.subheader("➕ Add New Section")
        st.info("A new sheet will be created for this section")
        col1, col2 = st.columns(2)
        with col1:
            new_department_name = st.text_input("📝 Section name:", key="new_department_name", placeholder="e.g. Mechanics, Electrical, Water Station")
            if new_department_name and new_department_name in sheets_edit:
                st.error(f"Section '{new_department_name}' already exists!")
            elif new_department_name:
                st.success(f"Section name '{new_department_name}' is available")
        with col2:
            st.markdown("#### Column Settings")
            use_default = st.checkbox("Use default columns", value=True, key="use_default_columns")
            if use_default:
                columns_list = APP_CONFIG["DEFAULT_SHEET_COLUMNS"]
                st.info(f"Columns: {', '.join(columns_list)}")
            else:
                columns_text = st.text_area("Custom columns (one per line):", value="\n".join(APP_CONFIG["DEFAULT_SHEET_COLUMNS"]), key="custom_columns", height=150)
                columns_list = [col.strip() for col in columns_text.split("\n") if col.strip()]
                if not columns_list:
                    columns_list = APP_CONFIG["DEFAULT_SHEET_COLUMNS"]
        st.markdown("---")
        st.markdown("### Preview")
        st.dataframe(pd.DataFrame(columns=columns_list), use_container_width=True)
        st.caption(f"Columns: {len(columns_list)} | Empty section will be created")
        if st.button("✅ Create Section", key="create_department_btn", type="primary", use_container_width=True):
            if not new_department_name:
                st.error("Please enter a section name")
                return sheets_edit
            clean_name = re.sub(r'[\\/*?:"<>|]', '_', new_department_name.strip())
            if clean_name != new_department_name:
                st.warning(f"Section name adjusted to: {clean_name}")
                new_department_name = clean_name
            if new_department_name in sheets_edit:
                st.error(f"Section '{new_department_name}' already exists!")
                return sheets_edit
            sheets_edit[new_department_name] = pd.DataFrame(columns=columns_list)
            if save_and_push_to_github(sheets_edit, f"Create section: {new_department_name}"):
                st.success(f"Section '{new_department_name}' created!")
                st.cache_data.clear()
                st.balloons()
                st.rerun()
            else:
                st.error("Failed to save section")
                return sheets_edit

        st.markdown("---")
        st.subheader("🗑️ Delete Existing Section")
        st.warning("Deleting a section permanently removes all its data (machines, faults, spare parts, maintenance tasks).")
        deletable_sections = [name for name in sheets_edit.keys() if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]
        if not deletable_sections:
            st.info("No sections to delete.")
        else:
            selected_dept = st.selectbox("Select section to delete:", deletable_sections, key="delete_department_select")
            if selected_dept:
                st.error(f"You are about to delete section **'{selected_dept}'** permanently.")
                confirm = st.text_input("Type section name to confirm:", key="delete_confirm")
                if confirm == selected_dept:
                    if st.button("🗑️ Delete Section", key="delete_department_btn", type="primary"):
                        sheets_edit = delete_maintenance_tasks_for_section(selected_dept, sheets_edit)
                        spare_df = load_spare_parts()
                        if not spare_df.empty:
                            spare_df = spare_df[spare_df["Section"] != selected_dept]
                            sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = spare_df
                            st.info(f"Deleted spare parts for section '{selected_dept}'.")
                        del sheets_edit[selected_dept]
                        if save_and_push_to_github(sheets_edit, f"Delete section: {selected_dept}"):
                            log_activity("delete_section", f"Deleted section '{selected_dept}' and all related data", section=selected_dept)
                            st.success(f"Section '{selected_dept}' deleted!")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Failed to save changes after deleting section.")
                elif confirm:
                    st.warning("Name does not match. Section not deleted.")
    else:
        st.info("Only admin can add or delete sections.")
    st.markdown("---")
    st.markdown("### Current Sections:")
    if sheets_edit:
        for dept_name in sheets_edit.keys():
            if dept_name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
                st.write(f"- 🏭 {dept_name}")
    else:
        st.info("No sections yet")
    return sheets_edit

def add_new_machine(sheets_edit, sheet_name):
    st.markdown(f"### 🔧 Add New Machine to Section: {sheet_name}")
    df = sheets_edit[sheet_name]
    equipment_list = get_equipment_list_from_sheet(df)
    st.markdown(f"**Machines in this section:**")
    if equipment_list:
        for eq in equipment_list:
            st.markdown(f"- 🔹 {eq}")
    else:
        st.info("No machines registered yet")
    st.markdown("---")
    new_machine = st.text_input("📝 New machine name:", key=f"new_machine_{sheet_name}")
    if st.button("➕ Add Machine", key=f"add_machine_{sheet_name}", type="primary"):
        if new_machine:
            success, msg = add_equipment_to_sheet_data(sheets_edit, sheet_name, new_machine)
            if success:
                if save_and_push_to_github(sheets_edit, f"Add machine: {new_machine} to {sheet_name}"):
                    st.success(msg)
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("Save failed")
            else:
                st.error(msg)
        else:
            st.warning("Please enter machine name")
    return sheets_edit

def manage_machines(sheets_edit, sheet_name, unique_suffix=""):
    st.markdown(f"### 🔧 Manage Machines in Section: {sheet_name}")
    df = sheets_edit[sheet_name]
    equipment_list = get_equipment_list_from_sheet(df)
    if equipment_list:
        st.markdown("#### Machine list:")
        for eq in equipment_list:
            st.markdown(f"- 🔹 {eq}")
    else:
        st.info("No machines registered")
    st.markdown("---")

    with st.form(key=f"add_machine_form_{sheet_name}_{unique_suffix}"):
        new_machine = st.text_input("➕ New machine name:", key=f"new_machine_input_{sheet_name}_{unique_suffix}")
        submitted_add = st.form_submit_button("➕ Add Machine")
        if submitted_add:
            if new_machine:
                success, msg = add_equipment_to_sheet_data(sheets_edit, sheet_name, new_machine)
                if success:
                    if save_and_push_to_github(sheets_edit, f"Add machine: {new_machine} to {sheet_name}"):
                        st.success(msg)
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("Save failed")
                else:
                    st.error(msg)
            else:
                st.warning("Please enter machine name")

    if equipment_list:
        st.markdown("#### Delete Machine")
        if st.session_state.get("username") == "admin":
            with st.form(key=f"delete_machine_form_{sheet_name}_{unique_suffix}"):
                machine_to_delete = st.selectbox("Select machine to delete:", equipment_list, key=f"delete_machine_select_{sheet_name}_{unique_suffix}")
                st.warning("Warning: deleting a machine removes all its fault records and maintenance tasks permanently!")
                submitted_del = st.form_submit_button("🗑️ Delete Machine")
                if submitted_del:
                    success, msg = remove_equipment_from_sheet_data(sheets_edit, sheet_name, machine_to_delete)
                    if success:
                        if save_and_push_to_github(sheets_edit, f"Delete machine: {machine_to_delete} from {sheet_name}"):
                            st.success(msg)
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Save failed")
                    else:
                        st.error(msg)
        else:
            st.info("Only admin can delete machines.")
    else:
        st.info("No machines to delete")

def add_new_event(sheets_edit, sheet_name):
    st.markdown(f"### 📝 Add New Event in Section: {sheet_name}")
    df = sheets_edit[sheet_name]
    equipment_list = get_equipment_list_from_sheet(df)
    if not equipment_list:
        st.warning("No machines registered in this section. Please add a machine first.")
        return sheets_edit

    selected_equipment = st.selectbox("🔧 Select Machine:", equipment_list, key="equipment_select")

    if "last_selected_equipment" not in st.session_state:
        st.session_state.last_selected_equipment = selected_equipment
    if st.session_state.last_selected_equipment != selected_equipment:
        st.session_state.last_selected_equipment = selected_equipment
        if "event_desc_area" in st.session_state:
            del st.session_state.event_desc_area
        if "correction_desc_area" in st.session_state:
            del st.session_state.correction_desc_area

    df_equip = df[df["Equipment"] == selected_equipment]
    previous_events = df_equip["Event_Fault"].dropna().unique()
    previous_events = [str(e).strip() for e in previous_events if str(e).strip() != ""]

    if not previous_events:
        st.info("No previous faults for this machine. You can enter a new event.")

    event_options = ["-- Select from previous --"] + sorted(previous_events)
    selected_event_option = st.selectbox("Select previous event:", event_options, key="event_old_select")

    if selected_event_option != "-- Select from previous --":
        st.session_state.event_desc_area = selected_event_option

    previous_corrections = df_equip["Corrective_Action"].dropna().unique()
    previous_corrections = [str(c).strip() for c in previous_corrections if str(c).strip() != ""]

    if not previous_corrections:
        st.info("No previous corrective actions for this machine.")

    correction_options = ["-- Select from previous --"] + sorted(previous_corrections)
    selected_correction_option = st.selectbox("Select previous corrective action:", correction_options, key="correction_old_select")

    if selected_correction_option != "-- Select from previous --":
        st.session_state.correction_desc_area = selected_correction_option

    part_name = ""
    consume_qty = 0
    warning_msg = ""

    spare_parts_list = get_spare_parts_for_section(sheet_name)

    with st.form(key="add_event_form"):
        col1, col2 = st.columns(2)
        with col1:
            event_date = st.date_input("📅 Date:", value=datetime.now())
            repair_duration = st.number_input("⏱️ Repair duration (hours):", min_value=0.0, step=0.5, format="%.1f")
            st.text_area("📝 Event/Fault:", height=100, key="event_desc_area")
            fault_type = st.selectbox("🏷️ Fault type:", ["Mechanical", "Electrical", "Electronic", "Hydraulic", "Service", "Maintenance", "Other"])
            uploaded_image = st.file_uploader("🖼️ Upload image (optional):", type=APP_CONFIG["ALLOWED_IMAGE_TYPES"])
        with col2:
            st.text_area("🔧 Corrective action:", height=100, key="correction_desc_area")
            servised_by = st.text_input("👨‍🔧 Performed by:")
            technician_rating = st.select_slider("⭐ Technician skill (solve/think/initiative/decision):", options=[1, 2, 3, 4, 5], value=3)
            safety_compliance = st.selectbox("🛡️ Safety compliance:", ["", "Fully_Compliant", "Partially_Compliant", "Not_Compliant", "Not_Applicable"])
            st.markdown("---")
            st.markdown("**🔩 Spare Parts Used**")
            if spare_parts_list:
                part_names = [f"{name} (Qty: {qty})" for name, qty in spare_parts_list]
                selected_part_display = st.selectbox("Select part:", [""] + part_names, key="spare_part_select")
                if selected_part_display:
                    part_name = selected_part_display.split(" (")[0]
                    current_qty = next((qty for name, qty in spare_parts_list if name == part_name), 0)
                    st.caption(f"Current stock: {current_qty}")
                    consume_qty = st.number_input("Quantity used:", min_value=1, max_value=max(1, current_qty), value=1, step=1, key="consume_qty")
                    if consume_qty > current_qty:
                        st.error(f"Insufficient stock (Available {current_qty})")
                    else:
                        st.success(f"{consume_qty} will be deducted")
                else:
                    part_name = ""
                    consume_qty = 0
            else:
                st.info("No spare parts registered for this section.")
                part_name = ""
                consume_qty = 0

        submitted = st.form_submit_button("✅ Add Event", type="primary")
        if submitted:
            event_desc = st.session_state.get("event_desc_area", "")
            correction_desc = st.session_state.get("correction_desc_area", "")

            spare_part_used = ""
            if part_name and consume_qty > 0:
                success, msg, new_qty = consume_spare_part(part_name, consume_qty)
                if success:
                    spare_part_used = f"{part_name} (Qty {consume_qty})"
                    critical_parts = get_critical_spare_parts()
                    for cp in critical_parts:
                        if cp["Part_Name"] == part_name:
                            warning_msg = f"Warning: Part '{part_name}' is critical and its stock is now {new_qty}. Please reorder."
                            break
                else:
                    st.error(msg)
                    return sheets_edit

            image_url = None
            if uploaded_image is not None:
                event_id = str(uuid.uuid4())[:8]
                image_url = upload_image_to_github(uploaded_image, "event", event_id)
                if image_url:
                    st.success("Image uploaded successfully!")
                else:
                    st.warning("Image upload failed, event saved without image")

            new_row = {
                "Repair_Duration": repair_duration if repair_duration > 0 else "",
                "Date": event_date.strftime("%Y-%m-%d"),
                "Equipment": selected_equipment,
                "Event_Fault": event_desc,
                "Corrective_Action": correction_desc,
                "Performed_By": servised_by,
                "Spare_Parts_Used": spare_part_used,
                "Fault_Type": fault_type if fault_type else "",
                "Technician_Skill": technician_rating,
                "Safety_Compliance": safety_compliance if safety_compliance else "",
                "Image_URL": image_url or ""
            }
            for col in df.columns:
                if col not in new_row:
                    new_row[col] = ""

            if is_duplicate_event(
                df,
                new_row,
                compare_columns=["Date", "Equipment", "Event_Fault", "Corrective_Action", "Performed_By"],
                ignore_columns=[],
                time_window_days=1
            ):
                st.warning("This event is already registered for the same machine, action, and technician within the last day.")
                return sheets_edit

            new_row_df = pd.DataFrame([new_row])
            sheets_edit[sheet_name] = pd.concat([df, new_row_df], ignore_index=True)

            if "temp_spare_parts_df" in st.session_state:
                sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = st.session_state.temp_spare_parts_df
                del st.session_state.temp_spare_parts_df

            commit_message = f"Add event with spare part {part_name}" if part_name else "Add event without spare part"

            if save_and_push_to_github(sheets_edit, commit_message):
                st.cache_data.clear()
                log_activity("add_event", f"Added fault: {event_desc[:50]} for machine {selected_equipment}", section=sheet_name)
                if "event_desc_area" in st.session_state:
                    del st.session_state.event_desc_area
                if "correction_desc_area" in st.session_state:
                    del st.session_state.correction_desc_area
                st.success("Event added and pushed to GitHub!")
                if warning_msg:
                    st.warning(warning_msg)

                try:
                    subject = f"New Fault Event - {selected_equipment} in {sheet_name}"
                    body = f"""A new fault event has been added to CMMS:

Date: {event_date.strftime('%Y-%m-%d')}
Section: {sheet_name}
Equipment: {selected_equipment}
Fault: {event_desc}
Corrective Action: {correction_desc}
Performed By: {servised_by}
Spare Parts Used: {spare_part_used if spare_part_used else 'None'}

--- Current Notifications ---
{get_current_notifications_text()}
                    """
                    send_email(subject, body)
                except Exception as e:
                    st.warning(f"Could not send email: {e}")

                st.rerun()
            else:
                st.error("Save failed")
    return sheets_edit

# ------------------------------- Preventive Maintenance Helpers -------------------------------
def execute_maintenance_with_date(sheets_edit, equipment_name, task_name, execution_date, performed_by, used_spare_part="", used_quantity=1, image_url=None, section=None):
    if APP_CONFIG["MAINTENANCE_SHEET"] not in sheets_edit:
        sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = pd.DataFrame(columns=APP_CONFIG["MAINTENANCE_COLUMNS"])
    df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
    if df.empty:
        return False, "No maintenance tasks"
    mask = (df["Equipment"] == equipment_name) & (df["Task_Name"] == task_name)
    if not mask.any():
        return False, f"Task '{task_name}' not found for equipment '{equipment_name}'"
    idx = df[mask].index[0]
    period_days = df.loc[idx, "Period_Days"]

    last_exec = df.loc[idx, "Last_Execution"]
    if pd.notna(last_exec) and hasattr(last_exec, 'date'):
        if last_exec.date() == execution_date:
            return False, f"Maintenance '{task_name}' for '{equipment_name}' already executed on this date."

    df.loc[idx, "Last_Execution"] = pd.to_datetime(execution_date)
    next_date = execution_date + timedelta(days=period_days)
    df.loc[idx, "Next_Date"] = next_date
    old_notes = df.loc[idx, "Notes"] if pd.notna(df.loc[idx, "Notes"]) else ""
    new_entry = f"{execution_date.strftime('%Y-%m-%d')} | Performed by: {performed_by}"
    warning_msg = ""
    if used_spare_part and used_quantity > 0:
        success, msg, new_qty = consume_spare_part(used_spare_part, used_quantity)
        if not success:
            return False, f"Spare part deduction failed: {msg}"
        new_entry += f" | Used {used_spare_part} qty {used_quantity} - {msg}"
        critical_parts = get_critical_spare_parts()
        for cp in critical_parts:
            if cp["Part_Name"] == used_spare_part:
                warning_msg = f"Warning: Part '{used_spare_part}' is critical and stock is now {new_qty}."
                break
    if image_url:
        new_entry += f" | Image: {image_url}"
    df.loc[idx, "Notes"] = (old_notes + "\n" + new_entry).strip()
    sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = df

    if section is None:
        for sheet_name, sh_df in sheets_edit.items():
            if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
                continue
            if equipment_name in get_equipment_list_from_sheet(sh_df):
                section = sheet_name
                break
        if section is None:
            section = "Unknown"

    log_activity("execute_maintenance", f"Executed maintenance '{task_name}' for {equipment_name} by {performed_by}", section=section)
    result_msg = f"Maintenance '{task_name}' executed on {execution_date.strftime('%Y-%m-%d')} by {performed_by}. Next date: {next_date.strftime('%Y-%m-%d')}" + (f" {warning_msg}" if warning_msg else "")

    try:
        subject = f"Maintenance Executed - {equipment_name} - {task_name}"
        body = f"""Preventive maintenance executed in CMMS:

Equipment: {equipment_name}
Task: {task_name}
Execution Date: {execution_date.strftime('%Y-%m-%d')}
Performed By: {performed_by}
Spare Parts Used: {used_spare_part if used_spare_part else 'None'}
Next Date: {next_date.strftime('%Y-%m-%d')}

--- Current Notifications ---
{get_current_notifications_text()}
        """
        send_email(subject, body)
    except Exception as e:
        st.warning(f"Could not send email: {e}")

    return True, result_msg

def add_maintenance_as_event(sheets_edit, equipment_name, task_name, execution_date, performed_by, used_spare_part="", used_quantity=1, image_url=None):
    target_sheet = None
    target_df = None
    for sheet_name, df in sheets_edit.items():
        if sheet_name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
            if equipment_name in get_equipment_list_from_sheet(df):
                target_sheet = sheet_name
                target_df = df
                break
    if target_sheet is None:
        return False, f"No section contains equipment '{equipment_name}'"
    spare_part_used = f"{used_spare_part} (Qty {used_quantity})" if used_spare_part else ""
    new_row = {
        "Repair_Duration": 0, "Date": execution_date.strftime("%Y-%m-%d"), "Equipment": equipment_name,
        "Event_Fault": f"Preventive Maintenance: {task_name}",
        "Corrective_Action": f"Executed periodic maintenance '{task_name}' by {performed_by}",
        "Performed_By": performed_by, "Spare_Parts_Used": spare_part_used, "Fault_Type": "Preventive_Maintenance",
        "Technician_Skill": 5, "Safety_Compliance": "Fully_Compliant", "Image_URL": image_url or ""
    }
    for col in target_df.columns:
        if col not in new_row:
            new_row[col] = ""
    sheets_edit[target_sheet] = pd.concat([target_df, pd.DataFrame([new_row])], ignore_index=True)
    return True, f"Maintenance logged as event in section '{target_sheet}' by {performed_by}"

# ------------------------------- Spare Parts Tab -------------------------------
def manage_spare_parts_tab(sheets_edit):
    st.header("📦 Spare Parts Management")
    st.info("Add/edit spare parts per section. Parts in 'General' section are available to all sections.")
    username = st.session_state.get("username")
    all_sheets = load_all_sheets()
    real_sections = get_allowed_sections(all_sheets, username, "view")
    allowed_sections = real_sections.copy()
    if real_sections or username == "admin":
        if APP_CONFIG["GENERAL_SECTION"] not in allowed_sections:
            allowed_sections = [APP_CONFIG["GENERAL_SECTION"]] + allowed_sections
    else:
        st.warning("No sections available for you.")
        return sheets_edit

    selected_section = st.selectbox("🏭 Select Section:", allowed_sections, key="spare_section")
    spare_df = load_spare_parts()
    view_mode = st.radio("View mode:", ["Table", "Cards with images"], horizontal=True, key="spare_view_mode")
    st.subheader("📋 Spare Parts List")
    filtered_df = spare_df[spare_df["Section"] == selected_section].copy()
    filtered_df.reset_index(drop=False, inplace=True)
    filtered_df.rename(columns={'index': 'original_index'}, inplace=True)
    filtered_df["id"] = filtered_df.index

    if filtered_df.empty:
        st.info(f"No spare parts registered for section '{selected_section}'.")
    else:
        part_name_filter = st.text_input("Filter by part name:", placeholder="Type part of name...", key="spare_name_filter")
        if part_name_filter:
            filtered_df = filtered_df[filtered_df["Part_Name"].str.contains(part_name_filter, case=False, na=False)]

        if view_mode == "Table":
            display_cols = [c for c in filtered_df.columns if c not in ["original_index", "id", "Image_URL"]]
            st.dataframe(filtered_df[display_cols], use_container_width=True)
            st.markdown("#### Edit/Delete Part")
            part_options = filtered_df["Part_Name"].tolist()
            selected_part_name = st.selectbox("Select part:", part_options, key="edit_part_name_select")
            if selected_part_name:
                part_row = filtered_df[filtered_df["Part_Name"] == selected_part_name].iloc[0]
                with st.expander(f"Edit part: {selected_part_name}", expanded=True):
                    new_name = st.text_input("Part name", value=part_row["Part_Name"], key="edit_name")
                    new_size = st.text_input("Size", value=part_row["Size"], key="edit_size")
                    new_qty = st.number_input("Stock", value=int(part_row["Available_Quantity"]), step=1, key="edit_qty")
                    new_lead = st.text_input("Lead time", value=part_row["Lead_Time"], key="edit_lead")
                    new_critical = st.checkbox("Critical part", value=(part_row["Is_Critical"] == "Yes"), key="edit_critical")
                    new_threshold = st.number_input("Alert threshold", value=int(part_row.get("Alert_Threshold", 1)), step=1, key="edit_threshold")
                    if st.button("💾 Save Changes", key="save_edit_part"):
                        original_idx = part_row["original_index"]
                        spare_df.loc[original_idx, "Part_Name"] = new_name
                        spare_df.loc[original_idx, "Size"] = new_size
                        spare_df.loc[original_idx, "Available_Quantity"] = new_qty
                        spare_df.loc[original_idx, "Lead_Time"] = new_lead
                        spare_df.loc[original_idx, "Is_Critical"] = "Yes" if new_critical else "No"
                        spare_df.loc[original_idx, "Alert_Threshold"] = new_threshold
                        sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = spare_df
                        if save_and_push_to_github(sheets_edit, f"Edit part: {selected_part_name}"):
                            log_activity("add_spare_part", f"Edited spare part '{selected_part_name}' in section {selected_section}", section=selected_section)
                            st.success("Updated")
                            st.rerun()
                if st.button("🗑️ Delete Part", key="delete_part_btn"):
                    original_idx = part_row["original_index"]
                    spare_df = spare_df.drop(index=original_idx)
                    sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = spare_df
                    if save_and_push_to_github(sheets_edit, f"Delete part: {selected_part_name}"):
                        st.success("Deleted")
                        st.rerun()
        else:
            cols_per_row = 2
            for i in range(0, len(filtered_df), cols_per_row):
                row_cols = st.columns(cols_per_row)
                for j, col in enumerate(row_cols):
                    idx = i + j
                    if idx < len(filtered_df):
                        row = filtered_df.iloc[idx]
                        with col:
                            with st.container(border=True):
                                img_url = row.get("Image_URL", "")
                                if img_url and isinstance(img_url, str) and img_url.strip():
                                    try:
                                        st.image(img_url, use_container_width=True)
                                    except:
                                        st.write("Image unavailable")
                                else:
                                    st.write("No image")
                                st.markdown(f"**🔩 {row['Part_Name']}**")
                                st.markdown(f"**Size:** {row['Size']}")
                                st.markdown(f"**Stock:** {row['Available_Quantity']}")
                                st.markdown(f"**Critical:** {row['Is_Critical']}")
                                if row.get('Lead_Time'):
                                    st.markdown(f"**Lead time:** {row['Lead_Time']}")
                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1:
                                    if st.button("✏️ Edit", key=f"edit_card_{row['id']}"):
                                        st.session_state[f"edit_mode_{row['id']}"] = True
                                with col_btn2:
                                    if st.button("🗑️ Delete", key=f"delete_card_{row['id']}"):
                                        original_idx = row["original_index"]
                                        spare_df = spare_df.drop(index=original_idx)
                                        sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = spare_df
                                        if save_and_push_to_github(sheets_edit, f"Delete part: {row['Part_Name']}"):
                                            st.success("Deleted")
                                            st.rerun()
                                if st.session_state.get(f"edit_mode_{row['id']}", False):
                                    with st.form(key=f"edit_form_{row['id']}"):
                                        new_name = st.text_input("Part name", value=row['Part_Name'])
                                        new_size = st.text_input("Size", value=row['Size'])
                                        new_qty = st.number_input("Stock", value=int(row['Available_Quantity']))
                                        new_lead = st.text_input("Lead time", value=row['Lead_Time'])
                                        new_critical = st.checkbox("Critical", value=(row['Is_Critical'] == "Yes"))
                                        if st.form_submit_button("💾 Save"):
                                            original_idx = row["original_index"]
                                            spare_df.loc[original_idx, "Part_Name"] = new_name
                                            spare_df.loc[original_idx, "Size"] = new_size
                                            spare_df.loc[original_idx, "Available_Quantity"] = new_qty
                                            spare_df.loc[original_idx, "Lead_Time"] = new_lead
                                            spare_df.loc[original_idx, "Is_Critical"] = "Yes" if new_critical else "No"
                                            sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = spare_df
                                            if save_and_push_to_github(sheets_edit, f"Edit part: {row['Part_Name']}"):
                                                st.success("Updated")
                                                del st.session_state[f"edit_mode_{row['id']}"]
                                                st.rerun()
                                            else:
                                                st.error("Save failed")

    st.subheader("➕ Add New Spare Part")
    with st.form(key="add_spare_part_form"):
        col1, col2 = st.columns(2)
        with col1:
            part_name = st.text_input("🔩 Part name:")
            part_size = st.text_input("📏 Size:")
            part_image = st.file_uploader("🖼️ Part image (optional):", type=APP_CONFIG["ALLOWED_IMAGE_TYPES"], key="spare_part_image")
        with col2:
            initial_qty = st.number_input("📦 Initial stock:", min_value=0, step=1, value=0)
            lead_time = st.text_input("⏱️ Lead time (days or text):")
            is_critical = st.checkbox("⚠️ Critical part (will appear in notifications on shortage)")
            critical_threshold = st.number_input("⚠️ Alert threshold (when stock falls below):", min_value=1, step=1, value=1)
        submitted = st.form_submit_button("✅ Add Part")
        if submitted:
            if not part_name:
                st.error("Please enter part name")
            else:
                existing = spare_df[(spare_df["Part_Name"] == part_name) & (spare_df["Section"] == selected_section)]
                if not existing.empty:
                    st.error(f"Part '{part_name}' already exists in '{selected_section}'")
                else:
                    image_url = None
                    if part_image is not None:
                        part_id = str(uuid.uuid4())[:8]
                        image_url = upload_image_to_github(part_image, "spare_part", part_id)
                        if image_url:
                            st.success("Image uploaded")
                        else:
                            st.warning("Image upload failed")
                    new_row = pd.DataFrame([{
                        "Part_Name": part_name,
                        "Size": part_size,
                        "Available_Quantity": initial_qty,
                        "Lead_Time": lead_time,
                        "Is_Critical": "Yes" if is_critical else "No",
                        "Section": selected_section,
                        "Image_URL": image_url or "",
                        "Alert_Threshold": critical_threshold
                    }])
                    sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = pd.concat([spare_df, new_row], ignore_index=True)
                    if save_and_push_to_github(sheets_edit, f"Add spare part: {part_name} to {selected_section}"):
                        log_activity("add_spare_part", f"Added spare part '{part_name}' to {selected_section} (Stock: {initial_qty})", section=selected_section)
                        st.success("Spare part added")
                        st.rerun()
                    else:
                        st.error("Save failed")
    return sheets_edit

# ------------------------------- Preventive Maintenance Tab -------------------------------
def preventive_maintenance_tab(sheets_edit):
    st.header("🛠 Preventive Maintenance")
    st.info("Manage periodic maintenance tasks. Data is saved automatically.")
    username = st.session_state.get("username")
    all_sheets = load_all_sheets()
    allowed_sections = get_allowed_sections(all_sheets, username, "view")
    if not allowed_sections:
        st.warning("No sections available.")
        return sheets_edit
    selected_section = st.selectbox("🏭 Select Section:", allowed_sections, key="pm_section")
    df_section = sheets_edit[selected_section]
    equipment_list = get_equipment_list_from_sheet(df_section)
    if not equipment_list:
        st.warning(f"No machines in section '{selected_section}'.")
        return sheets_edit
    selected_equipment = st.selectbox("🔧 Select Equipment:", equipment_list, key="pm_equipment")
    if APP_CONFIG["MAINTENANCE_SHEET"] in sheets_edit:
        tasks_df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]].copy()
    else:
        tasks_df = load_maintenance_tasks()
        sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = tasks_df
    tasks_df = tasks_df[tasks_df["Equipment"] == selected_equipment].copy()
    st.subheader(f"📋 Maintenance Tasks for {selected_equipment}")
    if tasks_df.empty:
        st.info("No maintenance tasks for this equipment. You can add a new task below.")
    else:
        view_mode = st.radio("View mode:", ["Table", "Cards with images"], horizontal=True, key="maintenance_view_mode")
        today = datetime.now().date()
        tasks_display = tasks_df.copy()
        tasks_display.reset_index(drop=False, inplace=True)
        tasks_display.rename(columns={"index": "original_index"}, inplace=True)
        def days_remaining(row):
            if pd.isna(row["Next_Date"]):
                return "Not Set"
            return (row["Next_Date"].date() - today).days
        tasks_display["Days_Remaining"] = tasks_display.apply(days_remaining, axis=1)
        tasks_display["Status"] = tasks_display["Days_Remaining"].apply(lambda x: "🔴 Overdue" if (isinstance(x, int) and x < 0) else ("🟡 Upcoming" if (isinstance(x, int) and x <= 3) else "🟢 OK"))
        tasks_display["Executions_Count"] = tasks_display["Last_Execution"].apply(lambda x: 1 if pd.notna(x) else 0)
        if view_mode == "Table":
            cols_to_show = ["Maintenance_Type", "Task_Name", "Period_Days", "Last_Execution", "Next_Date", "Days_Remaining", "Status", "Executions_Count", "Notes"]
            st.dataframe(tasks_display[cols_to_show], use_container_width=True)
            st.markdown("#### Edit/Delete Task")
            task_options = tasks_display["Task_Name"].tolist()
            selected_task_name = st.selectbox("Select task:", task_options, key="edit_task_select")
            if selected_task_name:
                task_row = tasks_display[tasks_display["Task_Name"] == selected_task_name].iloc[0]
                original_idx = task_row["original_index"]
                with st.expander(f"Edit task: {selected_task_name}", expanded=True):
                    new_name = st.text_input("Task name", value=task_row["Task_Name"], key="edit_task_name")
                    new_period_hours = st.number_input("Hours between maintenance", min_value=1, value=int(task_row["Period_Days"]*24), key="edit_period_hours")
                    new_notes = st.text_area("Notes", value=task_row["Notes"] if pd.notna(task_row["Notes"]) else "", key="edit_task_notes")
                    if st.button("💾 Save Changes", key="save_task_edit"):
                        new_period_days = new_period_hours / 24.0
                        main_df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
                        main_df.loc[original_idx, "Task_Name"] = new_name
                        main_df.loc[original_idx, "Period_Days"] = new_period_days
                        main_df.loc[original_idx, "Maintenance_Type"] = f"{new_period_hours} hours"
                        main_df.loc[original_idx, "Notes"] = new_notes
                        last_exec = main_df.loc[original_idx, "Last_Execution"]
                        if pd.notna(last_exec) and hasattr(last_exec, 'date'):
                            main_df.loc[original_idx, "Next_Date"] = last_exec + timedelta(days=new_period_days)
                        else:
                            main_df.loc[original_idx, "Next_Date"] = datetime.now().date() + timedelta(days=new_period_days)
                        sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = main_df
                        if save_and_push_to_github(sheets_edit, f"Edit maintenance task: {selected_task_name}"):
                            st.success("Updated")
                            st.rerun()
                if st.button("🗑️ Delete Task", key="delete_task_btn"):
                    main_df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
                    main_df = main_df.drop(index=original_idx)
                    sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = main_df
                    if save_and_push_to_github(sheets_edit, f"Delete maintenance task: {selected_task_name}"):
                        st.success("Deleted")
                        st.rerun()
        else:
            cols_per_row = 2
            for i in range(0, len(tasks_display), cols_per_row):
                row_cols = st.columns(cols_per_row)
                for j, col in enumerate(row_cols):
                    idx = i + j
                    if idx < len(tasks_display):
                        row = tasks_display.iloc[idx]
                        original_idx = row["original_index"]
                        with col:
                            with st.container(border=True):
                                img_url = row.get("Image_URL", "")
                                if img_url and isinstance(img_url, str) and img_url.strip():
                                    try:
                                        st.image(img_url, use_container_width=True)
                                    except:
                                        st.write("Image unavailable")
                                else:
                                    st.write("No image")
                                st.markdown(f"**{row['Task_Name']}**")
                                st.markdown(f"**Type:** {row['Maintenance_Type']}")
                                st.markdown(f"**Period:** {row['Period_Days']:.2f} days")
                                last_exec_val = row['Last_Execution']
                                if pd.notna(last_exec_val) and hasattr(last_exec_val, 'strftime'):
                                    last_exec_str = last_exec_val.strftime('%Y-%m-%d')
                                else:
                                    last_exec_str = 'Not executed yet'
                                st.markdown(f"**Last execution:** {last_exec_str}")
                                next_date_val = row['Next_Date']
                                if pd.notna(next_date_val) and hasattr(next_date_val, 'strftime'):
                                    next_date_str = next_date_val.strftime('%Y-%m-%d')
                                else:
                                    next_date_str = 'Not Set'
                                st.markdown(f"**Next date:** {next_date_str}")
                                st.markdown(f"**Status:** {row['Status']}")
                                col_btn1, col_btn2 = st.columns(2)
                                with col_btn1:
                                    if st.button("✏️ Edit", key=f"edit_task_card_{original_idx}"):
                                        st.session_state[f"edit_task_mode_{original_idx}"] = True
                                with col_btn2:
                                    if st.button("🗑️ Delete", key=f"delete_task_card_{original_idx}"):
                                        main_df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
                                        main_df = main_df.drop(index=original_idx)
                                        sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = main_df
                                        if save_and_push_to_github(sheets_edit, f"Delete maintenance task: {row['Task_Name']}"):
                                            st.success("Deleted")
                                            st.rerun()
                                if st.session_state.get(f"edit_task_mode_{original_idx}", False):
                                    with st.form(key=f"edit_task_form_{original_idx}"):
                                        new_name = st.text_input("Task name", value=row['Task_Name'])
                                        new_period_hours = st.number_input("Hours", min_value=1, value=int(row['Period_Days']*24))
                                        new_notes = st.text_area("Notes", value=row['Notes'] if pd.notna(row['Notes']) else "")
                                        if st.form_submit_button("💾 Save"):
                                            new_period_days = new_period_hours / 24.0
                                            main_df = sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]]
                                            main_df.loc[original_idx, "Task_Name"] = new_name
                                            main_df.loc[original_idx, "Period_Days"] = new_period_days
                                            main_df.loc[original_idx, "Maintenance_Type"] = f"{new_period_hours} hours"
                                            main_df.loc[original_idx, "Notes"] = new_notes
                                            last_exec = main_df.loc[original_idx, "Last_Execution"]
                                            if pd.notna(last_exec) and hasattr(last_exec, 'date'):
                                                main_df.loc[original_idx, "Next_Date"] = last_exec + timedelta(days=new_period_days)
                                            else:
                                                main_df.loc[original_idx, "Next_Date"] = datetime.now().date() + timedelta(days=new_period_days)
                                            sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = main_df
                                            if save_and_push_to_github(sheets_edit, f"Edit maintenance task: {row['Task_Name']}"):
                                                st.success("Updated")
                                                del st.session_state[f"edit_task_mode_{original_idx}"]
                                                st.rerun()
                                            else:
                                                st.error("Save failed")
        st.markdown("---")
        st.subheader("✅ Execute Maintenance")
        task_options = tasks_df["Task_Name"].tolist()
        if task_options:
            selected_task = st.selectbox("Select task to execute:", task_options, key="execute_task_select")
            if selected_task:
                execution_date = st.date_input("📅 Execution date:", value=datetime.now().date(), key="execution_date_input")
                performed_by = st.text_input("👨‍🔧 Performed by:", key="maintenance_performed_by")
                spare_parts_list = get_spare_parts_for_section(selected_section)
                st.markdown("**🔩 Spare parts used (optional)**")
                part_name = ""
                consume_qty = 0
                use_part = True
                if spare_parts_list:
                    part_names = [""] + [f"{name} (Qty: {qty})" for name, qty in spare_parts_list]
                    selected_part_display = st.selectbox("Select part:", part_names, key="pm_spare_part")
                    if selected_part_display:
                        part_name = selected_part_display.split(" (")[0]
                        current_qty = next((qty for name, qty in spare_parts_list if name == part_name), 0)
                        st.caption(f"Current stock: {current_qty}")
                        consume_qty = st.number_input("Quantity used:", min_value=1, max_value=max(1, current_qty), value=1, step=1, key="pm_consume_qty")
                        if consume_qty > current_qty:
                            st.error("Insufficient stock")
                            use_part = False
                else:
                    st.info("No spare parts registered for this section")
                execution_image = st.file_uploader("🖼️ Upload execution image (optional):", type=APP_CONFIG["ALLOWED_IMAGE_TYPES"], key="maintenance_execution_image")
                link_to_event = st.checkbox("🔗 Log as fault event", value=False)
                if st.button("✅ Execute Maintenance", type="primary"):
                    if not performed_by:
                        st.error("Please enter who performed the maintenance")
                    elif not use_part:
                        st.error("Cannot execute due to insufficient stock")
                    else:
                        image_url = None
                        if execution_image:
                            maint_id = str(uuid.uuid4())[:8]
                            image_url = upload_image_to_github(execution_image, "maintenance_execution", maint_id)
                        success, msg = execute_maintenance_with_date(sheets_edit, selected_equipment, selected_task, execution_date, performed_by, part_name, consume_qty, image_url, section=selected_section)
                        if success:
                            if link_to_event:
                                event_success, event_msg = add_maintenance_as_event(sheets_edit, selected_equipment, selected_task, execution_date, performed_by, part_name, consume_qty, image_url)
                                if event_success:
                                    st.success(f"{msg} - logged as fault event")
                                else:
                                    st.warning(f"{msg} but event log failed: {event_msg}")
                            else:
                                st.success(msg)
                            if "temp_spare_parts_df" in st.session_state:
                                sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = st.session_state.temp_spare_parts_df
                                del st.session_state.temp_spare_parts_df
                            if save_and_push_to_github(sheets_edit, f"Execute maintenance '{selected_task}' for {selected_equipment} by {performed_by}"):
                                st.rerun()
                        else:
                            st.error(msg)
        else:
            st.info("No tasks to execute")
    st.markdown("---")
    st.subheader("➕ Add New Maintenance Task")
    use_custom_start = st.checkbox("📅 Specify start date", key="use_custom_start_checkbox")
    start_date = None
    if use_custom_start:
        start_date = st.date_input("Start date:", value=datetime.now().date(), key="maintenance_start_date")
    with st.form(key="add_maintenance_form"):
        col1, col2 = st.columns(2)
        with col1:
            task_name = st.text_input("Task name:")
            period_hours = st.number_input("⏱️ Hours between maintenance:", min_value=1, step=1, value=24)
            st.caption(f"Period: {period_hours} hours = {period_hours/24:.2f} days")
            task_image = st.file_uploader("🖼️ Illustrative image:", type=APP_CONFIG["ALLOWED_IMAGE_TYPES"], key="maintenance_task_image")
        with col2:
            notes = st.text_area("Notes:")
            default_spare = st.text_input("Default spare part:", placeholder="Optional")
        submitted = st.form_submit_button("➕ Add Task")
        if submitted:
            if not task_name:
                st.error("Please enter task name")
            else:
                image_url = None
                if task_image:
                    task_id = str(uuid.uuid4())[:8]
                    image_url = upload_image_to_github(task_image, "maintenance_task", task_id)
                sheets_edit = add_maintenance_task(sheets_edit, selected_equipment, task_name, period_hours, start_date, notes, default_spare, image_url, section=selected_section)
                if save_and_push_to_github(sheets_edit, f"Add maintenance task '{task_name}'"):
                    st.success("Task added")
                    st.rerun()
                else:
                    st.error("Save failed")
    return sheets_edit

# ------------------------------- Main Data Management -------------------------------
def manage_data_edit(sheets_edit):
    if sheets_edit is None:
        st.warning("File not found. Use 'Refresh from GitHub' button in sidebar first.")
        return sheets_edit

    sheets_edit = clean_orphan_maintenance_tasks(sheets_edit)

    if APP_CONFIG["SPARE_PARTS_SHEET"] not in sheets_edit:
        sheets_edit[APP_CONFIG["SPARE_PARTS_SHEET"]] = load_spare_parts()
    if APP_CONFIG["MAINTENANCE_SHEET"] not in sheets_edit:
        sheets_edit[APP_CONFIG["MAINTENANCE_SHEET"]] = load_maintenance_tasks()

    tab_names = ["📋 View & Edit Sections", "🔧 Manage Machines", "➕ Add New Section", "📦 Spare Parts", "🛠 Preventive Maintenance"]
    tabs_edit = st.tabs(tab_names)

    username = st.session_state.get("username")

    with tabs_edit[0]:
        st.subheader("🗂️ View & Edit Section Data")
        st.info("Search and filter (by text, date, machine), then edit directly.")

        all_dept_names = [name for name in sheets_edit.keys() if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]
        dept_names = []
        for dept in all_dept_names:
            if username == "admin" or has_section_permission(username, dept, "edit"):
                dept_names.append(dept)

        if not dept_names:
            st.info("No sections available for editing.")
        else:
            selected_dept = st.selectbox("🏭 Select Section:", dept_names, key="edit_dept_select")
            df_original = sheets_edit[selected_dept].copy()

            st.markdown("### 🔎 Filter Data")
            col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 1])

            with col_f1:
                search_text = st.text_input("🔍 Global search:", placeholder="Enter search term...", key="search_text_edit")
            with col_f2:
                equipment_list = get_equipment_list_from_sheet(df_original)
                equipment_options = ["All"] + equipment_list
                selected_equipment = st.selectbox("🔧 Machine filter:", equipment_options, key="equipment_filter_edit")
            with col_f3:
                use_date_filter = st.checkbox("📅 Filter by date", key="use_date_filter_edit")
                if use_date_filter:
                    date_col_candidates = [col for col in df_original.columns if "date" in col.lower() or "Date" in col]
                    if date_col_candidates:
                        date_col = st.selectbox("Date column:", date_col_candidates, key="date_col_edit")
                    else:
                        date_col = None
                        st.warning("No date column in this section")
                else:
                    date_col = None
            with col_f4:
                st.write("")
                if st.button("🔄 Clear Filters", key="clear_filters_edit"):
                    for key in ["search_text_edit", "equipment_filter_edit", "use_date_filter_edit", "start_date_edit", "end_date_edit"]:
                        if key in st.session_state:
                            st.session_state[key] = None if key != "equipment_filter_edit" else "All"
                    st.rerun()

            if use_date_filter and date_col:
                col_f5, col_f6 = st.columns(2)
                with col_f5:
                    start_date = st.date_input("From:", value=None, key="start_date_edit")
                with col_f6:
                    end_date = st.date_input("To:", value=None, key="end_date_edit")
            else:
                start_date = None
                end_date = None

            df_filtered = df_original.copy()

            if selected_equipment != "All" and "Equipment" in df_filtered.columns:
                df_filtered = df_filtered[df_filtered["Equipment"] == selected_equipment]

            if search_text:
                mask = pd.Series([False] * len(df_filtered), index=df_filtered.index)
                for col in df_filtered.columns:
                    if col not in ["Image_URL"]:
                        mask |= df_filtered[col].astype(str).str.contains(search_text, case=False, na=False)
                df_filtered = df_filtered[mask]

            if use_date_filter and date_col and start_date and end_date:
                try:
                    df_filtered[date_col] = pd.to_datetime(df_filtered[date_col], errors='coerce')
                    df_filtered = df_filtered.dropna(subset=[date_col])
                    mask_date = (df_filtered[date_col] >= pd.to_datetime(start_date)) & (df_filtered[date_col] <= pd.to_datetime(end_date) + timedelta(days=1))
                    df_filtered = df_filtered[mask_date]
                except Exception as e:
                    st.warning(f"Date filter error: {e}")

            col_stat1, col_stat2, col_stat3 = st.columns(3)
            with col_stat1:
                st.metric("📊 Total Records", len(df_original))
            with col_stat2:
                st.metric("🔍 Filtered Records", len(df_filtered))
            with col_stat3:
                if "Equipment" in df_filtered.columns:
                    st.metric("🏭 Unique Machines", df_filtered["Equipment"].nunique())
                else:
                    st.metric("🏭 Unique Machines", "-")

            st.markdown("### ✏️ Edit Data")
            st.caption("Edit cells directly, add rows using 'Add Row', delete rows using 🗑️.")

            display_cols = [col for col in df_filtered.columns if col not in ["Image_URL"]]
            df_display = df_filtered[display_cols].copy()
            df_display = df_display.astype(str).replace('nan', '').replace('None', '')

            edited_df = st.data_editor(
                df_display,
                num_rows="dynamic",
                use_container_width=True,
                height=500,
                key=f"editor_{selected_dept}"
            )

            img_col = None
            if "Image_URL" in df_filtered.columns:
                img_col = "Image_URL"
            if img_col:
                with st.expander("🖼️ View Attached Images"):
                    cols_per_row = 4
                    for i in range(0, min(len(df_filtered), 20), cols_per_row):
                        row_cols = st.columns(cols_per_row)
                        for j, col in enumerate(row_cols):
                            idx = i + j
                            if idx < len(df_filtered):
                                row = df_filtered.iloc[idx]
                                img_url = row.get(img_col, "")
                                if img_url and isinstance(img_url, str) and img_url.strip():
                                    with col:
                                        try:
                                            st.image(img_url, caption=f"Row {idx+1}", width=150)
                                        except Exception:
                                            st.caption(f"Image unavailable: [link]({img_url})")
                                else:
                                    with col:
                                        st.write("No image")

            col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
            with col_btn1:
                if st.button("💾 Save Changes", key=f"save_edit_{selected_dept}", type="primary"):
                    try:
                        merged_df = df_original.copy()
                        for idx in df_filtered.index:
                            if idx in edited_df.index:
                                for col in edited_df.columns:
                                    if col in merged_df.columns:
                                        merged_df.loc[idx, col] = edited_df.loc[idx, col]
                        new_rows = edited_df[~edited_df.index.isin(df_filtered.index)]
                        if not new_rows.empty:
                            merged_df = pd.concat([merged_df, new_rows], ignore_index=True)
                        sheets_edit[selected_dept] = merged_df
                        if save_and_push_to_github(sheets_edit, f"Edit data in section {selected_dept}"):
                            st.cache_data.clear()
                            st.success("Changes saved and pushed to GitHub!")
                            st.rerun()
                        else:
                            st.error("Save failed")
                    except Exception as e:
                        st.error(f"Data save error: {e}")

            with col_btn2:
                excel_file = export_filtered_results_to_excel(df_filtered, selected_dept)
                st.download_button("📥 Download Filtered (Excel)", excel_file, f"{selected_dept}_filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"export_filtered_{selected_dept}")
            with col_btn3:
                all_excel = export_sheet_to_excel({selected_dept: df_original}, selected_dept)
                st.download_button("📥 Download All (Excel)", all_excel, f"{selected_dept}_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"export_all_{selected_dept}")
            with col_btn4:
                full_excel = export_all_sheets_to_excel(sheets_edit)
                st.download_button("📥 Download All Sections", full_excel, f"all_sheets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="export_full_all")

    with tabs_edit[1]:
        if sheets_edit:
            all_dept_names = [name for name in sheets_edit.keys() if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]
            dept_names_machines = []
            for dept in all_dept_names:
                if username == "admin" or has_section_permission(username, dept, "manage_machines"):
                    dept_names_machines.append(dept)
            if dept_names_machines:
                sheet_name = st.selectbox("Select Section:", dept_names_machines, key="manage_machines_sheet_edit")
                manage_machines(sheets_edit, sheet_name, unique_suffix=f"edit_{sheet_name}")
            else:
                st.info("No sections available for machine management.")
        else:
            st.warning("No data")

    with tabs_edit[2]:
        sheets_edit = add_new_department(sheets_edit)

    with tabs_edit[3]:
        sheets_edit = manage_spare_parts_tab(sheets_edit)

    with tabs_edit[4]:
        sheets_edit = preventive_maintenance_tab(sheets_edit)

    return sheets_edit

# ------------------------------- Main UI -------------------------------
with st.sidebar:
    st.header("Session")
    if not st.session_state.get("logged_in"):
        if not login_ui():
            st.stop()
    else:
        state = cleanup_sessions(load_state())
        username = st.session_state.username
        rem = remaining_time(state, username)
        if rem:
            mins, secs = divmod(int(rem.total_seconds()), 60)
            st.success(f"👋 {username} | ⏳ {mins:02d}:{secs:02d}")
        st.markdown("---")
        if st.button("🔄 Refresh", key="sidebar_refresh"):
            if fetch_from_github_requests():
                st.rerun()
        if st.button("Clear Cache", key="clear_cache"):
            st.cache_data.clear()
            st.rerun()
        if st.button("🚪 Logout", key="logout_button"):
            logout_action()

all_sheets = load_all_sheets()
sheets_edit = load_sheets_for_edit()
st.title(f"{APP_CONFIG['APP_ICON']} {APP_CONFIG['APP_TITLE']}")
user_role = st.session_state.get("user_role", "viewer")
username = st.session_state.get("username", "")

def user_can(permission_type):
    if username == "admin":
        return True
    perms = get_user_permissions(username)
    if perms.get("all_sections", False):
        return True
    for perms_list in perms.get("sections_permissions", {}).values():
        if permission_type in perms_list:
            return True
    return False

can_add_event = user_can("add_event")
can_manage_machines = user_can("manage_machines")
can_edit_data = user_can("edit")

tabs_list = ["🔍 Advanced Search", "📊 Fault Analysis", "🔔 Notifications"]
if can_add_event:
    tabs_list.append("➕ Add Fault Event")
if can_manage_machines:
    tabs_list.append("🔧 Manage Machines")
if can_edit_data:
    tabs_list.append("🛠 Data Management")
if username == "admin":
    tabs_list.append("👥 User Management")
tabs_list.append("📞 Support")

tabs = st.tabs(tabs_list)
idx = 0

with tabs[idx]:
    search_across_sheets(all_sheets)
idx += 1
with tabs[idx]:
    failures_analysis_tab(all_sheets)
idx += 1

# ------------------------------- Notifications Tab -------------------------------
with tabs[idx]:
    st.header("🔔 Notifications & Alerts")

    auto_refresh = st.checkbox("🔄 Auto refresh (every 30 seconds)", value=True, key="auto_refresh_checkbox")
    if auto_refresh:
        st.components.v1.html("""
        <script>
        setInterval(function() {
            location.reload();
        }, 30000);
        </script>
        """, height=0)
        st.info("Auto refresh enabled. Page will refresh every 30 seconds.")

    clean_old_activity_log(days_to_keep=1)

    username = st.session_state.get("username")
    user_role = st.session_state.get("user_role", "viewer")
    all_sheets = load_all_sheets()

    existing_sections = [name for name in all_sheets.keys()
                        if name not in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]]

    existing_equipment = []
    for sheet_name in existing_sections:
        if sheet_name in all_sheets:
            df = all_sheets[sheet_name]
            if "Equipment" in df.columns:
                existing_equipment.extend(df["Equipment"].dropna().unique())
    existing_equipment = [str(eq).strip() for eq in existing_equipment if str(eq).strip() != ""]

    allowed_sections = get_allowed_sections(all_sheets, username, "view")
    allowed_sections = [sec for sec in allowed_sections if sec in existing_sections]

    st.subheader("🛠️ Preventive Maintenance & Critical Spare Parts Alerts")

    allowed_equipment = []
    for sheet_name in allowed_sections:
        if sheet_name in all_sheets:
            df = all_sheets[sheet_name]
            if "Equipment" in df.columns:
                allowed_equipment.extend(df["Equipment"].dropna().unique())
    allowed_equipment = [str(eq).strip() for eq in allowed_equipment if str(eq).strip() != ""]

    overdue, upcoming = get_upcoming_maintenance(3)

    if not overdue.empty and "Equipment" in overdue.columns:
        overdue = overdue[overdue["Equipment"].isin(existing_equipment)]
    if not upcoming.empty and "Equipment" in upcoming.columns:
        upcoming = upcoming[upcoming["Equipment"].isin(existing_equipment)]

    if username != "admin" and user_role != "admin":
        overdue = overdue[overdue["Equipment"].isin(allowed_equipment)] if not overdue.empty else overdue
        upcoming = upcoming[upcoming["Equipment"].isin(allowed_equipment)] if not upcoming.empty else upcoming

    maintenance_text_parts = []
    for _, row in overdue.iterrows():
        eq = row['Equipment']
        task = row['Task_Name']
        due_date = row['Next_Date'].strftime('%Y-%m-%d') if pd.notna(row['Next_Date']) else "Not Set"
        section = "Unknown"
        for sheet_name in allowed_sections:
            if sheet_name in all_sheets and eq in all_sheets[sheet_name]["Equipment"].values:
                section = sheet_name
                break
        maintenance_text_parts.append(f"OVERDUE: {eq} - {task} (Due: {due_date}) [Section: {section}]")

    for _, row in upcoming.iterrows():
        eq = row['Equipment']
        task = row['Task_Name']
        days = (row['Next_Date'].date() - datetime.now().date()).days
        due_date = row['Next_Date'].strftime('%Y-%m-%d') if pd.notna(row['Next_Date']) else "Not Set"
        section = "Unknown"
        for sheet_name in allowed_sections:
            if sheet_name in all_sheets and eq in all_sheets[sheet_name]["Equipment"].values:
                section = sheet_name
                break
        maintenance_text_parts.append(f"UPCOMING: {eq} - {task} (in {days} days - {due_date}) [Section: {section}]")

    critical = get_critical_spare_parts()
    if username != "admin" and user_role != "admin":
        critical = [part for part in critical if part.get("Section", "") in allowed_sections]

    for part in critical:
        maintenance_text_parts.append(f"CRITICAL PART: {part['Part_Name']} (Stock: {part['Available_Quantity']} < Threshold: {part['Alert_Threshold']}) [Section: {part['Section']}]")

    if maintenance_text_parts:
        text_to_scroll = " | ".join(maintenance_text_parts)
        st.markdown(f"""
        <style>
        @keyframes scroll-text {{
            0% {{ transform: translateX(-100%); }}
            100% {{ transform: translateX(100%); }}
        }}
        .scrolling-wrapper {{
            overflow: hidden;
            white-space: nowrap;
            background-color: #f8f9fa;
            border: 2px solid #ffc107;
            border-radius: 8px;
            padding: 15px 0;
            width: 100%;
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
            margin: 10px 0;
        }}
        .scrolling-content {{
            display: inline-block;
            animation: scroll-text 120s linear infinite;
            font-size: 22px;
            font-weight: bold;
            color: #1a1a2e;
            padding-left: 100%;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}
        .scrolling-content:hover {{
            animation-play-state: paused;
        }}
        </style>
        <div class="scrolling-wrapper">
            <div class="scrolling-content">
                {text_to_scroll}
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.success("No overdue or upcoming maintenance, and no critical spare parts.")

    st.markdown("---")
    st.subheader("📋 Maintenance Details (Table)")

    maintenance_df_data = []
    for _, row in overdue.iterrows():
        eq = row['Equipment']
        task = row['Task_Name']
        due_date = row['Next_Date'].strftime('%Y-%m-%d') if pd.notna(row['Next_Date']) else "Not Set"
        section = "Unknown"
        for sheet_name in allowed_sections:
            if sheet_name in all_sheets and eq in all_sheets[sheet_name]["Equipment"].values:
                section = sheet_name
                break
        maintenance_df_data.append({"Equipment": eq, "Status": "🔴 Overdue", "Task": task, "Due_Date": due_date, "Section": section})

    for _, row in upcoming.iterrows():
        eq = row['Equipment']
        task = row['Task_Name']
        days = (row['Next_Date'].date() - datetime.now().date()).days
        due_date = row['Next_Date'].strftime('%Y-%m-%d') if pd.notna(row['Next_Date']) else "Not Set"
        section = "Unknown"
        for sheet_name in allowed_sections:
            if sheet_name in all_sheets and eq in all_sheets[sheet_name]["Equipment"].values:
                section = sheet_name
                break
        maintenance_df_data.append({"Equipment": eq, "Status": f"🟡 Upcoming (in {days} days)", "Task": task, "Due_Date": due_date, "Section": section})

    if maintenance_df_data:
        df_display = pd.DataFrame(maintenance_df_data)
        st.dataframe(df_display, use_container_width=True, height=400)
    else:
        st.info("No maintenance data to display.")

    st.markdown("---")
    st.subheader("📋 Events & Spare Parts (Collapsible)")

    with st.expander("📋 Recent Events", expanded=False):
        activity_log = load_activity_log()
        filtered_log = []
        for entry in activity_log:
            section = entry.get("section", "")
            if username == "admin" or user_role == "admin":
                filtered_log.append(entry)
            else:
                if not section or section in existing_sections:
                    filtered_log.append(entry)
        recent_log = filtered_log[:20]
        if recent_log:
            with st.container(height=200):
                for entry in recent_log:
                    timestamp = datetime.fromisoformat(entry["timestamp"]).strftime("%Y-%m-%d %H:%M:%S")
                    action_type = entry.get("action_type", "event")
                    username_act = entry.get("username", "unknown")
                    details = entry.get("details", "")
                    section = entry.get("section", "")
                    icon = {"add_event": "🆕", "execute_maintenance": "✅", "add_spare_part": "🔩", "add_maintenance_task": "🛠️", "delete_section": "🗑️"}.get(action_type, "📌")
                    section_display = f" (Section: {section})" if section else ""
                    st.info(f"{icon} **{timestamp}** - **{username_act}**{section_display}: {details}")
        else:
            st.info("No events in the last 24 hours.")

    with st.expander("⚠️ Critical Spare Parts (Details)", expanded=False):
        critical = get_critical_spare_parts()
        if username != "admin" and user_role != "admin":
            critical = [part for part in critical if part.get("Section", "") in allowed_sections]
        if critical:
            with st.container(height=150):
                for part in critical:
                    threshold = part.get('Alert_Threshold', 1)
                    section_name = part.get('Section', 'Unknown')
                    st.error(f"🔴 **{part['Part_Name']}** (Section: {section_name}) - Stock: {part['Available_Quantity']} < Threshold: {threshold}")
        else:
            st.success("No critical spare parts.")

    if st.button("🔄 Refresh Now", key="manual_refresh"):
        st.rerun()

idx += 1

if can_add_event:
    with tabs[idx]:
        if sheets_edit:
            allowed_for_add = []
            for sheet_name in sheets_edit.keys():
                if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
                    continue
                if has_section_permission(username, sheet_name, "add_event"):
                    allowed_for_add.append(sheet_name)
            if allowed_for_add:
                sheet_name = st.selectbox("Select Section:", allowed_for_add, key="add_event_sheet_main")
                sheets_edit = add_new_event(sheets_edit, sheet_name)
            else:
                st.warning("No sections available for adding events.")
        else:
            st.warning("No data")
    idx += 1

if can_manage_machines:
    with tabs[idx]:
        if sheets_edit:
            allowed_for_machines = []
            for sheet_name in sheets_edit.keys():
                if sheet_name in [APP_CONFIG["SPARE_PARTS_SHEET"], APP_CONFIG["MAINTENANCE_SHEET"]]:
                    continue
                if has_section_permission(username, sheet_name, "manage_machines"):
                    allowed_for_machines.append(sheet_name)
            if allowed_for_machines:
                sheet_name = st.selectbox("Select Section:", allowed_for_machines, key="manage_machines_sheet_main")
                manage_machines(sheets_edit, sheet_name, unique_suffix="main")
            else:
                st.warning("No sections available for machine management.")
        else:
            st.warning("No data")
    idx += 1

if can_edit_data:
    with tabs[idx]:
        sheets_edit = manage_data_edit(sheets_edit)
    idx += 1

if username == "admin":
    with tabs[idx]:
        admin_users_management_tab()
    idx += 1

with tabs[idx]:
    st.header("📞 Technical Support")
    st.markdown("### Designed and built by **Eng. Mohamed Abdallah**")
    st.markdown("#### Head of zahra & Preparations Dept., Bil Yarn 1 Factory")
    st.markdown("---")
    st.markdown("📧 **Contact & Support:** `01274424062`")
    st.markdown("---")
    YOUTUBE_LINK = "https://youtube.com/@cardtrutchler?si=bayhxhRXgCzWSpCl"
    st.markdown(f"[📺 Official YouTube Channel]({YOUTUBE_LINK})")
    st.caption(f"Channel link: {YOUTUBE_LINK}")
    st.markdown("---")
    support_config = load_support_config()
    current_image_url = support_config.get("image_url", "")
    st.subheader("🖼️ Developer Image")
    if current_image_url and current_image_url.strip():
        try:
            st.image(current_image_url, use_container_width=True)
            st.caption("Current developer image")
        except:
            st.warning("Could not display saved image")
    else:
        if st.session_state.get("username") == "admin":
            st.info("No developer image yet. Upload now (one time only).")
            uploaded_img = st.file_uploader("Upload developer image (jpg, png, ...)", type=APP_CONFIG["ALLOWED_IMAGE_TYPES"], key="support_img_upload_once")
            if uploaded_img is not None:
                with st.spinner("Uploading..."):
                    image_url = upload_image_to_github(uploaded_img, "support", "developer_image_final")
                    if image_url:
                        support_config["image_url"] = image_url
                        save_support_config(support_config)
                        st.success("Image uploaded successfully!")
                        st.rerun()
                    else:
                        st.error("Image upload failed, try again.")
        else:
            st.info("No developer image yet. It will be uploaded by system admin.")

    st.markdown("---")
    st.subheader("📧 Email Test")
    if st.button("📧 Send Test Email"):
        test_subject = "Test Email from CMMS System"
        test_body = f"""This is a test message from CMMS system.

Sent at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
User: {st.session_state.get('username', 'unknown')}

If you received this, email settings are working correctly.

--- Current Notifications ---
{get_current_notifications_text()}
        """
        if send_email(test_subject, test_body):
            st.success("Test email sent successfully!")
        else:
            st.error("Failed to send test email. Check SMTP settings in secrets.")
