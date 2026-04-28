"""
Baldovin Construction Co. — Project Agent
Monitors baldovinconstructionco@gmail.com and Google Drive.
Handles project kickoff, assignments, triggers, follow-ups, and reporting.
"""

import os
import re
import time
import base64
import logging
import json
from datetime import datetime, timedelta, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import anthropic
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account

# ─────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger("baldovin-agent")

# ─────────────────────────────────────────────────────────────
# ENVIRONMENT VARIABLES
# All secrets live here — never hardcoded in the script.
# Set these in Railway dashboard under Variables.
# ─────────────────────────────────────────────────────────────
ANTHROPIC_API_KEY       = os.environ["ANTHROPIC_API_KEY"]
GOOGLE_CREDS_JSON       = os.environ["GOOGLE_CREDS_JSON"]       # Full JSON string of service account creds
AGENT_EMAIL             = os.environ.get("AGENT_EMAIL", "baldovinconstructionco@gmail.com")
SETTINGS_SPREADSHEET_ID = os.environ["SETTINGS_SPREADSHEET_ID"] # Google Sheets ID for Agent Settings
DRIVE_PROJECTS_FOLDER   = os.environ["DRIVE_PROJECTS_FOLDER"]   # ID of the "Projects" folder in Drive
DRIVE_TEMPLATE_FOLDER   = os.environ["DRIVE_TEMPLATE_FOLDER"]   # ID of the "Template" folder
DRIVE_UPCOMING_FOLDER   = os.environ["DRIVE_UPCOMING_FOLDER"]   # ID of "Upcoming" folder
DRIVE_CURRENT_FOLDER    = os.environ["DRIVE_CURRENT_FOLDER"]    # ID of "Current" folder
COMPANYCAM_API_KEY      = os.environ.get("COMPANYCAM_API_KEY", "")
POLL_INTERVAL_SECONDS   = int(os.environ.get("POLL_INTERVAL_SECONDS", "60"))

# ─────────────────────────────────────────────────────────────
# GOOGLE API SCOPES
# ─────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ─────────────────────────────────────────────────────────────
# COMMAND PATTERNS
# Subject line formats the agent watches for.
# ─────────────────────────────────────────────────────────────
CMD_NEW_PROJECT     = re.compile(r"NEW PROJECT\s*[-—]\s*(.+)", re.IGNORECASE)
CMD_ASSIGN_DESIGNER = re.compile(r"ASSIGN DESIGNER\s*[-—]\s*(P-\d+)\s*[-—]\s*(.+)", re.IGNORECASE)
CMD_ASSIGN_PM       = re.compile(r"ASSIGN PM\s*[-—]\s*(P-\d+)\s*[-—]\s*(.+)", re.IGNORECASE)
CMD_ASSIGN_SUB      = re.compile(r"ASSIGN SUB\s*[-—]\s*(P-\d+)\s*[-—]\s*(.+)", re.IGNORECASE)
CMD_ASSIGN_SUPER    = re.compile(r"ASSIGN SUPER\s*[-—]\s*(P-\d+)\s*[-—]\s*(.+)", re.IGNORECASE)
CMD_WEEKLY_REPORT   = re.compile(r"REPORT REQUEST\s*[-—]\s*WEEKLY CLOSEOUT", re.IGNORECASE)

# ─────────────────────────────────────────────────────────────
# ROLE PERMISSIONS
# Maps which roles can issue which assignment commands.
# ─────────────────────────────────────────────────────────────
ROLE_PERMISSIONS = {
    "Estimator":         ["ASSIGN DESIGNER", "ASSIGN PM"],
    "Project Manager":   ["ASSIGN SUB", "ASSIGN SUPER"],
    "Director Ops & BD": ["ASSIGN DESIGNER", "ASSIGN PM", "ASSIGN SUB", "ASSIGN SUPER"],
    "Leadership":        ["ASSIGN DESIGNER", "ASSIGN PM", "ASSIGN SUB", "ASSIGN SUPER"],
    "Admin":             ["ASSIGN DESIGNER", "ASSIGN PM", "ASSIGN SUB", "ASSIGN SUPER"],
}

# ─────────────────────────────────────────────────────────────
# GOOGLE SERVICES INITIALIZATION
# ─────────────────────────────────────────────────────────────
def build_google_services():
    creds_info = json.loads(GOOGLE_CREDS_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES
    )
    gmail   = build("gmail",   "v1",  credentials=creds)
    drive   = build("drive",   "v3",  credentials=creds)
    sheets  = build("sheets",  "v4",  credentials=creds)
    return gmail, drive, sheets

# ─────────────────────────────────────────────────────────────
# ANTHROPIC CLIENT
# ─────────────────────────────────────────────────────────────
claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

def ask_claude(prompt: str, system: str = None) -> str:
    """Send a prompt to Claude and return the text response."""
    messages = [{"role": "user", "content": prompt}]
    kwargs = {"model": "claude-sonnet-4-20250514", "max_tokens": 1000, "messages": messages}
    if system:
        kwargs["system"] = system
    response = claude.messages.create(**kwargs)
    return response.content[0].text.strip()

# ─────────────────────────────────────────────────────────────
# SETTINGS SHEET READERS
# ─────────────────────────────────────────────────────────────
def read_sheet(sheets, range_name: str) -> list:
    """Read a range from the Agent Settings Google Sheet."""
    result = sheets.spreadsheets().values().get(
        spreadsheetId=SETTINGS_SPREADSHEET_ID,
        range=range_name
    ).execute()
    return result.get("values", [])

def write_sheet(sheets, range_name: str, values: list):
    """Write values to the Agent Settings Google Sheet."""
    sheets.spreadsheets().values().update(
        spreadsheetId=SETTINGS_SPREADSHEET_ID,
        range=range_name,
        valueInputOption="RAW",
        body={"values": values}
    ).execute()

def get_team_roster(sheets) -> dict:
    """
    Returns a dict keyed by lowercase alias -> {name, email, phone, role}.
    Used to resolve first names in assignment commands.
    """
    rows = read_sheet(sheets, "Team Roster!A:F")
    roster = {}
    for row in rows[1:]:  # skip header
        if len(row) < 5 or not row[1]:  # need at least name + email
            continue
        name, email, phone, role = row[0], row[1], row[2] if len(row) > 2 else "", row[3] if len(row) > 3 else ""
        aliases_raw = row[4] if len(row) > 4 else name
        for alias in aliases_raw.split(","):
            key = alias.strip().lower()
            if key:
                roster[key] = {"name": name, "email": email, "phone": phone, "role": role}
    return roster

def get_whitelist(sheets) -> dict:
    """Returns dict keyed by email -> {name, role, access_level, active}."""
    rows = read_sheet(sheets, "Team Whitelist!A:F")
    whitelist = {}
    for row in rows[1:]:
        if len(row) < 6 or not row[1]:
            continue
        email = row[1].strip().lower()
        whitelist[email] = {
            "name":         row[0],
            "role":         row[2] if len(row) > 2 else "",
            "access_level": row[3] if len(row) > 3 else "",
            "active":       (row[5].strip().upper() == "Y") if len(row) > 5 else False,
        }
    return whitelist

def get_project_registry(sheets) -> list:
    """Returns list of project dicts from Project Registry sheet."""
    rows = read_sheet(sheets, "Project Registry!A:P")
    if not rows or len(rows) < 2:
        return []
    headers = rows[0]
    projects = []
    for row in rows[1:]:
        if not row or not row[0]:
            continue
        proj = {}
        for i, h in enumerate(headers):
            proj[h] = row[i] if i < len(row) else ""
        projects.append(proj)
    return projects

def get_next_project_id(sheets) -> str:
    """Reads existing project IDs and returns the next one (P-001, P-002, etc.)."""
    projects = get_project_registry(sheets)
    if not projects:
        return "P-001"
    ids = [p.get("Project ID", "") for p in projects if p.get("Project ID", "").startswith("P-")]
    nums = []
    for pid in ids:
        try:
            nums.append(int(pid.replace("P-", "")))
        except ValueError:
            pass
    next_num = max(nums) + 1 if nums else 1
    return f"P-{next_num:03d}"

def add_project_to_registry(sheets, project: dict):
    """Appends a new project row to the Project Registry."""
    rows = read_sheet(sheets, "Project Registry!A:P")
    next_row = len(rows) + 1
    row_data = [
        project.get("Project ID", ""),
        project.get("Project Name", ""),
        project.get("Address", ""),
        project.get("Client Name", ""),
        project.get("Estimator", ""),
        "Chase, Samantha",          # Leadership always auto-assigned
        "",                          # Designer — TBD
        "",                          # PM — TBD
        "",                          # Sub Coordinator — TBD
        "",                          # Field Super — TBD
        "",                          # Renegade — TBD
        "Michelina",                 # Finance always auto-assigned
        project.get("Kickoff Date", str(date.today())),
        "Estimating",                # Starting phase
        "Y",                         # Active
        "Upcoming",                  # Folder location
    ]
    write_sheet(sheets, f"Project Registry!A{next_row}:P{next_row}", [row_data])
    log.info(f"Project {project['Project ID']} added to registry.")

def update_project_field(sheets, project_id: str, field: str, value: str):
    """Updates a single field on a project row in the registry."""
    rows = read_sheet(sheets, "Project Registry!A:P")
    if not rows:
        return
    headers = rows[0]
    if field not in headers:
        log.warning(f"Field '{field}' not found in Project Registry headers.")
        return
    col_idx = headers.index(field)
    for row_idx, row in enumerate(rows[1:], start=2):
        if row and row[0] == project_id:
            col_letter = chr(ord("A") + col_idx)
            write_sheet(sheets, f"Project Registry!{col_letter}{row_idx}", [[value]])
            log.info(f"Updated {project_id} field '{field}' = '{value}'")
            return
    log.warning(f"Project {project_id} not found in registry.")

# ─────────────────────────────────────────────────────────────
# GMAIL HELPERS
# ─────────────────────────────────────────────────────────────
def send_email(gmail, to: str, subject: str, body: str, cc: str = None):
    """Send an email from the agent Gmail account."""
    msg = MIMEMultipart()
    msg["To"]      = to
    msg["From"]    = AGENT_EMAIL
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = cc
    msg.attach(MIMEText(body, "plain"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    gmail.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()
    log.info(f"Email sent to {to} — Subject: {subject}")

def get_unread_messages(gmail) -> list:
    """Fetch unread messages in the agent inbox."""
    result = gmail.users().messages().list(
        userId="me", q="is:unread"
    ).execute()
    return result.get("messages", [])

def get_message_detail(gmail, msg_id: str) -> dict:
    """Get full message details including sender, subject, body."""
    msg = gmail.users().messages().get(
        userId="me", id=msg_id, format="full"
    ).execute()
    headers = {h["name"].lower(): h["value"] for h in msg["payload"]["headers"]}
    subject = headers.get("subject", "")
    sender  = headers.get("from", "")
    # Extract email address from "Name <email>" format
    email_match = re.search(r"<(.+?)>", sender)
    sender_email = email_match.group(1).lower() if email_match else sender.lower()
    # Extract body
    body = ""
    if "parts" in msg["payload"]:
        for part in msg["payload"]["parts"]:
            if part["mimeType"] == "text/plain":
                data = part["body"].get("data", "")
                body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
                break
    elif "body" in msg["payload"]:
        data = msg["payload"]["body"].get("data", "")
        body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
    return {
        "id":           msg_id,
        "subject":      subject,
        "sender_email": sender_email,
        "sender_name":  headers.get("from", ""),
        "body":         body,
        "date":         headers.get("date", ""),
    }

def mark_as_read(gmail, msg_id: str):
    gmail.users().messages().modify(
        userId="me", id=msg_id,
        body={"removeLabelIds": ["UNREAD"]}
    ).execute()

# ─────────────────────────────────────────────────────────────
# GOOGLE DRIVE HELPERS
# ─────────────────────────────────────────────────────────────
def copy_template_folder(drive, project_name: str) -> str:
    """
    Copies the Template folder, renames it to project_name,
    moves it to the Upcoming folder. Returns new folder ID.
    NOTE: Drive API copies files not folders natively.
    We create a new folder and copy each file from Template into it.
    """
    # Create new project folder in Upcoming
    folder_meta = {
        "name":     project_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents":  [DRIVE_UPCOMING_FOLDER],
    }
    new_folder = drive.files().create(body=folder_meta, fields="id").execute()
    new_folder_id = new_folder["id"]

    # List all files and subfolders in Template
    template_children = drive.files().list(
        q=f"'{DRIVE_TEMPLATE_FOLDER}' in parents and trashed=false",
        fields="files(id, name, mimeType)"
    ).execute().get("files", [])

    # Copy each item into the new project folder
    for item in template_children:
        if item["mimeType"] == "application/vnd.google-apps.folder":
            # Create matching subfolder
            sub_meta = {
                "name":     item["name"],
                "mimeType": "application/vnd.google-apps.folder",
                "parents":  [new_folder_id],
            }
            drive.files().create(body=sub_meta).execute()
        else:
            # Copy the file
            drive.files().copy(
                fileId=item["id"],
                body={"name": item["name"], "parents": [new_folder_id]}
            ).execute()

    log.info(f"Project folder '{project_name}' created in Upcoming. ID: {new_folder_id}")
    return new_folder_id

def detect_folder_moves(drive, sheets) -> list:
    """
    Checks if any project folder has been moved to the Current folder.
    Returns list of projects that moved.
    """
    moved = []
    projects = get_project_registry(sheets)
    current_children = drive.files().list(
        q=f"'{DRIVE_CURRENT_FOLDER}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)"
    ).execute().get("files", [])
    current_names = {f["name"].lower() for f in current_children}

    for proj in projects:
        if proj.get("Folder Location") == "Upcoming" and proj.get("Active (Y/N)") == "Y":
            if proj.get("Project Name", "").lower() in current_names:
                moved.append(proj)
    return moved

def read_schedule_close_date(drive, project_folder_id: str) -> str:
    """
    Looks inside the project's Schedule subfolder for an Excel file.
    Reads the latest Finish date as the project close date.
    Returns date string or empty string if not found.
    """
    try:
        # Find Schedule subfolder
        subfolders = drive.files().list(
            q=f"'{project_folder_id}' in parents and name='Schedule' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id)"
        ).execute().get("files", [])
        if not subfolders:
            return ""
        schedule_folder_id = subfolders[0]["id"]

        # Find Excel file in Schedule folder
        excel_files = drive.files().list(
            q=f"'{schedule_folder_id}' in parents and trashed=false",
            fields="files(id, name)"
        ).execute().get("files", [])
        if not excel_files:
            return ""

        # Download first Excel file and read with pandas
        file_id = excel_files[0]["id"]
        request = drive.files().get_media(fileId=file_id)
        import io
        fh = io.BytesIO()
        from googleapiclient.http import MediaIoBaseDownload
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        df = pd.read_csv(fh) if excel_files[0]["name"].endswith(".csv") else pd.read_excel(fh)

        # Find the Finish column and get the max date
        finish_col = next((c for c in df.columns if "finish" in c.lower()), None)
        if not finish_col:
            return ""
        dates = pd.to_datetime(df[finish_col], errors="coerce").dropna()
        if dates.empty:
            return ""
        close_date = dates.max().strftime("%Y-%m-%d")
        log.info(f"Schedule close date detected: {close_date}")
        return close_date
    except Exception as e:
        log.warning(f"Could not read schedule close date: {e}")
        return ""

# ─────────────────────────────────────────────────────────────
# AUTHORIZATION CHECK
# ─────────────────────────────────────────────────────────────
def is_authorized(sender_email: str, command: str, whitelist: dict) -> tuple:
    """
    Checks if the sender is in the whitelist and has permission
    to issue the given command. Returns (allowed: bool, reason: str).
    """
    person = whitelist.get(sender_email.lower())
    if not person:
        return False, f"Your email ({sender_email}) is not registered in the Baldovin Agent system."
    if not person["active"]:
        return False, f"Your account is currently inactive in the agent system. Contact Chase to activate."
    role = person["role"]
    allowed_commands = ROLE_PERMISSIONS.get(role, [])
    command_upper = command.upper()
    if not any(cmd in command_upper for cmd in allowed_commands):
        return False, (
            f"Your role ({role}) does not have permission to issue this command. "
            f"If you believe this is an error, contact Chase."
        )
    return True, "authorized"

# ─────────────────────────────────────────────────────────────
# NAME RESOLVER
# ─────────────────────────────────────────────────────────────
def resolve_name(name_input: str, roster: dict) -> dict | None:
    """
    Resolves a partial name or alias to a full roster record.
    Returns None if ambiguous or not found.
    """
    key = name_input.strip().lower()
    matches = {alias: record for alias, record in roster.items() if key in alias}
    # Deduplicate by email
    unique = {v["email"]: v for v in matches.values()}
    if len(unique) == 1:
        return list(unique.values())[0]
    return None  # ambiguous or not found

# ─────────────────────────────────────────────────────────────
# PARSE KICKOFF EMAIL BODY
# ─────────────────────────────────────────────────────────────
def parse_kickoff_body(body: str) -> dict:
    """
    Parses project data from either:
    - A JSON string passed from the Claude brain (already extracted)
    - A raw email body (needs Claude to extract it)
    """
    # If the body is already JSON from the Claude brain, use it directly
    try:
        data = json.loads(body)
        if isinstance(data, dict) and any(k in data for k in ["project_name", "address", "client_name"]):
            return data
    except Exception:
        pass

    # Otherwise use Claude to extract from raw email body
    prompt = f"""
You are parsing a project kickoff email for Baldovin Construction Co.
Extract the following fields and return ONLY valid JSON, no markdown, no preamble.
Fields: project_name, address, client_name, client_email, client_phone, scope, estimated_value, notes.
If a field is missing use an empty string.

Email body:
{body}
"""
    raw = ask_claude(prompt)
    try:
        clean = raw.replace("```json", "").replace("```", "").strip()
        return json.loads(clean)
    except Exception:
        log.warning("Could not parse kickoff body — using empty project dict.")
        return {}

# ─────────────────────────────────────────────────────────────
# COMMAND HANDLERS
# ─────────────────────────────────────────────────────────────
def handle_new_project(msg: dict, gmail, drive, sheets, whitelist: dict, roster: dict):
    """Handles NEW PROJECT kickoff email from estimator."""
    sender = msg["sender_email"]
    allowed, reason = is_authorized(sender, "NEW PROJECT", whitelist)
    # New project can be initiated by Estimator, Director, or Leadership
    person = whitelist.get(sender, {})
    if person.get("role") not in ["Estimator", "Director Ops & BD", "Leadership", "Admin"]:
        send_email(gmail, sender, "Agent — Project Kickoff Rejected",
            f"Hi,\n\nOnly Estimators can initiate new projects via the agent.\n\nReason: {reason}\n\n— Baldovin Agent")
        return

    project_data = parse_kickoff_body(msg["body"])
    match = CMD_NEW_PROJECT.search(msg["subject"])
    if match:
        project_data["project_name"] = project_data.get("project_name") or match.group(1).strip()

    if not project_data.get("project_name"):
        send_email(gmail, sender, "Agent — Kickoff Failed: Missing Project Name",
            "Hi,\n\nI couldn't detect a project name from your email. "
            "Please use the subject format:\n\nNEW PROJECT — [Project Name]\n\n— Baldovin Agent")
        return

    project_id = get_next_project_id(sheets)
    project_data["Project ID"] = project_id
    project_data["Project Name"] = project_data.get("project_name", "")
    project_data["Address"] = project_data.get("address", "")
    project_data["Client Name"] = project_data.get("client_name", "")
    project_data["Kickoff Date"] = str(date.today())
    project_data["Estimator"] = person.get("name", sender)

    # Create Drive folder
    folder_id = copy_template_folder(drive, project_data["Project Name"])
    project_data["drive_folder_id"] = folder_id

    # Register project
    add_project_to_registry(sheets, project_data)

    # Confirm to estimator
    send_email(gmail, sender,
        f"Agent — Project Created: {project_data['Project Name']} ({project_id})",
        f"Hi {person.get('name', '')},\n\n"
        f"Project {project_id} has been created and registered.\n\n"
        f"Project: {project_data['Project Name']}\n"
        f"Client: {project_data.get('client_name', 'TBD')}\n"
        f"Address: {project_data.get('address', 'TBD')}\n"
        f"Drive folder has been created in Upcoming.\n\n"
        f"Next steps:\n"
        f"  • Complete your estimate\n"
        f"  • When ready, assign the designer:\n"
        f"    ASSIGN DESIGNER — {project_id} — [Designer Name]\n"
        f"  • Then assign the PM:\n"
        f"    ASSIGN PM — {project_id} — [PM Name]\n\n"
        f"— Baldovin Agent"
    )

    # Notify leadership
    leadership_emails = [v["email"] for v in whitelist.values()
                         if v["role"] == "Leadership" and v["active"]]
    for le in leadership_emails:
        send_email(gmail, le,
            f"New Project Registered: {project_data['Project Name']} ({project_id})",
            f"A new project has been created.\n\n"
            f"Project ID: {project_id}\n"
            f"Name: {project_data['Project Name']}\n"
            f"Client: {project_data.get('client_name', 'TBD')}\n"
            f"Address: {project_data.get('address', 'TBD')}\n"
            f"Estimator: {project_data['Estimator']}\n"
            f"Scope: {project_data.get('scope', 'TBD')}\n"
            f"Est. Value: {project_data.get('estimated_value', 'TBD')}\n\n"
            f"— Baldovin Agent"
        )
    log.info(f"New project {project_id} — {project_data['Project Name']} created.")

def handle_assignment(msg: dict, command: str, project_id: str,
                       assignee_input: str, gmail, sheets,
                       whitelist: dict, roster: dict):
    """Generic handler for ASSIGN DESIGNER / ASSIGN PM / ASSIGN SUB / ASSIGN SUPER."""
    sender = msg["sender_email"]
    allowed, reason = is_authorized(sender, command, whitelist)
    if not allowed:
        send_email(gmail, sender, f"Agent — Assignment Rejected",
            f"Hi,\n\nYour assignment command could not be processed.\n\nReason: {reason}\n\n— Baldovin Agent")
        return

    sender_info = whitelist.get(sender, {})
    assignee = resolve_name(assignee_input, roster)

    if assignee is None:
        # Check for ambiguity — multiple matches
        key = assignee_input.strip().lower()
        matches = {alias: r for alias, r in roster.items() if key in alias}
        unique_emails = list({v["email"]: v for v in matches.values()}.values())
        if len(unique_emails) > 1:
            names_list = "\n".join([f"  • {r['name']} ({r['role']})" for r in unique_emails])
            send_email(gmail, sender, f"Agent — Ambiguous Name: {assignee_input}",
                f"Hi {sender_info.get('name', '')},\n\n"
                f"I found multiple people matching '{assignee_input}':\n\n{names_list}\n\n"
                f"Please resend the command using the full last name to clarify.\n\n— Baldovin Agent")
        else:
            send_email(gmail, sender, f"Agent — Name Not Found: {assignee_input}",
                f"Hi {sender_info.get('name', '')},\n\n"
                f"I couldn't find '{assignee_input}' in the team roster. "
                f"Please check the name and try again.\n\n— Baldovin Agent")
        return

    # Map command to registry field
    field_map = {
        "ASSIGN DESIGNER": "Designer",
        "ASSIGN PM":       "PM Assigned",
        "ASSIGN SUB":      "Sub Coordinator",
        "ASSIGN SUPER":    "Field Super",
    }
    field = field_map.get(command.upper())
    if not field:
        return

    update_project_field(sheets, project_id, field, assignee["name"])

    # Update phase if PM is being assigned (moves to Pre-Construction & Design)
    if command.upper() == "ASSIGN PM":
        update_project_field(sheets, project_id, "Current Phase", "Pre-Construction & Design")

    # Confirm to sender
    send_email(gmail, sender,
        f"Agent — {command.title()} Confirmed: {project_id}",
        f"Hi {sender_info.get('name', '')},\n\n"
        f"{assignee['name']} has been assigned as {field} on {project_id}.\n\n"
        f"I've notified them directly.\n\n— Baldovin Agent"
    )

    # Notify the assignee
    role_label = field.replace(" Assigned", "")
    notification_body = (
        f"Hi {assignee['name']},\n\n"
        f"You have been assigned as {role_label} on project {project_id}.\n\n"
        f"Assigned by: {sender_info.get('name', sender)}\n"
        f"Project Registry has been updated.\n\n"
    )
    if command.upper() == "ASSIGN PM":
        notification_body += (
            f"As Project Manager, please complete the following:\n"
            f"  1. Reply YES or NO to this email: Does this project require a Certificate of Occupancy (CO)?\n"
            f"  2. Build your construction schedule and upload it to the Schedule folder in Drive.\n"
            f"  3. Assign subs as needed:\n"
            f"     ASSIGN SUB — {project_id} — [Sub Name]\n\n"
        )
    notification_body += "— Baldovin Agent"

    send_email(gmail, assignee["email"],
        f"You've Been Assigned: {role_label} — Project {project_id}",
        notification_body
    )
    log.info(f"{command} complete: {assignee['name']} on {project_id}")

def handle_weekly_report(msg: dict, gmail, sheets, whitelist: dict):
    """Generates and sends the weekly closeout report to Sam."""
    sender = msg["sender_email"]
    person = whitelist.get(sender, {})
    projects = get_project_registry(sheets)
    active = [p for p in projects if p.get("Active (Y/N)") == "Y"]

    closed_this_week = []
    # A project is "closed this week" if Current Phase == Close-Out and recently updated
    # For now we flag all in Close-Out phase
    closed_this_week = [p for p in active if p.get("Current Phase") == "Close-Out"]

    report_body = (
        f"Hi {person.get('name', 'Samantha')},\n\n"
        f"Weekly Closeout Report — {date.today().strftime('%B %d, %Y')}\n\n"
        f"--- PROJECTS IN CLOSE-OUT THIS WEEK ---\n"
    )
    if closed_this_week:
        for p in closed_this_week:
            report_body += (
                f"\n  Project: {p.get('Project Name')}\n"
                f"  PM: {p.get('PM Assigned', 'TBD')}\n"
                f"  Client: {p.get('Client Name', 'TBD')}\n"
                f"  Status: Close-Out\n"
            )
    else:
        report_body += "\n  No projects currently in Close-Out.\n"

    report_body += (
        f"\n--- FULL ACTIVE PORTFOLIO ({len(active)} projects) ---\n"
    )
    for p in active:
        report_body += (
            f"\n  {p.get('Project ID')} | {p.get('Project Name')} "
            f"| Phase: {p.get('Current Phase')} "
            f"| PM: {p.get('PM Assigned', 'TBD')}\n"
        )
    report_body += "\n— Baldovin Agent"

    send_email(gmail, sender,
        f"Weekly Closeout Report — {date.today().strftime('%B %d, %Y')}",
        report_body
    )
    log.info("Weekly closeout report sent.")

# ─────────────────────────────────────────────────────────────
# DAILY REPORTS
# ─────────────────────────────────────────────────────────────
def send_sam_daily_report(gmail, sheets, whitelist: dict):
    """Sends Sam's daily 7am portfolio briefing."""
    sam = next((v for v in whitelist.values()
                if v["role"] == "Leadership" and "samantha" in v["email"].lower()), None)
    if not sam:
        return

    projects = get_project_registry(sheets)
    active = [p for p in projects if p.get("Active (Y/N)") == "Y"]

    today_str = date.today().strftime("%B %d, %Y")
    body = f"Good morning Samantha,\n\nHere is your daily portfolio summary for {today_str}.\n\n"
    body += f"ACTIVE PROJECTS — {len(active)} total\n"
    body += "─" * 40 + "\n"

    flags = []
    for p in active:
        close_date = p.get("Close Date", "")
        days_to_close = ""
        if close_date:
            try:
                delta = (datetime.strptime(close_date, "%Y-%m-%d").date() - date.today()).days
                days_to_close = f"{delta} days"
                if delta <= 30:
                    flags.append(f"  ⚠ {p.get('Project Name')} — {delta} days to close, punch list window.")
            except ValueError:
                pass

        # CompanyCam placeholder — replace with live API call when key is available
        photo_count = get_companycam_photo_count(p.get("Project ID", ""))
        photo_flag  = ""
        if photo_count == 0:
            photo_flag = "⚠ ZERO photos today"
            flags.append(f"  ⚠ {p.get('Project Name')} — No CompanyCam photos logged today.")
        elif photo_count < 3:
            photo_flag = "⚠ LOW photo activity"

        body += (
            f"\n  {p.get('Project ID')} | {p.get('Project Name')}\n"
            f"  Phase: {p.get('Current Phase', 'Unknown')}\n"
            f"  PM: {p.get('PM Assigned', 'TBD')} | "
            f"Estimator: {p.get('Estimator', 'TBD')}\n"
            f"  Location: {p.get('Folder Location', 'Upcoming')}\n"
            f"  CompanyCam Photos Today: {photo_count} {photo_flag}\n"
            f"  Days to Close: {days_to_close or 'Schedule not set'}\n"
        )

    if flags:
        body += "\n\nFLAGS & ATTENTION ITEMS\n" + "─" * 40 + "\n"
        body += "\n".join(flags)
    else:
        body += "\n\nNo flags today.\n"

    body += "\n\n— Baldovin Agent"

    send_email(gmail, sam["email"],
        f"Daily Briefing — {today_str} | {len(active)} Active Projects",
        body
    )
    log.info("Sam daily report sent.")

def get_companycam_photo_count(project_id: str) -> int:
    """
    Fetches today's photo count from CompanyCam for a given project.
    Returns 0 if API key not set or project not found.
    Full CompanyCam API integration goes here when key is provided.
    """
    if not COMPANYCAM_API_KEY:
        return 0
    # TODO: Implement CompanyCam API call
    # GET https://api.companycam.com/v2/projects/{project_id}/photos
    # Filter by created_at >= today
    return 0

def send_punchlist_warning(gmail, project: dict, days_remaining: int, pm_email: str):
    """Fires the 30-day punch list warning to the PM."""
    send_email(gmail, pm_email,
        f"Action Required — Punch List Window: {project.get('Project Name')}",
        f"Hi {project.get('PM Assigned')},\n\n"
        f"This is your 30-day punch list reminder for {project.get('Project Name')}.\n\n"
        f"Scheduled Close Date: {project.get('Close Date')}\n"
        f"Days Remaining: {days_remaining}\n\n"
        f"Now is the time to begin your punch list walkthrough and coordinate "
        f"with your subs for final completion items.\n\n"
        f"— Baldovin Agent"
    )
    log.info(f"Punch list warning sent for {project.get('Project ID')}")

# ─────────────────────────────────────────────────────────────
# PUNCH LIST COUNTDOWN CHECKER
# ─────────────────────────────────────────────────────────────
def check_punchlist_countdowns(gmail, drive, sheets, whitelist: dict):
    """
    For every active Construction-phase project, checks if we're
    within 30 days of the close date. Fires warning if so.
    """
    projects = get_project_registry(sheets)
    for proj in projects:
        if proj.get("Current Phase") != "Construction":
            continue
        if proj.get("Active (Y/N)") != "Y":
            continue
        close_date_str = proj.get("Close Date", "")
        if not close_date_str:
            continue
        try:
            close_date = datetime.strptime(close_date_str, "%Y-%m-%d").date()
            days_remaining = (close_date - date.today()).days
            if 0 < days_remaining <= 30:
                pm_name = proj.get("PM Assigned", "")
                pm_record = next(
                    (v for v in whitelist.values() if pm_name.lower() in v["name"].lower()),
                    None
                )
                if pm_record:
                    send_punchlist_warning(gmail, proj, days_remaining, pm_record["email"])
                    update_project_field(sheets, proj["Project ID"], "Current Phase", "Close-Out Prep")
        except ValueError:
            continue

# ─────────────────────────────────────────────────────────────
# FOLDER MOVE HANDLER
# ─────────────────────────────────────────────────────────────
def handle_folder_moves(gmail, drive, sheets, whitelist: dict):
    """Detects projects moved to Current and sends PM confirmation."""
    moved_projects = detect_folder_moves(drive, sheets)
    for proj in moved_projects:
        update_project_field(sheets, proj["Project ID"], "Folder Location", "Current")
        pm_name = proj.get("PM Assigned", "")
        pm_record = next(
            (v for v in whitelist.values() if pm_name.lower() in v.get("name", "").lower()),
            None
        )
        if pm_record:
            send_email(gmail, pm_record["email"],
                f"Great News: {proj['Project Name']} Is Now Current",
                f"Hey {pm_record['name']},\n\n"
                f"I noticed you moved {proj['Project Name']} to the Current folder — "
                f"that's a big step, nice work getting it there.\n\n"
                f"I've logged the move today and updated the project status. "
                f"I'll now begin monitoring your Schedule folder for the construction "
                f"timeline and will start the 30-day punch list countdown once your "
                f"close date is confirmed in the schedule.\n\n"
                f"If anything looks off or you moved it by accident, just reply and let me know.\n\n"
                f"— Baldovin Agent"
            )
        log.info(f"Folder move confirmed for {proj['Project ID']} — {proj['Project Name']}")

# ─────────────────────────────────────────────────────────────
# SCHEDULE WATCHER
# ─────────────────────────────────────────────────────────────
def check_schedule_updates(drive, sheets, whitelist: dict):
    """
    For active projects in Current folder, watches the Schedule subfolder
    for new or updated Excel files and updates the close date in registry.
    """
    # Get all project folders in Current
    current_folders = drive.files().list(
        q=f"'{DRIVE_CURRENT_FOLDER}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name, modifiedTime)"
    ).execute().get("files", [])

    projects = get_project_registry(sheets)
    proj_map = {p["Project Name"].lower(): p for p in projects}

    for folder in current_folders:
        proj = proj_map.get(folder["name"].lower())
        if not proj:
            continue
        close_date = read_schedule_close_date(drive, folder["id"])
        if close_date and close_date != proj.get("Close Date", ""):
            update_project_field(sheets, proj["Project ID"], "Close Date", close_date)
            log.info(f"Close date updated for {proj['Project ID']}: {close_date}")

# ─────────────────────────────────────────────────────────────
# SILENT ACTION LOG
# Logs emails from whitelisted senders that couldn't be actioned.
# Stored in memory for the session — future version can write to Sheet.
# ─────────────────────────────────────────────────────────────
_silent_log = []

def log_silent(sender: str, subject: str, reason: str):
    entry = {
        "timestamp": datetime.now().isoformat(),
        "sender":    sender,
        "subject":   subject,
        "reason":    reason,
    }
    _silent_log.append(entry)
    log.info(f"Silent log entry: {sender} | {subject} | {reason}")

# ─────────────────────────────────────────────────────────────
# CLAUDE EMAIL BRAIN
# Replaces rigid pattern matching with intelligent email routing.
# Claude reads the full email, knows who the sender is, what
# projects they're on, and decides what action to take.
# ─────────────────────────────────────────────────────────────
AGENT_SYSTEM_PROMPT = """
You are the Baldovin Construction Co. project agent — an intelligent assistant 
managing project communications for a second-generation family construction firm 
based in Peoria, IL, expanding into TX, AZ, and FL.

You operate within EOS/Traction principles. Every response you give is professional, 
direct, warm, and helpful. You sign all replies as "— Baldovin Agent".

YOUR CAPABILITIES:
- Create new projects (triggers Drive folder creation and project registration)
- Assign team members to projects by role (Designer, PM, Sub, Field Super)
- Look up current project status, phase, and team assignments
- Answer questions about projects the sender is assigned to
- Guide team members through what you can help with
- Generate status and closeout reports on request

YOUR RULES:
- You ONLY share information about projects the sender is assigned to
- Leadership and Director roles can see all projects
- Estimators, PMs, Designers, and Field see only their assigned projects
- If a sender asks about a project they're not on, politely tell them you can't share that
- If you need more information to complete a request, ask ONE question at a time conversationally
- Never make up project data — only report what exists in the registry
- If you cannot action something, respond with ACTION: NONE and a brief reason
- If you can action something, respond with the action type first, then your reply

RESPONSE FORMAT:
Always respond in this exact JSON format so the agent can parse your decision:
{
  "action": "NEW_PROJECT" | "ASSIGN_DESIGNER" | "ASSIGN_PM" | "ASSIGN_SUB" | "ASSIGN_SUPER" | "WEEKLY_REPORT" | "STATUS_REPLY" | "CLARIFY" | "NONE",
  "reply": "the email reply text to send to the user, or empty string if action is NONE",
  "data": {
    "project_name": "",
    "project_id": "",
    "assignee_name": "",
    "address": "",
    "client_name": "",
    "client_email": "",
    "client_phone": "",
    "scope": "",
    "estimated_value": "",
    "notes": ""
  }
}

Only include fields in data that are relevant to the action.
If action is NONE, reply must be empty string "".
If action is CLARIFY, reply contains the single clarifying question to ask.
If action is STATUS_REPLY, reply contains the status information formatted clearly.
"""

def claude_process_email(
    msg: dict,
    person: dict,
    assigned_projects: list,
    all_projects: list,
) -> dict:
    """
    Passes the email to Claude with full context.
    Claude decides what action to take and what to reply.
    Returns parsed JSON decision dict.
    """
    # Build project context scoped to sender's role
    role = person.get("role", "")
    is_leadership = role in ["Leadership", "Director Ops & BD", "Admin"]

    if is_leadership:
        visible_projects = all_projects
    else:
        visible_projects = assigned_projects

    projects_summary = []
    for p in visible_projects:
        projects_summary.append({
            "project_id":     p.get("Project ID", ""),
            "project_name":   p.get("Project Name", ""),
            "address":        p.get("Address", ""),
            "client":         p.get("Client Name", ""),
            "phase":          p.get("Current Phase", ""),
            "estimator":      p.get("Estimator", ""),
            "designer":       p.get("Designer", ""),
            "pm":             p.get("PM Assigned", ""),
            "folder":         p.get("Folder Location", ""),
            "close_date":     p.get("Close Date", ""),
            "active":         p.get("Active (Y/N)", ""),
        })

    user_prompt = f"""
SENDER INFORMATION:
Name: {person.get('name')}
Email: {msg['sender_email']}
Role: {person.get('role')}
Active: {person.get('active')}

SENDER'S VISIBLE PROJECTS:
{json.dumps(projects_summary, indent=2)}

TOTAL ACTIVE PROJECTS IN PORTFOLIO: {len(all_projects)}

EMAIL RECEIVED:
Subject: {msg['subject']}
Body:
{msg['body']}

Based on this email, decide what action to take and what to reply.
Remember: respond ONLY with valid JSON matching the format in your instructions.
"""

    raw = ask_claude(user_prompt, system=AGENT_SYSTEM_PROMPT)

    try:
        clean = raw.replace("```json", "").replace("```", "").strip()
        return json.loads(clean)
    except Exception as e:
        log.error(f"Claude response parse error: {e} | Raw: {raw}")
        return {"action": "NONE", "reply": "", "data": {}}


def get_sender_projects(sender_email: str, all_projects: list, person: dict) -> list:
    """Returns projects the sender is assigned to based on their role."""
    role = person.get("role", "")
    name = person.get("name", "").lower()

    if role in ["Leadership", "Director Ops & BD", "Admin"]:
        return all_projects

    assigned = []
    for p in all_projects:
        assigned_fields = [
            p.get("Estimator", "").lower(),
            p.get("Designer", "").lower(),
            p.get("PM Assigned", "").lower(),
            p.get("Sub Coordinator", "").lower(),
            p.get("Field Super", "").lower(),
            p.get("Finance Contact", "").lower(),
        ]
        if any(name in field for field in assigned_fields):
            assigned.append(p)
    return assigned


# ─────────────────────────────────────────────────────────────
# MAIN EMAIL PROCESSOR
# ─────────────────────────────────────────────────────────────
def process_inbox(gmail, drive, sheets):
    """
    Reads unread emails. Unknown senders are silently ignored.
    Whitelisted senders get Claude-powered intelligent handling.
    """
    whitelist     = get_whitelist(sheets)
    roster        = get_team_roster(sheets)
    all_projects  = get_project_registry(sheets)
    messages      = get_unread_messages(gmail)

    log.info(f"Inbox check — {len(messages)} unread message(s) found.")

    for msg_ref in messages:
        try:
            msg    = get_message_detail(gmail, msg_ref["id"])
            sender = msg["sender_email"]
            subject = msg["subject"]

            # ── WHITELIST GATE ──────────────────────────────
            person = whitelist.get(sender.lower())
            if not person or not person.get("active"):
                # Unknown or inactive sender — silent. Leave unread.
                log.info(f"Unknown/inactive sender: {sender} — ignored silently.")
                continue

            # Known sender — mark as read and process
            mark_as_read(gmail, msg_ref["id"])
            log.info(f"━━━ WHITELISTED EMAIL ━━━")
            log.info(f"  From: {person['name']} ({person['role']})")
            log.info(f"  Subject: {subject}")

            # ── SCOPE PROJECTS TO SENDER ────────────────────
            assigned_projects = get_sender_projects(sender, all_projects, person)

            # ── HAND TO CLAUDE ──────────────────────────────
            decision = claude_process_email(msg, person, assigned_projects, all_projects)
            action   = decision.get("action", "NONE")
            reply    = decision.get("reply", "")
            data     = decision.get("data", {})

            log.info(f"  Claude decision: {action}")

            # ── ROUTE THE ACTION ────────────────────────────
            if action == "NEW_PROJECT":
                # Merge Claude-extracted data with msg for handler
                msg["body"] = json.dumps(data) if data else msg["body"]
                handle_new_project(msg, gmail, drive, sheets, whitelist, roster)

            elif action == "ASSIGN_DESIGNER":
                handle_assignment(
                    msg, "ASSIGN DESIGNER",
                    data.get("project_id", ""),
                    data.get("assignee_name", ""),
                    gmail, sheets, whitelist, roster
                )

            elif action == "ASSIGN_PM":
                handle_assignment(
                    msg, "ASSIGN PM",
                    data.get("project_id", ""),
                    data.get("assignee_name", ""),
                    gmail, sheets, whitelist, roster
                )

            elif action == "ASSIGN_SUB":
                handle_assignment(
                    msg, "ASSIGN SUB",
                    data.get("project_id", ""),
                    data.get("assignee_name", ""),
                    gmail, sheets, whitelist, roster
                )

            elif action == "ASSIGN_SUPER":
                handle_assignment(
                    msg, "ASSIGN SUPER",
                    data.get("project_id", ""),
                    data.get("assignee_name", ""),
                    gmail, sheets, whitelist, roster
                )

            elif action == "WEEKLY_REPORT":
                handle_weekly_report(msg, gmail, sheets, whitelist)

            elif action in ("STATUS_REPLY", "CLARIFY"):
                # Claude already wrote the reply — just send it
                if reply:
                    send_email(
                        gmail, sender,
                        f"Re: {subject}",
                        reply
                    )
                    log.info(f"  Reply sent: {action}")

            elif action == "NONE":
                # Can't action — log silently, no reply
                log_silent(sender, subject, "Claude determined no actionable intent")

            else:
                log.warning(f"  Unknown action from Claude: {action}")
                log_silent(sender, subject, f"Unknown Claude action: {action}")

        except Exception as e:
            log.error(f"Error processing message {msg_ref['id']}: {e}", exc_info=True)

# ─────────────────────────────────────────────────────────────
# SCHEDULED TASKS
# Runs on each poll cycle — time-gated so they only fire once/day
# ─────────────────────────────────────────────────────────────
_last_daily_report = None
_last_punchlist_check = None
_last_schedule_check = None

def run_scheduled_tasks(gmail, drive, sheets):
    global _last_daily_report, _last_punchlist_check, _last_schedule_check

    now = datetime.now()
    today = date.today()

    # Sam daily report — 7:00 AM weekdays
    if (now.hour == 7 and now.weekday() < 5 and _last_daily_report != today):
        whitelist = get_whitelist(sheets)
        send_sam_daily_report(gmail, sheets, whitelist)
        _last_daily_report = today

    # Punch list countdown — once per day
    if _last_punchlist_check != today:
        whitelist = get_whitelist(sheets)
        check_punchlist_countdowns(gmail, drive, sheets, whitelist)
        _last_punchlist_check = today

    # Schedule file watcher — once per day
    if _last_schedule_check != today:
        whitelist = get_whitelist(sheets)
        check_schedule_updates(drive, sheets, whitelist)
        _last_schedule_check = today

    # Folder move detection — runs every poll cycle
    whitelist = get_whitelist(sheets)
    handle_folder_moves(gmail, drive, sheets, whitelist)

# ─────────────────────────────────────────────────────────────
# BUSINESS HOURS CHECK
# Agent only runs between 7am–5pm Monday–Friday (Central Time).
# Outside those hours it sleeps and checks again every 15 minutes
# in case it missed the window. No API calls are made while sleeping.
# ─────────────────────────────────────────────────────────────
BUSINESS_HOUR_START = int(os.environ.get("BUSINESS_HOUR_START", "7"))   # 7 AM
BUSINESS_HOUR_END   = int(os.environ.get("BUSINESS_HOUR_END",   "17"))  # 5 PM
SLEEP_OUTSIDE_HOURS = 900   # 15 minutes — just to recheck the clock
SLEEP_BUSINESS_HOURS = int(os.environ.get("POLL_INTERVAL_SECONDS", "3600"))  # 1 hour default

def is_business_hours() -> bool:
    """Returns True if current time is Mon–Fri, 7am–5pm Central."""
    # Railway servers run UTC — Central is UTC-5 (CST) or UTC-6 (CDT)
    # We use UTC-5 as a safe default; adjust BUSINESS_HOUR_START/END in Railway vars if needed
    from datetime import timezone, timedelta
    central = datetime.now(timezone(timedelta(hours=-5)))
    is_weekday   = central.weekday() < 5          # Mon=0 … Fri=4
    is_work_hours = BUSINESS_HOUR_START <= central.hour < BUSINESS_HOUR_END
    return is_weekday and is_work_hours

# ─────────────────────────────────────────────────────────────
# MAIN LOOP
# ─────────────────────────────────────────────────────────────
def main():
    log.info("Baldovin Construction Agent starting up...")
    gmail, drive, sheets = build_google_services()
    log.info("Google services connected.")
    log.info(f"Monitoring inbox: {AGENT_EMAIL}")
    log.info(f"Business hours: Mon–Fri {BUSINESS_HOUR_START}am–{BUSINESS_HOUR_END}pm Central")
    log.info(f"Poll interval during business hours: {SLEEP_BUSINESS_HOURS}s")

    while True:
        if not is_business_hours():
            log.info("Outside business hours — sleeping 15 minutes.")
            time.sleep(SLEEP_OUTSIDE_HOURS)
            continue

        try:
            log.info("Business hours active — running inbox and scheduled tasks.")
            process_inbox(gmail, drive, sheets)
            run_scheduled_tasks(gmail, drive, sheets)
        except Exception as e:
            log.error(f"Main loop error: {e}", exc_info=True)

        log.info(f"Cycle complete — sleeping {SLEEP_BUSINESS_HOURS // 60} minutes.")
        time.sleep(SLEEP_BUSINESS_HOURS)

if __name__ == "__main__":
    main()
