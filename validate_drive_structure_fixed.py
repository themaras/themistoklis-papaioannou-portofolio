
import re
from datetime import datetime
from openpyxl import Workbook
from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
SERVICE_ACCOUNT_FILE = "loco-translate-404819-b2d4f374d7e2.json"
FOLDER_ID = "1tAlVuQphn_E3ESYugK3BOXDS69eMFkWi"

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
drive = build("drive", "v3", credentials=credentials)

def list_folder(folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    results = drive.files().list(q=query, fields="files(id, name, mimeType)").execute()
    return results.get("files", [])

def extract_order_number(name: str):
    match = re.match(r"^(\d+)_", name)
    return int(match.group(1)) if match else None

def starts_with_invalid_separator(name: str):
    return not re.match(r"^\d+_.*", name)

def validate_course(course_name, course_id):
    print(f"\n🎓 Validating Course: {course_name}")
    items = list_folder(course_id)
    sections = {item['name']: item['id'] for item in items if item['mimeType'] == 'application/vnd.google-apps.folder'}
    files = [item['name'] for item in items if item['mimeType'] != 'application/vnd.google-apps.folder']
    if not sections and not files:
        ws.append([course_name, "—", "—", "WARNING", "Course folder is empty"])
        return
    for f in files:
        ws.append([course_name, "—", f, "INVALID", "Top-level file is not a section folder"])

    section_orders = []
    for section in sections:
        if section.lower() == "metadata":
            continue
        if starts_with_invalid_separator(section):
            ws.append([course_name, section, "—", "INVALID", "Section name must use '_' as separator"])
        order = extract_order_number(section)
        if order:
            section_orders.append(order)
    if section_orders:
        if min(section_orders) != 1:
            ws.append([course_name, "—", "—", "INVALID", "Section order must start from 1"])
        for i in range(1, max(section_orders) + 1):
            if i not in section_orders:
                ws.append([course_name, "—", "—", "INVALID", f"Missing expected section order: {i}"])

    for section, sec_id in sections.items():
        if section.lower() == "metadata":
            continue
        print(f"  📁 Section: {section}")
        content = list_folder(sec_id)
        filenames = [f['name'] for f in content]
        file_orders = [extract_order_number(name) for name in filenames if extract_order_number(name)]
        if file_orders:
            if min(file_orders) != 1:
                ws.append([course_name, section, "—", "INVALID", "File order must start from 1"])
            for i in range(1, max(file_orders)+1):
                if i not in file_orders:
                    ws.append([course_name, section, "—", "INVALID", f"Missing expected file order: {i}"])
        for f in content:
            name = f['name']
            if starts_with_invalid_separator(name):
                ws.append([course_name, section, name, "INVALID", "File name must use '_' as separator"])
            elif extract_order_number(name):
                ws.append([course_name, section, name, "VALID", "✓ File order correct"])
            else:
                ws.append([course_name, section, name, "INVALID", "File name missing order prefix"])

def check_metadata(course_name, course_id):
    items = list_folder(course_id)
    metadata = next((item for item in items if item['name'].lower() == "metadata" and item['mimeType'] == 'application/vnd.google-apps.folder'), None)
    if not metadata:
        return
    meta_id = metadata['id']
    meta_items = list_folder(meta_id)
    settings_file = next((item for item in meta_items if 'settings' in item['name'].lower()), None)
    attachments_folder = next((item for item in meta_items if item['name'].lower() == "attachments" and item['mimeType'] == 'application/vnd.google-apps.folder'), None)

    if settings_file:
        ws.append([course_name, "metadata", settings_file['name'], "VALID", "✓ Valid settings file"])
    if attachments_folder:
        attachment_items = list_folder(attachments_folder['id'])
        for f in attachment_items:
            ws.append([course_name, "attachments", f['name'], "VALID", "Valid digital download"])

    if settings_file and attachments_folder:
        ws.append([course_name, "attachments", attachments_folder['name'], "VALID", f"settings file found: {settings_file['name']}, with valid attachments"])
    elif settings_file:
        ws.append([course_name, "metadata", settings_file['name'], "VALID", f"Settings file found: {settings_file['name']}, no attachments folder"])
    elif attachments_folder:
        ws.append([course_name, "attachments", attachments_folder['name'], "INVALID", "Attachments folder present but no settings file"])
    else:
        ws.append([course_name, "metadata", "—", "INVALID", "metadata/ folder found but missing both settings.* and attachments/"])

print("📦 Connecting to Google Drive and preparing validation...")
courses = list_folder(FOLDER_ID)

wb = Workbook()
ws = wb.active
ws.append(["Course", "Section", "File", "Status", "Reason"])

for item in courses:
    if item["mimeType"] == "application/vnd.google-apps.folder":
        validate_course(item["name"], item["id"])
        check_metadata(item["name"], item["id"])
    else:
        ws.append([item["name"], "—", "—", "INVALID", "Top-level item is not a folder — skipping."])

output_filename = f"validation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
wb.save(output_filename)
print(f"\n✅ Validation complete. Report saved as: {output_filename}")
