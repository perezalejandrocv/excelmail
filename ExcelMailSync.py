import msal
import requests
from datetime import datetime
import os

print("ExcelMailSync iniciado (App-Only)")

# ======================
# CONFIG desde GitHub Secrets
# ======================
CLIENT_ID = os.environ["CLIENT_ID"]
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
USER_ID = os.environ["USER_ID"]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

EXCEL_FILE_NAME = "EXCELMAIL.xlsx"
TABLE_NAME = "Tabla1"

# ======================
# AUTH (Client Credentials)
# ======================
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_for_client(scopes=SCOPES)

if "access_token" not in result:
    raise SystemExit(f"Error obteniendo token: {result}")

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}
print("Access token OK")

# ======================
# OBTENER CARPETAS
# ======================
def get_all_mailfolders(user_id, headers):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders?$top=100"
    folder_map = {}
    while url:
        resp = requests.get(url, headers=headers)
        print("MAILFOLDERS STATUS:", resp.status_code)
        if resp.status_code != 200:
            print("MAILFOLDERS RESPONSE:", resp.text)
            resp.raise_for_status()
        data = resp.json()
        for f in data.get("value", []):
            folder_map[f["id"]] = f["displayName"]
            if f.get("childFolderCount", 0) > 0:
                child_map = get_all_mailfolders_recursive(user_id, f["id"], headers)
                folder_map.update(child_map)
        url = data.get("@odata.nextLink")
    return folder_map

def get_all_mailfolders_recursive(user_id, parent_id, headers):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders/{parent_id}/childFolders?$top=100"
    folder_map = {}
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            resp.raise_for_status()
        data = resp.json()
        for f in data.get("value", []):
            folder_map[f["id"]] = f["displayName"]
            if f.get("childFolderCount", 0) > 0:
                child_map = get_all_mailfolders_recursive(user_id, f["id"], headers)
                folder_map.update(child_map)
        url = data.get("@odata.nextLink")
    return folder_map

folder_map = get_all_mailfolders(USER_ID, headers)
print("Carpetas obtenidas:", len(folder_map))

# ======================
# OBTENER EMAILS
# ======================
emails_url = (
    f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages"
    "?$top=100"
    "&$orderby=receivedDateTime desc"
    "&$select=receivedDateTime,from,subject,bodyPreview,parentFolderId"
)
resp = requests.get(emails_url, headers=headers)
if resp.status_code != 200:
    print(resp.text)
    raise SystemExit("Error al obtener emails")
emails = resp.json().get("value", [])
print("Emails obtenidos:", len(emails))

# ======================
# LOCALIZAR EXCEL EN ONEDRIVE
# ======================
drive_url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root/children"
resp = requests.get(drive_url, headers=headers)
if resp.status_code != 200:
    print(resp.text)
    raise SystemExit("Error al listar archivos OneDrive")
files = resp.json().get("value", [])
file_id = next((f["id"] for f in files if f["name"] == EXCEL_FILE_NAME), None)
if not file_id:
    raise SystemExit(f"{EXCEL_FILE_NAME} no encontrado")
print("FILE ID encontrado:", file_id)

# ======================
# PREPARAR FILAS
# ======================
rows = []
for email in emails:
    folder_name = folder_map.get(email.get("parentFolderId"), "DESCONOCIDA")
    fecha = datetime.strptime(email.get("receivedDateTime"), "%Y-%m-%dT%H:%M:%SZ").strftime("%d/%m/%Y")
    rows.append([
        fecha,
        email.get("from", {}).get("emailAddress", {}).get("address"),
        email.get("subject"),
        folder_name,
        email.get("bodyPreview")
    ])

# ======================
# INSERTAR FILAS EN EXCEL ONLINE
# ======================
url_add_rows = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/tables('{TABLE_NAME}')/rows/add"
resp = requests.post(url_add_rows, headers={**headers, "Content-Type": "application/json"}, json={"values": rows})
if resp.status_code not in (200, 201):
    print(resp.text)
    raise SystemExit("Error al insertar filas en Excel")

print("✔ Sincronización completada correctamente")
