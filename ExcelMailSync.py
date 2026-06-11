import msal
import requests
from datetime import datetime
import os

print("ExcelMailSync iniciado (App-Only)")

# ======================
# CONFIG desde GitHub Secrets
# ======================
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID = os.environ["TENANT_ID"]

# Mailbox que quieres leer
USER_ID = os.environ.get("USER_ID")  # El Object ID del usuario o su email principal

EXCEL_FILE_NAME = "EXCELMAIL.xlsx"
TABLE_NAME = "Tabla1"

# ======================
# AUTH (Client Credentials)
# ======================
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_for_client(scopes=SCOPES)
if "access_token" not in result:
    raise SystemExit(result)

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

print("Access token OK")

# ======================
# MAPA DE CARPETAS
# ======================
def get_all_mailfolders(user_id, headers):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders?$top=100"
    folder_map = {}
    
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print("MAILFOLDERS STATUS:", resp.status_code)
            print("MAILFOLDERS RESPONSE:", resp.text)
            resp.raise_for_status()
        data = resp.json()
        for f in data.get("value", []):
            folder_map[f["id"]] = f["displayName"]
            # Recursivo si hay subcarpetas
            if f.get("childFolderCount", 0) > 0:
                folder_map.update(get_all_mailfolders_recursive(user_id, f["id"], headers))
        url = data.get("@odata.nextLink")
    return folder_map

def get_all_mailfolders_recursive(user_id, parent_id, headers):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders/{parent_id}/childFolders?$top=100"
    folder_map = {}
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            print("MAILFOLDERS RECURSIVE STATUS:", resp.status_code)
            print("MAILFOLDERS RECURSIVE RESPONSE:", resp.text)
            resp.raise_for_status()
        data = resp.json()
        for f in data.get("value", []):
            folder_map[f["id"]] = f["displayName"]
            if f.get("childFolderCount", 0) > 0:
                folder_map.update(get_all_mailfolders_recursive(user_id, f["id"], headers))
        url = data.get("@odata.nextLink")
    return folder_map

# ======================
# OBTENER CARPETAS
# ======================
folder_map = get_all_mailfolders(USER_ID, headers)
print("Carpetas obtenidas:", len(folder_map))

# ======================
# OBTENER EMAILS
# ======================
url = (
    f"https://graph.microsoft.com/v1.0/users/{USER_ID}/messages"
    "?$top=100"
    "&$orderby=receivedDateTime desc"
    "&$select=receivedDateTime,from,subject,bodyPreview,parentFolderId"
)
resp = requests.get(url, headers=headers)
if resp.status_code != 200:
    print(resp.text)
    raise SystemExit("Error al obtener emails")

emails = resp.json().get("value", [])
print("Emails obtenidos:", len(emails))

# ======================
# UBICAR EXCEL EN ONEDRIVE
# ======================
url_drive = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root/children"
resp = requests.get(url_drive, headers=headers)
if resp.status_code != 200:
    print(resp.text)
    raise SystemExit("Error al listar archivos OneDrive")

files = resp.json().get("value", [])
file_id = None
for f in files:
    if f["name"] == EXCEL_FILE_NAME:
        file_id = f["id"]
        break

if not file_id:
    raise SystemExit(f"{EXCEL_FILE_NAME} no encontrado en OneDrive")

print("FILE ID encontrado:", file_id)

# ======================
# PREPARAR FILAS
# ======================
rows = []
for email in emails:
    folder_id = email.get("parentFolderId")
    folder_name = folder_map.get(folder_id, "DESCONOCIDA")
    fecha = datetime.strptime(email.get("receivedDateTime"), "%Y-%m-%dT%H:%M:%SZ").strftime("%d/%m/%Y")
    rows.append([
        fecha,
        email.get("from", {}).get("emailAddress", {}).get("address"),
        email.get("subject"),
        folder_name,
        email.get("bodyPreview")
    ])

# ======================
# INSERTAR FILAS EN EXCEL
# ======================
url_add = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/tables('{TABLE_NAME}')/rows/add"
payload = {"values": rows}

resp = requests.post(url_add, headers={**headers, "Content-Type": "application/json"}, json=payload)
if resp.status_code not in (200, 201):
    print(resp.text)
    raise SystemExit("Error al insertar filas en Excel")

print("✔ Sincronización completada correctamente")
