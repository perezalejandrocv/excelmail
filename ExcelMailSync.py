import msal
import requests
from datetime import datetime
import os

print("ExcelMailSync iniciado")

# ======================
# CONFIG (desde GitHub Secrets)
# ======================
CLIENT_ID = os.environ["CLIENT_ID"]
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

USER_ID = "9017f3a7-dcf2-4475-adff-cc5b4661df93"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

print("CLIENT_ID OK:", bool(CLIENT_ID))
print("TENANT_ID OK:", bool(TENANT_ID))
print("CLIENT_SECRET OK:", bool(CLIENT_SECRET))
print("AUTHORITY:", AUTHORITY)





EXCEL_FILE_NAME = "EXCELMAIL.xlsx"
TABLE_NAME = "Tabla1"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = ["https://graph.microsoft.com/.default"]

# ======================
# AUTH (CLIENT CREDENTIALS)
# ======================
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
# CARPETAS
# ======================
folder_map = {}

USER_ID = "9017f3a7-dcf2-4475-adff-cc5b4661df93"
folders_url = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/mailFolders?$top=100"
resp_folders = requests.get(folders_url, headers=headers)

if resp_folders.status_code == 200:
    folders = resp_folders.json().get("value", [])
    for f in folders:
        folder_map[f["id"]] = f["displayName"]

print("Carpetas cargadas:", len(folder_map))

# ======================
# EMAILS
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
    raise SystemExit("Error emails")

emails = resp.json().get("value", [])
print("Emails obtenidos:", len(emails))

# ======================
# ONEDRIVE FILE
# ======================
url_drive = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/root/children"
resp = requests.get(url_drive, headers=headers)

files = resp.json().get("value", [])

file_id = None
for f in files:
    if f["name"] == EXCEL_FILE_NAME:
        file_id = f["id"]
        break

if not file_id:
    raise SystemExit("Excel no encontrado")

print("FILE ID:", file_id)

# ======================
# FILAS
# ======================
rows = []

for email in emails:

    folder_id = email.get("parentFolderId")
    folder_name = folder_map.get(folder_id, "DESCONOCIDA")

    fecha_raw = email.get("receivedDateTime")
    fecha = datetime.strptime(fecha_raw, "%Y-%m-%dT%H:%M:%SZ").strftime("%d/%m/%Y")

    rows.append([
        fecha,
        email.get("from", {}).get("emailAddress", {}).get("address"),
        email.get("subject"),
        folder_name,
        email.get("bodyPreview")
    ])

# ======================
# INSERT EXCEL
# ======================
url_add = f"https://graph.microsoft.com/v1.0/users/{USER_ID}/drive/items/{file_id}/workbook/tables('{TABLE_NAME}')/rows/add"

payload = {"values": rows}

resp = requests.post(url_add, headers={**headers, "Content-Type": "application/json"}, json=payload)

if resp.status_code not in (200, 201):
    print(resp.text)
    raise SystemExit("Error Excel")

print("✔ Sincronización completada correctamente")
