import msal
import requests
import json
from datetime import datetime

print("Codespace funcionando OK")

# ======================
# CONFIG
# ======================
CLIENT_ID = "aa4a5948-53ce-42e6-b1e1-b4e37b5723b6"
TENANT_ID = "1ae0bd69-c867-4cf3-b57f-3855fe6628f1"

AUTHORITY = "https://login.microsoftonline.com/consumers"

SCOPES = [
    "User.Read",
    "Mail.Read",
    "Files.ReadWrite.All"
]

EXCEL_FILE_NAME = "EXCELMAIL.xlsx"
TABLE_NAME = "Tabla1"

# ======================
# LOGIN MSAL (Device Flow)
# ======================
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise SystemExit(flow)

print("👉 Ve a:", flow["verification_uri"])
print("👉 Código:", flow["user_code"])

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise SystemExit(result)

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}


# ======================
# MAPA DE CARPETAS
# ======================

folder_map = {}

folders_url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100"

resp_folders = requests.get(folders_url, headers=headers)

if resp_folders.status_code == 200:

    folders = resp_folders.json().get("value", [])

    for folder in folders:

        folder_map[folder["id"]] = folder["displayName"]


for k, v in folder_map.items():
    print(v)




    print("Carpetas cargadas:", len(folder_map))

else:

    print("Error obteniendo carpetas")
    print(resp_folders.status_code)
    print(resp_folders.text)


print("Access token OK")

# ======================
# OBTENER EMAILS
# ======================


url = (
    "https://graph.microsoft.com/v1.0/me/messages"
    "?$top=100"
    "&$orderby=receivedDateTime desc"
    "&$select=receivedDateTime,from,subject,bodyPreview,parentFolderId"
)

resp = requests.get(url, headers=headers)
if resp.status_code != 200:
    print(resp.text)
    raise SystemExit("Error al obtener emails")

emails = resp.json().get("value", [])
print(f"Emails obtenidos: {len(emails)}")

# ======================
# BUSCAR EL EXCEL EN ONEDRIVE
# ======================
url_drive = "https://graph.microsoft.com/v1.0/me/drive/root/children"
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

print(f"FILE ID encontrado: {file_id}")

# ======================
# PREPARAR FILAS
# ======================

rows = []

for email in emails:

    folder_id = email.get("parentFolderId")

    folder_name = folder_map.get(
        folder_id,
        "DESCONOCIDA"
    )

    fecha_original = email.get("receivedDateTime")

    fecha_formateada = datetime.strptime(
        fecha_original,
        "%Y-%m-%dT%H:%M:%SZ"
    ).strftime("%d/%m/%Y")

    rows.append([
        fecha_formateada,
        email.get("from", {}).get("emailAddress", {}).get("address"),
        email.get("subject"),
        folder_name,
        email.get("bodyPreview")
    ])

# ======================
# INSERTAR FILAS EN TABLA ONLINE
# ======================



url_add_rows = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/tables('{TABLE_NAME}')/rows/add"

payload = {"values": rows}

resp = requests.post(
    url_add_rows,
    headers={**headers, "Content-Type": "application/json"},
    json=payload
)

if resp.status_code not in (200, 201):
    print(resp.text)
    raise SystemExit("Error al insertar filas en Excel")

print("✔ Sincronización completada correctamente")
