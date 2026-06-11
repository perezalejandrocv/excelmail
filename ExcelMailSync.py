import msal
import requests
from datetime import datetime

print("ExcelMailSync iniciado (Delegated)")

# ======================
# CONFIG
# ======================
CLIENT_ID = "<CLIENT_ID>"  # o desde env
TENANT_ID = "consumers"    # o tu tenant
SCOPES = ["User.Read", "Mail.Read", "Files.ReadWrite.All"]

EXCEL_FILE_NAME = "EXCELMAIL.xlsx"
TABLE_NAME = "Tabla1"

# ======================
# LOGIN MSAL (Device Flow)
# ======================
app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}"
)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise SystemExit(flow)

print("👉 Ve a:", flow["verification_uri"])
print("👉 Código:", flow["user_code"])

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise SystemExit("Error obteniendo token")

token = result["access_token"]
headers = {"Authorization": f"Bearer {token}"}

print("Access token OK")

# ======================
# OBTENER CARPETAS
# ======================
folder_map = {}
folders_url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100"

resp = requests.get(folders_url, headers=headers)
if resp.status_code != 200:
    print("Error carpetas:", resp.status_code, resp.text)
    raise SystemExit()

for f in resp.json().get("value", []):
    folder_map[f["id"]] = f["displayName"]

print("Carpetas cargadas:", len(folder_map))

# ======================
# OBTENER EMAILS
# ======================
url_emails = (
    "https://graph.microsoft.com/v1.0/me/messages"
    "?$top=100"
    "&$orderby=receivedDateTime desc"
    "&$select=receivedDateTime,from,subject,bodyPreview,parentFolderId"
)

resp = requests.get(url_emails, headers=headers)
if resp.status_code != 200:
    print("Error emails:", resp.status_code, resp.text)
    raise SystemExit()

emails = resp.json().get("value", [])
print(f"Emails obtenidos: {len(emails)}")

# ======================
# BUSCAR EXCEL EN ONEDRIVE
# ======================
url_drive = "https://graph.microsoft.com/v1.0/me/drive/root/children"
resp = requests.get(url_drive, headers=headers)
if resp.status_code != 200:
    print(resp.text)
    raise SystemExit("Error listando OneDrive")

files = resp.json().get("value", [])
file_id = next((f["id"] for f in files if f["name"] == EXCEL_FILE_NAME), None)
if not file_id:
    raise SystemExit(f"{EXCEL_FILE_NAME} no encontrado")

print(f"FILE ID: {file_id}")

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
# INSERTAR FILAS EN TABLA
# ======================
url_add = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/tables('{TABLE_NAME}')/rows/add"
resp = requests.post(url_add, headers={**headers, "Content-Type": "application/json"}, json={"values": rows})
if resp.status_code not in (200, 201):
    print(resp.text)
    raise SystemExit("Error insertando filas en Excel")

print("✔ Sincronización completada correctamente")
