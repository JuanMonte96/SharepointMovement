import pandas as pd
import requests
import io
from msal import ConfidentialClientApplication


# === CONFIGURACIÓN AZURE ===
# Aqui se pone el tenant id
# Aqui se pone el client id 
# Aqui se pone el client secret 
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

# === CONFIGURACIÓN SHAREPOINT ===
site_name = "prueba"
site_domain = "gramivys.sharepoint.com"
file_path_name = 'prueba.xlsx'
# https://gramivys.sharepoint.com/sites/prueba/Documentos%20compartidos/prueba.xlsx

# === OBTENER TOKEN DE ACCESSO ===
app = ConfidentialClientApplication(
    client_id, authority=authority, client_credential=client_secret
)
token_response = app.acquire_token_for_client(scopes=scope)

if "access_token" not in token_response:
    print("❌ Error al obtener token:", token_response.get("error_description"))
    exit()

access_token = token_response["access_token"]
headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

# === OBTENER EL ID DEL SITIO ===
site_url = f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/sites/{site_name}"
site_resp = requests.get(site_url, headers=headers)
site_resp.raise_for_status()
site_id = site_resp.json()["id"]

# === OBTENER EL ID DE LA BIBLIOTECA ("Documentos") ===
drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
drives_resp = requests.get(drives_url, headers=headers)
drives_resp.raise_for_status()
print(drives_resp) 

# Busca una biblioteca llamada "Documentos" o "Documents"
drives = drives_resp.json()["value"]
doc_drive = next(d for d in drives if "document" in d["name"].lower() or d["name"] == "Documents")
drive_id = doc_drive["id"]
print(drive_id)

# === DESCARGAR ARCHIVO ===
file_path = f"/{file_path_name}"
file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:{file_path}:/content"
file_resp = requests.get(file_url, headers={'Authorization': f'Bearer {access_token}'})

if file_resp.status_code != 200:
    print("❌ Error al descargar el archivo:", file_resp.status_code)
    print("➡️ Respuesta:", file_resp.text[:300])
    exit()

# === LEER EXCEL COMO DATAFRAME ===
df = pd.read_excel(io.BytesIO(file_resp.content), engine='openpyxl')
print("✅ DataFrame cargado correctamente:")
print(df.head())