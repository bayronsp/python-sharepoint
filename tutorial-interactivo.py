import streamlit as st
import msal
import requests
import json

# Configuración inicial
st.title("Tutorial para Conectar a SharePoint usando MS Graph API")

st.header("1. Configuración Inicial en Azure")
st.write("""
- Registra una aplicación en Azure AD.
- Obtén el **Client ID**, **Client Secret** y **Tenant ID** desde el portal de Azure.
- Asigna permisos para **Microsoft Graph API** con el alcance `Sites.ReadWrite.All`.
""")

st.header("2. Autenticación con MSAL")
st.code('''
client_id = "TU_CLIENT_ID"
client_secret = "TU_CLIENT_SECRET"
tenant_id = "TU_TENANT_ID"
authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    client_id, authority=authority, client_credential=client_secret
)
token_response = app.acquire_token_for_client(scopes=scopes)

access_token = token_response.get("access_token")
if not access_token:
    raise Exception("No se pudo obtener el token de acceso")
''', language="python")

# Token Authentication Example
if st.button("Autenticar y Obtener Token"):
    client_id = st.text_input("Client ID")
    client_secret = st.text_input("Client Secret", type="password")
    tenant_id = st.text_input("Tenant ID")

    if client_id and client_secret and tenant_id:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scopes = ["https://graph.microsoft.com/.default"]
        
        app = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )
        token_response = app.acquire_token_for_client(scopes=scopes)
        access_token = token_response.get("access_token")
        
        if access_token:
            st.success("Token obtenido exitosamente")
            st.code(access_token, language="text")
        else:
            st.error("Error al obtener el token")

st.header("3. Conexión a SharePoint")
st.write("""
- Obtén el **Site ID** de tu sitio de SharePoint desde la URL.
- Configura la URL de la lista:

`https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/items`
""")

st.code('''
site_id = "TU_SITE_ID"
list_name = "flota"
graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_name}/items?expand=fields"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}
''', language="python")

st.header("4. Crear Elementos en la Lista")
st.code('''
datos_vehiculo = {
    "fields": {
        "Title": "Camioneta Ford",
        "Marca": "Ford",
        "Modelo": "Ranger",
        "Matricula": "ABC-123",
        "Estado": "Operativo",
        "Fecha de ultima mantencion": "2024-10-01",
        "Tipo de combustible": "Diésel"
    }
}

response_post = requests.post(
    graph_url, headers=headers, data=json.dumps(datos_vehiculo)
)

if response_post.status_code == 201:
    print("Datos del vehículo agregados exitosamente.")
else:
    print(f"Error al agregar datos: {response_post.status_code} - {response_post.text}")
''', language="python")

st.header("5. Obtener Datos de la Lista")
st.code('''
response_get = requests.get(graph_url, headers=headers)

if response_get.status_code == 200:
    data = response_get.json()
    for item in data['value']:
        print(f"Vehículo: {item['fields']['Title']} - Marca: {item['fields']['Marca']}")
else:
    print(f"Error al obtener datos: {response_get.status_code} - {response_get.text}")
''', language="python")

# Simular resultado de datos de ejemplo
if st.button("Simular Obtener Datos"):
    ejemplo_datos = [
        {"fields": {"Title": "Camioneta Ford", "Marca": "Ford"}},
        {"fields": {"Title": "SUV Toyota", "Marca": "Toyota"}}
    ]
    for item in ejemplo_datos:
        st.write(f"Vehículo: {item['fields']['Title']} - Marca: {item['fields']['Marca']}")
