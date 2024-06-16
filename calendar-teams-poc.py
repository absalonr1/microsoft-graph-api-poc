import msal
import requests

client_id = 'xx'
client_secret = 'xx'
tenant_id = 'xx'

# Crear una instancia de la aplicación de cliente pública
app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}")

# Obtener el token de acceso
# Si es la primera vez, se redirigirá al usuario a iniciar sesión
flow = app.initiate_device_flow(scopes=["User.Read", "Calendars.Read","Group.Read.All","Team.ReadBasic.All"])
if "user_code" not in flow:
    raise ValueError("No se pudo crear el flujo de dispositivo. ¿Están los permisos correctos configurados?")

print(flow["message"])  # Muestra el mensaje al usuario para iniciar sesión

# Esperar a que el usuario complete el flujo
result = app.acquire_token_by_device_flow(flow)
print(result)
if "access_token" in result:
    access_token = result["access_token"]
    print("Token de acceso obtenido con éxito.")
else:
    print("Error al obtener el token de acceso.")
    print(result)

# authority = f'https://login.microsoftonline.com/{tenant_id}'
# scopes = ['https://graph.microsoft.com/.default']

# app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)

# result = app.acquire_token_for_client(scopes=scopes)

# access_token = ""

# if 'access_token' in result:
#     access_token = result['access_token']
#     print(access_token)
# else:
#     print("Error al obtener el token de acceso.")
#     print(result)


team_name = 'Subgerencia Servicios TI'
url = f'https://graph.microsoft.com/v1.0/groups?$filter=displayName eq \'{team_name}\''
headers = {
    'Authorization': f'Bearer {access_token}'
}

response = requests.get(url, headers=headers)

group_id = ""

if response.status_code == 200:
    groups = response.json()
    if groups['value']:
        group_id = groups['value'][0]['id']
        print(f"ID del grupo: {group_id}")
    else:
        print("No se encontró el grupo.")
else:
    print(f"Error al obtener el ID del grupo: {response.status_code}")
    print(response.json())

headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Obtener los canales del equipo
response = requests.get(f'https://graph.microsoft.com/v1.0/teams/{group_id}/channels', headers=headers)

if response.status_code == 200:
    channels = response.json()
    for channel in channels['value']:
        print(f"Canal: {channel['displayName']} - ID: {channel['id']}")
else:
    print(f"Error al obtener los canales: {response.status_code}")
    print(response.json())


url = f'https://graph.microsoft.com/v1.0/groups/xx/events'
headers = {
    'Authorization': f'Bearer {access_token}'
}

response = requests.get(url, headers=headers)

print(response.json())

while True:
    if response.status_code == 200:
        events = response.json()
        for event in events['value']:
            print(f"Evento: {event['subject']}")
            print(f"Inicio: {event['start']['dateTime']}")
            print(f"Fin: {event['end']['dateTime']}")
            print("-" * 20)
        nextLink = events['@odata.nextLink']
        if nextLink is None:
            break
        else:
            response = requests.get(nextLink, headers=headers)
    else:
        print(f"Error al obtener los eventos: {response.status_code}")
        print(response.json())
