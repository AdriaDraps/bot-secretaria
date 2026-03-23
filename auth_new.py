"""
Script para regenerar el token OAuth2 con todos los scopes necesarios.

Uso:
  1. Asegúrate de tener credentials.json en esta carpeta
  2. Ejecuta: python auth_new.py
  3. Se abrirá el navegador para autorizar
  4. Copia el valor de GOOGLE_TOKEN_B64 que aparece en pantalla
  5. Actualiza esa variable en Railway → Settings → Variables
"""
from google_auth_oauthlib.flow import InstalledAppFlow
import json
import base64

SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/tasks',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
creds = flow.run_local_server(port=0)

token_data = {
    'token': creds.token,
    'refresh_token': creds.refresh_token,
    'token_uri': creds.token_uri,
    'client_id': creds.client_id,
    'client_secret': creds.client_secret,
    'scopes': list(creds.scopes),
}

b64 = base64.b64encode(json.dumps(token_data).encode()).decode()

print('\n' + '='*60)
print('Nuevo GOOGLE_TOKEN_B64 (copia todo el valor de abajo):')
print('='*60)
print(b64)
print('='*60)
print('\nPasos siguientes:')
print('  1. Ve a Railway → proyecto intuitive-integrity → servicio empowering-wisdom')
print('  2. Settings → Variables → GOOGLE_TOKEN_B64')
print('  3. Pega el valor anterior y guarda')
print('  4. El servicio se reiniciará automáticamente')
