import csv
import os.path

import gspread
from gspread.spreadsheet import Spreadsheet

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

BASE_FOLDER = "/"
CREDENTIALS_FILE=f'{BASE_FOLDER}/acesso/credentials.json'
TOKEN_FILE=f'{BASE_FOLDER}/acesso/token.json'

TEMPLATE_FOLDER_ID = f'1WrriGQaE5cScXL8UK16FD-fCJogUcls4'
TEMPLATE_ID = f'17uf2kDas1SHk_HJn3tM6gevVTjt0ZX9aBaVafmAtodM'

CSV_FILE = f'{BASE_FOLDER}/acoes.csv'

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

def get_credentials():
    credentials = None
    if os.path.exists(TOKEN_FILE):
        credentials = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            credentials = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(credentials.to_json())
    return credentials

def main():
    client = gspread.authorize(get_credentials())

    planilha_nova = 'Versão nova'    
    print(f"Verificando se planilha {planilha_nova} existe")
    try:
        sh = client.open(planilha_nova, TEMPLATE_FOLDER_ID)
        print(f"Panilha {planilha_nova} já existe e será removida")
        client.del_spreadsheet(sh.id)
    except:
        print(f"Panilha {planilha_nova} ainda não existe")

    print(f"Copiando a template para: {planilha_nova}")
    client.copy(TEMPLATE_ID, title=planilha_nova, copy_permissions=True)

    print(f"Atualizando CSV: {CSV_FILE}")
    spreadsheet: Spreadsheet = client.open(planilha_nova)

    with open(CSV_FILE, 'r') as file:
        spreadsheet.values_update(
            'Ações',
            params={'valueInputOption': 'USER_ENTERED'},
            body={'values': list(csv.reader(file))}
        )

    print(f"Formatando Cabecalho")
    aba_acoes = spreadsheet.get_worksheet(1)
    aba_acoes.format('A1:F1', {
        'horizontalAlignment': 'CENTER',
        'textFormat': {'bold': True}
        })
    

if __name__ == '__main__':
    main()