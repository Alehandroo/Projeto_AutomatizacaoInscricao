import os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Defina o ID do arquivo CSV do Google Drive
file_id = 'SEU_ID_DE_ARQUIVO_AQUI'

# Defina o caminho onde o arquivo CSV será salvo
save_path = 'C:\\Users\\alpoente - tipo3\\Desktop\\Create with code'

# Configure as credenciais de acesso
creds = None
if os.path.exists('credentials.json'):
    creds = Credentials.from_authorized_user_file('credentials.json')

# Crie a instância do cliente da API do Google Drive
service = build('drive', 'v3', credentials=creds)

# Exporte o arquivo CSV para o seu computador
request = service.files().export_media(fileId=file_id, mimeType='text/csv')
with open(save_path, 'wb') as f:
    f.write(request.execute())