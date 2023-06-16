import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define as credenciais do Google Sheets API
scope = ['https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\alpoente - tipo3\\Desktop\\projetoautomatizacao-d5b2f1d4925e.json", scope)

# Autentica a conexão com a API
client = gspread.authorize(creds)

# Abre uma planilha pelo ID
sheet = client.open_by_key('1pd_fM-G6Yf9e6OfkWTBl2cWKrImVmygEBoKUW62mEbQ')

# Seleciona uma determinada planilha pelo nome
worksheet = sheet.worksheet('Folha')

# Lê os valores da planilha em uma determinada faixa de células
values = worksheet.get('c2:c4')
values = sorted(values)
# Imprime os valores lidos
print(values)