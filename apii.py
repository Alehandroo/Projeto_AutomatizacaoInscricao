import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from googleapiclient.discovery import build

scope = ['https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\JOHNNYPENI\\Desktop\\Estagio\\Projeto_AutomatizacaoInscricao\\projetoautomatizacao-d5b2f1d4925e.json", scope)
client = gspread.authorize(creds)

sheet = client.open_by_key('1pd_fM-G6Yf9e6OfkWTBl2cWKrImVmygEBoKUW62mEbQ')
worksheet = sheet.worksheet('Respostas')


for i in range(2, worksheet.row_count + 1):
    linha = worksheet.row_values(i)
    if all(value == '' for value in linha):
        break
    acao,nProcesso,nome,ano,turma = linha[1:6]
    new_user = {
        "acao": acao,
        "nome": nome,
        "password": 123,
        "primaryEmail": "a" + str(nProcesso) + "@alpoente.org",
        "changePasswordAtNextLogin": True,
    }
    print(new_user["primaryEmail"])
    print(new_user)
#    service.users().insert(body=new_user).execute()
