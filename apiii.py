import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\alpoente - tipo3\\Desktop\\projetoautomatizacao-d5b2f1d4925e.json", scope)
client = gspread.authorize(creds)

sheet = client.open_by_key('1pd_fM-G6Yf9e6OfkWTBl2cWKrImVmygEBoKUW62mEbQ')
worksheet = sheet.worksheet('Respostas')


all_values = worksheet.get_all_values()
col_a = [row[2] for row in all_values[1:]]
print(col_a)

for nProcesso in col_a:
    nome = "a"
    new_user = {
        "nome": nome,
        "password": 123,
        "primaryEmail": "a" + str(nProcesso) + "@alpoente.org",
        "changePasswordAtNextLogin": True,
    }
    print(new_user["primaryEmail"])
#    service.users().insert(body=new_user).execute()