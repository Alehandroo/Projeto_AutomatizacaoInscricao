import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from googleapiclient.discovery import build

jsonFilePath = 'C:\\Users\\JOHNNYPENI\\Desktop\\Estagio\\Projeto_AutomatizacaoInscricao\\projectoauto-ae51294fd899.json'

def get_credentials():
    service_account_file = jsonFilePath
    sub_email = 'admin@alpoente.org'

    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    credentials = credentials.with_subject(sub_email)
    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())

    return credentials


def adicionar(aluno):
    group_key = aluno["list"]
    member_email = aluno["primaryEmail"]

    try:
        member = {
            'email': member_email,
            'role': 'MEMBER'
        }

        service.members().insert(groupKey=group_key, body=member).execute()
        print(f'O membro {member_email} foi adicionado à lista {group_key} com sucesso!')
        exportar(aluno,group_key)
    except Exception as e:
        print(f'Ocorreu um erro ao adicionar o membro {member_email}: {str(e)}')

def retirar(aluno):
    member_email = aluno["primaryEmail"]
    list_to_remove = aluno["list"]

    # Obtem as listas de e-mails às quais o membro pertence
    response = service.groups().list(
        userKey=member_email
    ).execute()

    # Extrai as informações relevantes das listas de e-mails
    mailing_lists = response.get('groups', [])
    if mailing_lists:
        print('Listas de e-mails às quais o membro pertence:')
        for mailing_list in mailing_lists:
            print('Nome:', mailing_list['name'])
            print('---')
            try:
                service.members().delete(memberKey=member_email, groupKey=mailing_list['email']).execute()
                print(f'O membro {member_email} foi retirado da lista ' + mailing_list['email'] + ' com sucesso!')
                exportar(aluno, mailing_list['email'])
            except Exception as e:
                print(f'Ocorreu um erro ao retirar o membro {member_email}: {str(e)}')
    else:
        print('O membro não pertence a nenhuma lista de e-mails.')

def alterar(aluno):
    retirar(aluno)
    adicionar(aluno)

def exportar(aluno,lista):
    spreadsheet_key = '1lhp_olC3TTMrSiX34CH1ePlkJUBmQweZS4-aNJNcLgI'
    worksheet_name = 'log_book'
    spreadsheet = client.open_by_key(spreadsheet_key)
    worksheet1 = spreadsheet.worksheet(worksheet_name)
    last_row = len(worksheet1.get_all_values()) + 1
    data_execucao = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Obtém a data e hora atual
    input1 = [[data_execucao, aluno["nProcesso"], lista, aluno["acao"]]]
    worksheet1.insert_row(input1[0], index=last_row, value_input_option='RAW')

def limpar():
    worksheet.clear()
    worksheet.append_row(headerDB)

headerDB = ["Carimbo de data/hora","Endereço de Email",	"Ação",	"Nº Processo","	Tipo de Curso",	"Ano Turma Regular", "Ano Turma Profissional", "Ano Turma CEF"]

# Credenciais para as contas de serviço (API)
scopes = ['https://www.googleapis.com/auth/admin.directory.group', 'https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name(jsonFilePath, scopes)

# Linhas necessárias para abrir o google sheets
client = gspread.authorize(creds)
sheet = client.open_by_key('1mEwN7mD1gxVVuiV8CdFOUYlsxlTPJrUCUqb4VbQgwGQ')
worksheet = sheet.worksheet('DataBase')

credentials = get_credentials()
service = build('admin', 'directory_v1', credentials=credentials)

values = worksheet.get_all_values()

for linha in values[1:]:
    if not any(linha):
        continue  # Parar se encontrar uma linha vazia

    receiver_email = linha[1]
    acao = linha[2]
    nProcesso = linha[3]
    tipo = linha[4]

    if tipo == "Ensino Regular":
        anoturma = linha[5]
    elif tipo == 'Curso Profissional':
        anoturma = linha[6]
    elif tipo == "CEF":
        anoturma = linha[7]
    aluno = {
        "acao": acao,
        "anoturma": anoturma,
        "nProcesso": nProcesso,
        "primaryEmail": "a" + str(nProcesso) + "@alpoente.org",
        "list": "alunos-" + (str(anoturma).lower()) + "@alpoente.org"}
    # Switch case para escolher a ação
    match acao:
        case "Adicionar":
            adicionar(aluno)
        case "Retirar":
            retirar(aluno)
        case "Alterar":
            alterar(aluno)
        case Default:
            print("Operação inválida")