import mysql.connector
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# Dados do banco
conn = mysql.connector.connect(host='xxx',
                               database="xxx",
                               user="xxx",
                               password="xxx")

query = '''
SELECT u.displayName, u.email, u.lastAccess, 
       u.lastPasswordModificationDate, 
       u.needChangePassword,
       u.username, u.lastPasswordValidationDate, 
       u.lastResetPasswordNotificationDate, u.costCenter, 
       u.department, u.organization, u.responsible
FROM User u
'''

# Conexão com o banco de dados
cursor = conn.cursor()
cursor.execute(query)
results = cursor.fetchall()

# Exceções de emails
excecoes = ['']

# Definição das colunas
columns = [
    'displayName', 'email', 'lastAccess', 
       'lastPasswordModificationDate', 
       'needChangePassword',
       'username', 'lastPasswordValidationDate', 
       'lastResetPasswordNotificationDate', 'costCenter', 
       'department', 'organization', 'responsible'
]

data = [list(row) for row in results]

# Filtrar as entradas usando a lista de exceções de emails
filtered_data = [row for row in data if row[1] not in excecoes] 

df = pd.DataFrame(filtered_data, columns=columns)

# Remover duplicatas
df = df.drop_duplicates(subset=['email'])

# Definir caminhos dos arquivos existentes
existing_file_path = r''
existing_file_path2 = r''

# Criar um novo workbook e obter a planilha ativa
workbook = Workbook()
sheet = workbook.active

# Preencher a planilha com os dados do DataFrame
for row in dataframe_to_rows(df, index=False, header=True):
    sheet.append(row)

# Criar uma tabela na planilha
tbl = Table(displayName='TabelaDinamica', ref=sheet.dimensions)
tbl.tableStyleInfo = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)

# Adicionar a tabela à planilha
sheet.add_table(tbl)

# Salvar o arquivo Excel nos dois caminhos especificados
workbook.save(existing_file_path)
workbook.save(existing_file_path2)

print(f'Deu certo, salvo no caminho: {existing_file_path}')

# Configurações do Sharepoint
site_url = 'xxx'
username = 'xxx'
password = 'xxx'
relative_url = 'xxx' # Caminho da biblioteca
caminho_arquivo_local = r"xxx"
nome_arquivo_sharepoint = "xxx"

# Autenticação
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(site_url, ctx_auth)

    with open(existing_file_path, 'rb') as f:
        content = f.read()

    target_folder = ctx.web.get_folder_by_server_relative_url(relative_url)
    target_file = target_folder.upload_file(nome_arquivo_sharepoint, content)
    ctx.execute_query()

    print(f"Arquivo '{nome_arquivo_sharepoint}' enviado com sucesso para o SharePoint!")
else:
    print("Erro de autenticação")