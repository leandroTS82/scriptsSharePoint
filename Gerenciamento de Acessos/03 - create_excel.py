"""
create_excel.py

Este script lê arquivos CSV de um diretório especificado, processa os dados, 
e os salva em um arquivo Excel com várias abas, além de gerar um resumo de acessos 
dos usuários. 

Dependências:
- pandas: biblioteca para manipulação e análise de dados.
- openpyxl: biblioteca para leitura e escrita de arquivos Excel (XLSX).
  
Como instalar as dependências:
1. Certifique-se de ter o Python instalado em seu sistema (versão 3.6 ou superior).
2. Instale as dependências necessárias usando pip:
   pip install pandas openpyxl

Como executar:
1. Coloque os arquivos CSV no diretório especificado em 'csv_directory'.
2. Execute o script pelo terminal ou prompt de comando:
   python "03 - create_excel.py"
3. O arquivo Excel será salvo no diretório 'output_directory' com o nome 
   "Butterfly - Gerenciamento de Acessos.xlsx".
"""

import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import PatternFill, Font

# Defina o diretório onde os arquivos CSV estão localizados
csv_directory = ".\\files\\07csv"
output_directory = ".\\files\\07Excel"
output_file_path = os.path.join(output_directory, "Butterfly - Gerenciamento de Acessos.xlsx")

# Crie o diretório de saída para os arquivos Excel, caso não exista
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Lista para armazenar os dados de Nome e Email sem duplicações
unique_users = pd.DataFrame(columns=["Nome", "Email"])

# Crie um objeto ExcelWriter para salvar as várias abas
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Itere sobre cada arquivo CSV no diretório
    for csv_file in os.listdir(csv_directory):
        if csv_file.endswith(".csv"):
            # Leia o arquivo CSV em um DataFrame
            csv_path = os.path.join(csv_directory, csv_file)
            df = pd.read_csv(csv_path)

            # Remova a coluna "Login", se existir
            if "Login" in df.columns:
                df = df.drop(columns=["Login"])

            # Adiciona dados únicos de Nome e Email ao DataFrame de usuários únicos
            unique_users = pd.concat([unique_users, df[["Nome", "Email"]]]).drop_duplicates()

            # Extraia o nome do site (sem extensão) para usar como nome da aba
            sheet_name = os.path.splitext(csv_file)[0]
            table_name = sheet_name.replace(" ", "_")

            # Escreva o DataFrame em uma nova aba no arquivo Excel
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Obtenha o objeto do workbook e a aba recém-criada
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # Defina o intervalo da tabela
            last_row = len(df) + 1  # +1 para incluir o cabeçalho
            table_range = f"A1:{chr(64 + len(df.columns))}{last_row}"

            # Crie a tabela
            table = Table(displayName=table_name, ref=table_range)
            style = TableStyleInfo(
                name="TableStyleMedium9", showFirstColumn=False,
                showLastColumn=False, showRowStripes=True,
                showColumnStripes=True
            )
            table.tableStyleInfo = style
            worksheet.add_table(table)

            print(f"Aba '{sheet_name}' criada com sucesso como tabela.")

    # Adiciona a aba de usuários únicos
    unique_users.to_excel(writer, sheet_name="Usuarios Unicos", index=False)
    worksheet = writer.sheets["Usuarios Unicos"]
    last_row = len(unique_users) + 1
    table_range = f"A1:B{last_row}"
    table = Table(displayName="UsuariosUnicos", ref=table_range)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
                           showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style
    worksheet.add_table(table)

    print("Aba 'Usuarios Unicos' criada com sucesso com dados únicos de Nome e Email.")

# Após salvar, agora é seguro ler o arquivo para adicionar a aba de Resumo de Acessos
with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a') as writer:
    all_users = set()

    # Lê o arquivo salvo e cria o resumo de acessos
    for sheet_name, sheet_data in pd.read_excel(output_file_path, sheet_name=None, engine='openpyxl').items():
        if 'Nome' in sheet_data.columns:
            users = sheet_data['Nome'].drop_duplicates().tolist()
            all_users.update(users)

    summary_data = {'Nome': list(all_users)}
    for sheet_name in pd.read_excel(output_file_path, sheet_name=None, engine='openpyxl').keys():
        summary_data[sheet_name] = [0] * len(all_users)

    for sheet_name, sheet_data in pd.read_excel(output_file_path, sheet_name=None, engine='openpyxl').items():
        if 'Nome' in sheet_data.columns:
            for user in all_users:
                if user in sheet_data['Nome'].values:
                    index = summary_data['Nome'].index(user)
                    summary_data[sheet_name][index] = 1

    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Resumo de Acessos', index=False)
    worksheet = writer.sheets['Resumo de Acessos']
    last_row = len(summary_df) + 1
    table_range = f"A1:{chr(64 + len(summary_df.columns))}{last_row}"
    table = Table(displayName="ResumoAcessos", ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True,
        showColumnStripes=True
    )
    table.tableStyleInfo = style
    worksheet.add_table(table)

    print("Aba 'Resumo de Acessos' criada com sucesso e movida para a primeira posição!")