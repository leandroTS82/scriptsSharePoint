import os
import json
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# Configurações do SharePoint
site_url = "https://butterflygrowth.sharepoint.com"
username = "leandro.silva@ButterflyGrowth.onmicrosoft.com"  # Substitua pelo seu email
password = "DlA685947"  # Substitua pela sua senha

# Diretório para salvar os arquivos JSON
output_directory = "./files/07json"

# Crie o diretório de saída, caso ele não exista
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

try:
    # Usar UserCredential para autenticação
    credentials = UserCredential(username, password)
    ctx = ClientContext(site_url).with_credentials(credentials)
    
    # Obtenha informações do site
    web = ctx.web.get().execute_query()
    print(f"Conectado ao site: {web.title}, URL: {web.url}")

except Exception as e:
    print(f"Erro ao conectar ao SharePoint: {e}")
