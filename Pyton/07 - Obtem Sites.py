import os
import json
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

# Configurações do SharePoint
site_url = "***"
username = "leandro.silva@ButterflyGrowth.onmicrosoft.com"  # Substitua pelo seu email
password = "****"               # Substitua pela sua senha

# Diretório para salvar os arquivos JSON
output_directory = "./files/07json"

# Crie o diretório de saída, caso ele não exista
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

try:
    # Conectar ao SharePoint
    ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

    # Obtenha todos os sites da coleção
    sites = ctx.web.site_collection.get_sites().execute_query()
    
    # Crie um array para armazenar os dados dos sites
    site_list = []

    # Preencha o array com os dados dos sites
    for site in sites:
        site_list.append({
            "Nome": site.title,
            "Url": site.url
        })

    # Converta a lista de sites para JSON e salve em um arquivo
    json_file_path = os.path.join(output_directory, "07 - SharePointSites.json")
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(site_list, json_file, ensure_ascii=False, indent=4)

    print("Arquivo JSON '07 - SharePointSites.json' criado com sucesso contendo todos os sites da coleção.")

except Exception as e:
    print(f"Erro ao obter informações da coleção de sites do SharePoint: {e}")
