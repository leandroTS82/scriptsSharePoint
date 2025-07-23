# Pré-requisitos para executar o script `ImportItemsToList.ps1`

## ✅ 1. Sistema Operacional
- Windows 10 ou superior
- PowerShell 5.x ou PowerShell 7+

## ✅ 2. Permissões necessárias
- Conta com acesso ao Microsoft 365 (Azure AD)
- Permissão de leitura e escrita na lista SharePoint Online

## ✅ 3. Instalação dos módulos PowerShell obrigatórios

Execute os comandos abaixo no PowerShell (como Administrador ou com `-Scope CurrentUser`):

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

- `PnP.PowerShell`: utilizado para autenticação e operações com SharePoint
- `ImportExcel`: permite ler arquivos `.xlsx` sem precisar do Excel instalado

## ✅ 4. Estrutura esperada de arquivos

```plaintext
ImportItemsToList
│
├── Config
│   ├── IgnoreColumns.json              # Colunas a ignorar no envio
│   └── Sites.json                      # Lista de sites disponíveis
│
├── Excel
│   └── DiarioBordo.xlsx                # Planilha com os dados a importar
│
├── 01 - ImportItemsToList.ps1          # Script principal
├── Instructions.md                     # Instruções e documentação
└── SchemaListColumns.json              # Schema com colunas da lista (sem listId)
```

## ✅ 5. Formato dos arquivos JSON

### `Sites.json`

```json
[
  {
    "name": "Site Leandro Grupo Teste",
    "url": "https://butterflygrowth.sharepoint.com/sites/leandrogrupoteste"
  }
]
```

### `SchemaListColumns.json`

```json
{
  "Colunas": [
    {
      "Title": "Data Início",
      "InternalName": "DataIn_x00ed_cio",
      "FieldTypeKind": 4
    }
    // outras colunas...
  ]
}
```

### `IgnoreColumns.json`

```json
[
  "LinkTitle",
  "ID",
  "ContentType",
  "Author"
  // outras colunas ignoradas...
]
```

## ✅ 6. Requisitos de execução

- O Excel deve ter cabeçalhos na primeira linha e não deve conter células mescladas
- Execute o script no PowerShell (não no ISE) para garantir que o `Read-Host` funcione corretamente
- A conexão com o SharePoint será feita via `-UseWebLogin` (janela interativa)

---
# 📝 Passo a passo para executar o script de importação no SharePoint

Este guia mostra como executar o script que importa dados de uma planilha Excel para uma lista do SharePoint Online.

---

## ✅ Pré-requisitos (fazer apenas na primeira vez)

1. **Abra o PowerShell como administrador**

   Clique com o botão direito no menu Iniciar → **Windows PowerShell (Admin)**

2. **Instale os módulos necessários**

   Copie e cole os comandos abaixo no PowerShell e pressione Enter:

   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   Install-Module ImportExcel -Scope CurrentUser
   ```

   Se for solicitado, digite `Y` e pressione Enter para confirmar.

---

## 📁 Estrutura de arquivos esperada

Certifique-se de que os arquivos estejam organizados assim:

```
ImportItemsToList
│
├── Config
│   ├── IgnoreColumns.json
│   └── Sites.json
│
├── Excel
│   └── DiarioBordo.xlsx
│
├── 01 - ImportItemsToList.ps1
├── Instructions.md
└── SchemaListColumns.json
```

---

## ▶️ Como executar o script

1. **Abra o PowerShell normal (sem ser como administrador)**

2. **Navegue até a pasta do projeto**

   Use o comando `cd` para acessar a pasta onde estão os arquivos. Exemplo:

   ```powershell
   cd "C:\Usuários\SeuNome\Documentos\ImportItemsToList"
   ```

   Substitua o caminho acima pelo local correto da pasta no seu computador.

3. **Execute o script**

   Digite o comando abaixo e pressione Enter:

   ```powershell
   .\01 - ImportItemsToList.ps1
   ```

4. **Selecione o site SharePoint**

   O script vai mostrar uma lista de sites disponíveis. Digite o número correspondente ao site desejado e pressione Enter.

5. **Selecione a lista**

   O script vai mostrar todas as listas disponíveis no site. Digite o número da lista onde os dados devem ser importados.

6. **Login no SharePoint**

   Será aberta uma janela pedindo para fazer login com sua conta Microsoft. Faça o login normalmente.

7. **Aguarde a importação**

   O script irá processar a planilha Excel e importar os dados linha por linha. Ao final, mostrará quantos itens foram importados.

---

## ℹ️ Observações importantes

- A planilha Excel **deve ter os nomes das colunas na primeira linha**
- Não use células mescladas ou formatos avançados
- O script ignora automaticamente colunas que não são compatíveis ou que estão na lista de exclusão (`IgnoreColumns.json`)

---

## 🆘 Suporte

Se ocorrer algum erro durante a execução, anote a mensagem exibida e envie para o responsável técnico pelo script.

