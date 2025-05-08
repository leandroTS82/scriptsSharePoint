function CreateListColumn(site, listDisplayName, columnTitle, InternalName, columnTipe) {
  // Obter o token __REQUESTDIGEST
  fetch(site + "/_api/contextinfo", {
    method: "POST",
    headers: {
      "Accept": "application/json;odata=verbose"
    }
  })
    .then(res => res.json())
    .then(data => {
      var requestDigest = data.d.GetContextWebInformation.FormDigestValue;

      // Montar o corpo da requisição para criação da coluna
      let body;

      if (columnTipe === 6) {
        // Se for coluna do tipo Escolha (Choice)
        let opcoesTexto = prompt("Informe as opções separadas por ponto e vírgula (;):", "Opção 1;Opção 2;Opção 3");
        let opcoesArray = opcoesTexto.split(";").map(o => o.trim());

        body = {
          "__metadata": { "type": "SP.FieldChoice" },
          "Title": InternalName,
          "FieldTypeKind": 6,
          "StaticName": InternalName,
          "Choices": { "results": opcoesArray },
          "EditFormat": 0
        };
      } else {
        // Outros tipos de coluna
        body = {
          "__metadata": { "type": "SP.Field" },
          "Title": InternalName,
          "FieldTypeKind": columnTipe,
          "StaticName": InternalName
        };
      }

      // Criar a coluna
      return fetch(site + "/_api/web/lists/getbytitle('" + listDisplayName + "')/fields", {
        method: "POST",
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": requestDigest
        },
        body: JSON.stringify(body)
      })
        .then(res => res.json())
        .then(data => {
          console.log("Coluna criada com sucesso:", data);

          // Atualizar o Title da coluna
          return fetch(site + "/_api/web/lists/getbytitle('" + listDisplayName + "')/fields/getbyinternalnameortitle('" + InternalName + "')", {
            method: "POST",
            headers: {
              "Accept": "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": requestDigest,
              "X-HTTP-Method": "MERGE",
              "IF-MATCH": "*"
            },
            body: JSON.stringify({
              "__metadata": { "type": "SP.Field" },
              "Title": columnTitle
            })
          });
        });
    })
    .then(res => {
      if (res && res.ok) {
        console.log("Título atualizado com sucesso!");
      } else {
        console.error("Erro ao atualizar título:", res.statusText);
      }
    })
    .catch(err => console.error("Erro:", err));
}
/*
Exemplos de acionamento do script:

Data/Hora (DateTime)
CreateListColumn(
  "https://butterflygrowth.sharepoint.com/sites/leandrogrupoteste", // URL do site
  "list Display Name", // Nome da Lista
  "Nova Coluna de Data", // Título da coluna
  "OutraColunaData", // Nome interno da coluna
  4 // FieldTypeKind para DateTime
);

Pessoa (People)
CreateListColumn(
  "https://butterflygrowth.sharepoint.com/sites/leandrogrupoteste", // URL do site
  "list Display Name", // Nome da listDisplayName
  "Nova Coluna de Pessoa", // Título da coluna
  "OutraColunaPessoa", // Nome interno da coluna
  20 // FieldTypeKind para Pessoa (Person or Group)
);

Escolha (Choice)
CreateListColumn(
  "https://butterflygrowth.sharepoint.com/sites/leandrogrupoteste", // URL do site
  "list Display Name", // Nome da listDisplayName
  "Nova Coluna de Escolha", // Título da coluna
  "OutraColunaEscolha", // Nome interno da coluna
  6 // FieldTypeKind para Escolha (Choice)
);




*/