async function GetStringJsonListColumns(site, list) {
    try {
      // Obter ID da list
      const listResponse = await fetch(`${site}/_api/web/lists/getbytitle('${list}')?$select=Id`, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose"
        }
      });
  
      if (!listResponse.ok) {
        throw new Error(`Erro ao obter ID da list: ${listResponse.statusText}`);
      }
  
      const listData = await listResponse.json();
      const listId = listData.d.Id;
  
      // Obter colunas visíveis com tipo e choices
      const colunasResponse = await fetch(`${site}/_api/web/lists/getbytitle('${list}')/fields?$select=Title,InternalName,FieldTypeKind,Choices&$filter=Hidden eq false`, {
        method: "GET",
        headers: {
          "Accept": "application/json;odata=verbose"
        }
      });
  
      if (!colunasResponse.ok) {
        throw new Error(`Erro ao obter colunas: ${colunasResponse.statusText}`);
      }
  
      const colunasData = await colunasResponse.json();
  
      const resultado = {
        listId: listId,
        Colunas: colunasData.d.results.map(field => {
          const coluna = {
            Title: field.Title,
            InternalName: field.InternalName,
            FieldTypeKind: field.FieldTypeKind
          };
  
          if (field.FieldTypeKind === 6 && field.Choices && field.Choices.results) {
            coluna.Choices = field.Choices.results;
          }
  
          return coluna;
        })
      };
  
      const resultadoComoString = JSON.stringify(resultado, null, 2);
      console.log("Resultado JSON como string:\n", resultadoComoString);
  
      return resultadoComoString;
  
    } catch (error) {
      console.error("Erro:", error);
    }
  }  
  
  // Exemplo de chamada:
  var site = "https://butterflygrowth.sharepoint.com/sites/CasadaVedacao";
 var list = "Cronograma de Aprovações"
  GetStringJsonListColumns(site, list);
  console.log("considerando o padrão, extraia os campos de Artes, texto e CAlendário na ordem e crie a chamadas das functions para criação das views para o site "+site+" e lista "+list)
  