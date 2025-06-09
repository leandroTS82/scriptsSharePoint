async function getRequestDigest(site) {
    const response = await fetch(site + "/_api/contextinfo", {
        method: "POST",
        headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
}

async function getViewByTitle(site, listTitle, viewTitle) {
    const response = await fetch(`${site}/_api/web/lists/getbytitle('${listTitle}')/views`, {
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    const view = data.d.results.find(v => v.Title === viewTitle);
    return view ? view : null;
}

async function getFieldByDisplayName(site, listTitle, displayName) {
    const response = await fetch(`${site}/_api/web/lists/getbytitle('${listTitle}')/fields`, {
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    const field = data.d.results.find(f => f.Title === displayName);
    return field ? field : null;
}

function buildCamlQuery(internalName, fieldType, filtro) {
    let operador;
    switch (filtro.Operador) {
        case "Igual":
            operador = "Eq";
            break;
        case "Diferente":
            operador = "Neq";
            break;
        default:
            throw new Error("Operador não suportado: " + filtro.Operador);
    }

    // Determina o tipo correto para CAML
    let valueType = "Text";
    if (fieldType === "Choice") valueType = "Text";
    else if (fieldType === "Number") valueType = "Number";
    else if (fieldType === "DateTime") valueType = "DateTime";

    return `
        <Where>
            <${operador}>
                <FieldRef Name='${internalName}'/>
                <Value Type='${valueType}'>${filtro.Valor}</Value>
            </${operador}>
        </Where>
    `;
}

async function updateViewQuery(site, listTitle, viewId, newCamlQuery, requestDigest) {
    const body = {
        "__metadata": { "type": "SP.View" },
        "ViewQuery": newCamlQuery
    };

    const response = await fetch(`${site}/_api/web/lists/getbytitle('${listTitle}')/views(guid'${viewId}')`, {
        method: "POST",
        headers: {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": requestDigest,
            "X-HTTP-Method": "MERGE",
            "If-Match": "*"
        },
        body: JSON.stringify(body)
    });

    if (!response.ok) {
        const errorData = await response.json();
        console.error("Erro ao atualizar o filtro da view:", errorData);
    } else {
        console.log(`Filtro da view atualizado com sucesso! Site: ${site} | Lista: ${listTitle} | View: '${viewId}'`);
    }
}

async function updateViewFilter(config) {
    const { site, lista: listTitle, View: viewTitle, filtro } = config;
    const requestDigest = await getRequestDigest(site);
    const view = await getViewByTitle(site, listTitle, viewTitle);

    if (!view) {
        console.warn(`View com título '${viewTitle}' não encontrada.`);
        return;
    }

    const field = await getFieldByDisplayName(site, listTitle, filtro.Coluna);

    if (!field) {
        console.warn(`Coluna '${filtro.Coluna}' não encontrada na lista '${listTitle}'.`);
        return;
    }

    if (field.FieldTypeKind === 7) { // Lookup
        console.warn(`A coluna '${filtro.Coluna}' é do tipo Lookup, ajuste necessário para suportar este tipo.`);
        return;
    }

    if (field.TypeAsString === "Choice") {
        const choices = field.Choices.results;
        if (!choices.includes(filtro.Valor)) {
            console.warn(`Valor '${filtro.Valor}' não é uma das opções válidas para a coluna '${filtro.Coluna}'. Opções válidas: ${choices.join(", ")}`);
            return;
        }
    }

    const camlQuery = buildCamlQuery(field.InternalName, field.TypeAsString, filtro);
    await updateViewQuery(site, listTitle, view.Id, camlQuery, requestDigest);
}

// Exemplo de uso:
var site = "https://butterflygrowth.sharepoint.com/sites/hopere";
var lista = "Cronograma de Aprovações"
var config = {
    "site": site,
    "lista": lista,
    "View": "01_Calendario",
    "filtro": {
        "Coluna": "Etapa", // display name correto
        "Operador": "Igual",
        "Valor": "Calendário"
    }
};
updateViewFilter(config);

config = {
    "site": site,
     "lista": lista,
     "View": "02_Texto",
     "filtro": {
         "Coluna": "Etapa", // display name correto
         "Operador": "Igual",
         "Valor": "Posts"
     }
 };
 updateViewFilter(config);

config = {
   "site": site,
    "lista": lista,
    "View": "03_Artes",
    "filtro": {
        "Coluna": "Etapa", // display name correto
        "Operador": "Igual",
        "Valor": "Artes"
    }
};
updateViewFilter(config);
