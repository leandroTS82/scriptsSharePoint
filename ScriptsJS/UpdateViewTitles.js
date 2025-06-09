async function getRequestDigest(site) {
    const response = await fetch(site + "/_api/contextinfo", {
        method: "POST",
        headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
}

async function getViews(site, listTitle) {
    const response = await fetch(`${site}/_api/web/lists/getbytitle('${listTitle}')/views`, {
        method: "GET",
        headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    return data.d.results;
}

async function renameView(site, listTitle, viewId, newTitle, requestDigest) {
    const body = {
        "__metadata": { "type": "SP.View" },
        "Title": newTitle
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
        console.error(`Erro ao renomear a view:`, errorData);
    } else {
        console.log(`View renomeada para '${newTitle}' com sucesso!`);
    }
}

async function updateViewTitles(config) {
    const { site, lista: listTitle, Views } = config;
    const requestDigest = await getRequestDigest(site);
    const views = await getViews(site, listTitle);

    for (const viewMap of Views) {
        const view = views.find(v => v.Title === viewMap.TituloViewAtual);
        if (view) {
            await renameView(site, listTitle, view.Id, viewMap.TituloViewAlterado, requestDigest);
        } else {
            console.warn(`View com título '${viewMap.TituloViewAtual}' não encontrada.`);
        }
    }
}

// Exemplo de uso:
const config = {
    "site": "https://butterflygrowth.sharepoint.com/sites/CasadaVedacao",
    "lista": "Cronograma de Aprovações",
    "Views": [
        { "TituloViewAtual": "03_Calendario", "TituloViewAlterado": "01_Calendario" },
        { "TituloViewAtual": "02_Texto", "TituloViewAlterado": "02_Texto" }
        { "TituloViewAtual": "01_Arte", "TituloViewAlterado": "03_Arte" }
    ]
};

updateViewTitles(config);