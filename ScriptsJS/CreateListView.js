async function UpdateViewFields(site, listTitle, viewId, viewFieldsArray) {
    const contextResponse = await fetch(site + "/_api/contextinfo", {
        method: "POST",
        headers: {
            "Accept": "application/json;odata=verbose"
        }
    });
    const contextData = await contextResponse.json();
    const requestDigest = contextData.d.GetContextWebInformation.FormDigestValue;

    // Remove todos os campos da View
    await fetch(`${site}/_api/web/lists/getbytitle('${listTitle}')/views(guid'${viewId}')/ViewFields/removeallviewfields`, {
        method: "POST",
        headers: {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": requestDigest
        }
    });

    // Adiciona os campos desejados
    for (const field of viewFieldsArray) {
        const addFieldResponse = await fetch(
            `${site}/_api/web/lists/getbytitle('${listTitle}')/views(guid'${viewId}')/ViewFields/addviewfield('${field}')`,
            {
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": requestDigest
                }
            }
        );

        if (!addFieldResponse.ok) {
            const err = await addFieldResponse.json();
            console.error(`Erro ao adicionar o campo '${field}':`, err);
        }
    }

    console.log("Campos atualizados com sucesso!");
}
async function CreateListView(site, listTitle, viewTitle, viewFieldsArray, rowLimit = 30) {
    try {
        const contextResponse = await fetch(site + "/_api/contextinfo", {
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose"
            }
        });
        const contextData = await contextResponse.json();
        const requestDigest = contextData.d.GetContextWebInformation.FormDigestValue;

        const createViewBody = {
            "__metadata": { "type": "SP.View" },
            "Title": viewTitle,
            "RowLimit": rowLimit,
            "ViewType": "HTML"
        };

        const createViewResponse = await fetch(`${site}/_api/web/lists/getbytitle('${listTitle}')/views`, {
            method: "POST",
            headers: {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": requestDigest
            },
            body: JSON.stringify(createViewBody)
        });

        const viewData = await createViewResponse.json();
        const viewId = viewData.d.Id;

        console.log("View criada com sucesso! Atualizando campos...");
        await UpdateViewFields(site, listTitle, viewId, viewFieldsArray);

    } catch (err) {
        console.error("Erro geral:", err);
    }
}
/*
CreateListView(
  "https://butterflygrowth.sharepoint.com/sites/leandrogrupoteste",
  "ListaTeste",
  "MinhaViewPersonalizada",
  ["Title", "NovaColunaTexto", "NovaColunaTexto2", "OutraColunaTexto3"],
  100
);

*/