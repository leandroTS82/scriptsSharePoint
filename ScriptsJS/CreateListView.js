function CreateListView(site, listTitle, viewTitle, viewFieldsArray, rowLimit = 30) {
    fetch(site + "/_api/contextinfo", {
        method: "POST",
        headers: {
            "Accept": "application/json;odata=verbose"
        }
    })
        .then(res => res.json())
        .then(data => {
            const requestDigest = data.d.GetContextWebInformation.FormDigestValue;

            // Monta o CAML com os campos
            const camlFields = viewFieldsArray.map(f => `<FieldRef Name='${f}' />`).join("");

            const body = {
                "__metadata": { "type": "SP.View" },
                "Title": viewTitle,
                "RowLimit": rowLimit,
                "ViewQuery": `<View><ViewFields>${camlFields}</ViewFields></View>`,
                "ViewType": "HTML"
            };

            return fetch(site + "/_api/web/lists/getbytitle('" + listTitle + "')/views", {
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": requestDigest
                },
                body: JSON.stringify(body)
            });
        })
        .then(res => {
            if (res.ok) {
                console.log("View criada com sucesso!");
            } else {
                res.json().then(err => console.error("Erro ao criar a view:", err));
            }
        })
        .catch(err => console.error("Erro:", err));
}
/*
CreateListView(
  "https://butterflygrowth.sharepoint.com/sites/leandrogrupoteste",
  "ListaTeste",
  "MinhaViewPersonalizada",
  ["Title", "NovaColunaTexto", "OutraColunaTextoChoice"],
  100
);

*/