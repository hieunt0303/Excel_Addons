//addContent(API_KEY,ACCESS_TOKEN,"null")
import {updateInformation_handleAPI} from "../function/Excelfunction.js"
export function addContent(APIKey, AccessToken, RefreshToken) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Handle API")
        // Create the headers and format them to stand out.


        // Create the product data rows.
        var dataRange = sheet.getRange(`A1:B4`);



        dataRange.values = [
            ["API Key", APIKey],
            ["Access Token", AccessToken],
            ["Refresh token", RefreshToken],
            ["Expired Access token at:", "data"]
        ];
        sheet.getRange(`A1:B4`).format.autofitColumns()
        // sheet.getRange(`B1`).format.autofitColumns()
        // sheet.getRange(`A4`).format.wrapText = true
        // sheet.getRange(`B2`).format.wrapText = true

        return context.sync();
    });
}

export function getInformationFromAPIKEY(apiKey) {
    var url = URL_ROOT + "token"
    fetch(url, {
        method: "POST",
        redirect: "follow", // manual, *follow, error
        mode: "cors", // no-cors, *cors, same-origin
        cache: "default", // *default, no-cache, reload, force-cache, only-if-cached
        credentials: "same-origin",
        headers: {
            'Authorization': apiKey,
            Accept: 'application/json',
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            "code": apiKey
        })

    })
        .then(function (respond) {
            return respond.json()
        })
        .then(function (result) {
            console.log(result)
            updateInformation_handleAPI(result,apiKey)
            document.getElementById("main").style.display = "block"
            document.getElementById("loader").style.display = "none"
        })
}