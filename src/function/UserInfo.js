import {ClearAllData} from "../function/Excelfunction.js"
import {setLoading} from "../function/Loading.js"
export function getUserInfo() {
    ClearAllData("UserInfo")
    fetch(URL_ROOT + "userInfo", {
        method: "GET",
        redirect: "follow", // manual, *follow, error
        mode: "cors", // no-cors, *cors, same-origin
        cache: "default", // *default, no-cache, reload, force-cache, only-if-cached
        credentials: "same-origin",
        headers: {
            authorization: `${ACCESS_TOKEN}`,
            "X-Auth-Token": `${ACCESS_TOKEN}`,
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Credentials": true,
        },
    })
        .then(function (respond) {
            return respond.json()
        })
        .then(function (data) {
            //return data["data"]
            console.log(handleUserInfo(data["data"]))
            showExcel(handleUserInfo(data["data"]))
        })
        .catch(function (error) {
            setTimeout(setLoading(false),2000) 
            console.log(error)
        })
}
//  return  vể 1 obj 
function handleUserInfo(data) {
    return {
        email: data["user"]["email"],
        name: data["business"]['name'],
        bankAccs: data["bankAccs"]
    }
}
function showExcel(objData) {
    Excel.run(function (ctx) {
        var sheet = ctx.workbook.worksheets.getItem("UserInfo");
        sheet.getRange("A1:A1").values = [["Personal Information"]]
        sheet.getRange("A1:A1").format.font.color = "#FFFFFF"
        sheet.getRange("A1:A1").format.fill.color = "#1966A6"
        sheet.getRange("A1:A1").format.horizontalAlignment = Excel.HorizontalAlignment.center
        sheet.getRange('A1:B1').merge(false);

        sheet.getRange("C1:C1").values = [["Bank Accounts"]]
        sheet.getRange("C1:C1").format.font.color = "#FFFFFF"
        sheet.getRange("C1:C1").format.fill.color = "#1966A6"
        sheet.getRange("C1:C1").format.horizontalAlignment = Excel.HorizontalAlignment.center
        sheet.getRange('C1:E1').merge(false);

        // table personal information
        var tablePersonalInformation = sheet.tables.add("A2:B2", true)
        tablePersonalInformation.clearFilters()
        tablePersonalInformation.name = "PersonalInformation"
        tablePersonalInformation.getHeaderRowRange().values = [
            ["   Email   ", "   Business   "]
        ];
        tablePersonalInformation.rows.add(null, [
            [objData["email"], objData["name"]]
        ])

        // table bank accounts
        var tableBankAcc = sheet.tables.add("C2:E2", true)
        tableBankAcc.clearFilters()
        tableBankAcc.name = "BankAccounts"
        tableBankAcc.getHeaderRowRange().values = [
            [ "   Bank Name   ","   BankSubAccId   ", "   BankAccountName   "]
        ];
        if (objData["bankAccs"].length != 0) {
            var arr = []
            console.warn("length not null")
            console.log(objData["bankAccs"])
            objData["bankAccs"].forEach(element => {
                arr.push([
                    element["bank"]["codeName"],
                    element["bankSubAccId"],
                    element["bankAccountName"]
                ])
            });
            tableBankAcc.rows.add(null,arr)
        }
     
        sheet.getRange("A2:E2").format.horizontalAlignment = Excel.HorizontalAlignment.center


        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();
            // sheet.getUsedRange().format.fill.color = "#FFFFFF"
        }

        // sheet.activate();
        return ctx.sync()
            .then(function () { });
    }).catch(errorHandler);
}

