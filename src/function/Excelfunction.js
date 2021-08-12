import { API_KEY, ACCESS_TOKEN } from "../valueConst.js"
export function ClearAllData(SHEET) {
    if (!SHEET) {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            sheet.getRange().clear();
            return ctx.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    else {
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getItem(SHEET);
            sheet.getRange().clear();
            return ctx.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
}
// Xóa các sheet tồn tại và thêm vào 2 sheet cần thiết trans và userinfo
export function addInitialSheet() {
    renameCurrentSheet("Transaction")

}
export function deleteSheet(nameSheet) {
    Excel.run(function (context) {
        var sheets = context.workbook.worksheets.getItem(nameSheet)
        sheets.load("items/name");

        return context.sync()
            .then(function () {
                if (sheets.items.length === 1) {
                    console.log("Unable to delete the only worksheet in the workbook");
                } else {
                    var lastSheet = sheets.items[sheets.items.length - 1];

                    console.log(`Deleting worksheet named "${lastSheet.name}"`);
                    lastSheet.delete();

                    return context.sync();
                };
            });
    }).catch(errorHandlerFunction);
}
export function renameCurrentSheet(nameSheet) {
    Excel.run(function (context) {
        var currentSheet = context.workbook.worksheets.getActiveWorksheet();
        currentSheet.name = nameSheet;

        return context.sync()
            .then(function () {
                console.log('rename success')
            })

    })

}
export function add1Sheet(nameSheet) {
    Excel.run(function (context) {
        var sheets = context.workbook.worksheets;

        var sheet = sheets.add(nameSheet);
        sheet.load("name, position");

        return context.sync()
            .then(function () {
                console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
            });
    })
}

// Handler date --> show transaction
// records = [[],[],[],[]]
function addData_Initial_HandleAPI() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Handle API")

        // Create the product data rows.
        var dataRange = sheet.getRange(`A1:A1`);
        //var dataRange = sheet.getRange("B3:E12");
        dataRange.values = [["Please enter API Key to Taskpane!!"]];
        sheet.getRange(`A1:A1`).format.font.color = "#FF0033"

        return context.sync();
    });
}
//<<===================================== Access token .js ====================================================>>
export function handleAccessToken(api_key) {
    console.log(api_key)
    var url = URL_ROOT + "token"
    fetch(url, {
        method: "POST",
        redirect: "follow", // manual, *follow, error
        mode: "cors", // no-cors, *cors, same-origin
        cache: "default", // *default, no-cache, reload, force-cache, only-if-cached
        credentials: "same-origin",
        headers: {
            'Authorization': api_key,
            Accept: 'application/json',
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            "code": api_key
        })

    })
        .then(function (respond) {
            return respond.json()
        })
        .then(function (result) {
            console.log(result)
            getAPIKey_fromTable()
            ClearAllData("Handle API")

            updateInformation_handleAPI(result, api_key)


        })
}
//<<===================================== Handle API .js ====================================================>>
export function getAPIKey_fromTable() {
    Excel.run(function (context) {
        try {
            var sheet = context.workbook.worksheets.getItem("Handle API");

        } catch (error) {
            console.log(error)
        }
        var expensesTable = sheet.tables.getItem("HandleAPI");

        // Get data from a single column
        var columnRange = expensesTable.columns.getItem("Values").getDataBodyRange().load("values");


        return context.sync()
            .then(function () {

                var merchantColumnValues = columnRange.values;
                API_KEY = merchantColumnValues[0][0]
                console.log(merchantColumnValues[0][0])

                // Sync to update the sheet in Excel
                return context.sync();
            });
    }).catch(function (error) {
        console.log(error)
    });
}
// result = {}
export function updateInformation_handleAPI(result, api_key) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Handle API")
        var nametable = sheet.tables.add("A1:B1", true)
        nametable.name = "HandleAPI"
        nametable.getHeaderRowRange().values = [[nametable.name, "Values"]]

        var current = new Date()
        current.getTime()
        current.setSeconds(current.getSeconds() + 21600)
        console.log(current)

        API_KEY = api_key
        ACCESS_TOKEN = result["access_token"]

        nametable.rows.add(null, [
            ["API Key", api_key],
            ["Access Token", result["access_token"]],
            ["Refresh token", result["refresh_token"]],
            ["Expired Access token at:", formatCurrentDate(current)]
        ])

        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();
        }

        sheet.activate();



        return context.sync();
    });
}

function formatCurrentDate(date) {
    console.log(typeof String(date))

    var arr = String(date).split(' ')

    var formatVN_day = {
        'Sun': "Chủ nhật",
        'Mon': "Thứ 2",
        'Tue': "Thứ 3",
        'Wed': "Thứ 4",
        'Thu': "Thứ 5",
        'Fri': "Thứ 6",
        'Sat': "Thứ 7"
    }
    var formatVN_months = {
        'Jan': "1",
        'Feb': "2",
        'Mar': "3",
        'Apr': "4",
        'May': "5",
        'Jun': "6",
        'Jul': "7",
        'Aug': "8",
        'Sep': "9",
        'Oct': "10",
        'Nov': "11",
        'Dec': "12"
    }
    console.log(arr)
    var output =
        `${formatVN_day[arr[0]]} ${arr[2]}/${formatVN_months[arr[1]]}/${arr[3]} vào lúc ${arr[4]} `

    console.log(output)
    return output
}

