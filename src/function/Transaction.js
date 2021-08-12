
export function getTransaction(dateFrom) {
    numberPageTransaction(dateFrom, function (data) {
        getFullTransaction(dateFrom, data["data"]["totalRecords"], function (totalRecords) {
            //console.log(totalRecords["data"]["records"])
            showTrans(handlerData(totalRecords["data"]["records"]))
        })
    })
}

function getFullTransaction(dateFrom, totalRecords, callback) {
    var url = URL_ROOT + "transactions?fromDate=" + dateFrom + "&pageSize=" + totalRecords
    fetch(url, {
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
        .then(callback)
        .catch(function (error) {
            console.log(error)
        })
}

function numberPageTransaction(dateFrom, callback) {
    console.log(ACCESS_TOKEN)
    var url = URL_ROOT + "transactions?fromDate=" + dateFrom
    fetch(url, {
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
        .then(callback)
        .catch(function (error) {
            console.log(error)
        })
}

// [{}{}{}{}{}]
function handlerData(records) {
    var arrTrans = []
    for (let i = 0; i < records.length; i++) {
        arrTrans.push(convertArray(records[i]))
        //console.log(arrTransaction_FromTo)
    }
    return arrTrans
}

function convertArray(obj) {
    return [obj["id"], obj["tid"], obj["amount"], obj["when"]]
}


function showTrans(TOTAL_TRANSACTION) {

    console.log(TOTAL_TRANSACTION.length)
    console.log(typeof TOTAL_TRANSACTION)
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Transaction")
        var tableTrans = sheet.tables.add("A1:D1", true)
        tableTrans.name = "tableTrans"

        tableTrans.getHeaderRowRange().values = [
            ["id", "ID giao dịch", "Số tiền", "Thời gian giao dịch"]
        ];

        tableTrans.rows.add(null, TOTAL_TRANSACTION)

        //var dataRange = sheet.getRange(`B3:E${3 + TOTAL_TRANSACTION.length - 1}`);
        //var dataRange = sheet.getRange("B3:E12");
        //dataRange.values = arrTransaction_FromTo;

        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();
        }

        sheet.activate();

        return context.sync();
    });

}