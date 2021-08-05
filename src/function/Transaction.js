export function lastTransPage(callback) {
    var url = URL_ROOT + "transactions?fromDate=2020-01-01"
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
        .then(function (result) {
            LAST_PAGE_RECORDS = result["data"]["totalPages"]
            return result["data"]["totalPages"]
        })
        .then(callback)
}
export function lastTransDate(lastPage, callback) {
    var url = URL_ROOT + "transactions?fromDate=2020-01-01&page=" + lastPage
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
        .then(function (result) {
            var arrLength = result["data"]["records"].length
            return (result["data"]["records"][arrLength - 1]["when"])
        })
        .then(callback)
}

export function getTransaction() {
    numberPageTransaction(function (data) {
        getFullTransaction(data["data"]["totalPages"])
    })

}

function getFullTransaction(numberPagesTrans) {
    for (let i = 1; i <= numberPagesTrans; ++i) {
        fetchEachPageData(i)
    }

}

function numberPageTransaction(callback) {
    var url = URL_ROOT + "transactions?fromDate=2020-01-01"
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

function fetchEachPageData(page) {
    var url = URL_ROOT + "transactions?fromDate=2020-01-01&page=" + page
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
        .then(function (data) {
            exportExcel(data["data"]["records"])
        })
        .catch(function (error) {
            console.log(error)
        })
}

function exportExcel(arrRecords) {
    // sai hieenr thij
    var arrArr = []
    for (let i = 0; i < arrRecords.length; ++i) {
        arrArr.push(convertArray(arrRecords[i]))
        if (TOTAL_TRANSACTION.length < 1130)
            TOTAL_TRANSACTION.push(convertArray(arrRecords[i]))
    }
    console.log(TOTAL_TRANSACTION.length)
    if (TOTAL_TRANSACTION.length == 1130)
        ExcelAPITransaction(arrArr)
}
function convertArray(obj) {
    return [obj["id"], obj["tid"], obj["amount"], obj["when"]]
}

function ExcelAPITransaction() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        // Create the headers and format them to stand out.
        var headers = [
            ["id", "ID giao dịch", "Số tiền", "Thời gian giao dịch"]
        ];
        var headerRange = sheet.getRange("B2:E2");
        headerRange.values = headers;
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";

        // Create the product data rows.
        var dataRange = sheet.getRange(`B3:E${3 + TOTAL_TRANSACTION.length - 1}`);
        //var dataRange = sheet.getRange("B3:E12");
        dataRange.values = TOTAL_TRANSACTION;
        for (let i = 3; i <= 3 + TOTAL_TRANSACTION.length - 1; ++i) {
            if (i % 2 == 0)
                sheet.getRange(`B${i}:E${i}`).format.fill.color = "#CCFFFF"
            else
                sheet.getRange(`B${i}:E${i}`).format.fill.color = "#6699FF"

        }
        return context.sync();
    });

}

// <<=================================================================================================>>
//
// CÁC HÀM ĐỂ LỌC DATA THEO THÁNG 
// objDate = "2/2020"
//
// BINARY SEARCH
function Filter(objDate) {
    var left = 1
    var right = LAST_PAGE_RECORDS
    var middle
    while (left <= right) {
        middle = (left + right) / 2
        if (true) {
            return
        }

    }
}
export function filterPage(page, objDate) {
    var url = URL_ROOT + "transactions?fromDate=2020-01-01&page=" + page
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
        .then(function (data) {
            //handleDataPage(data["data"]["page"], data["data"]["records"], objDate)
            console.log('hieu123')
            return true
        })
      
}
//      [{}{}{}{}]
// objDate -- 2/2020
