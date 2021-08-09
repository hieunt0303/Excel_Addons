import { showTrans } from "../function/Excelfunction.js"

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
// //////////////////////////////////////////////////////////////////////
// 
// Handle select from date to date

var arrTransaction_FromTo = []

export function getTransaction_fromTo(dateFrom, dateTo) {
    numberPageTransaction(dateFrom, function (data) {
        getFullTransaction(dateFrom, dateTo, data["data"]["totalPages"])
    })

}

var finishHandle = false
function getFullTransaction(dateFrom, dateTo, numberPagesTrans) {
    //console.log('total pages : ' + numberPagesTrans)
    for (let i = 1; i <= numberPagesTrans; ++i) {
        if (finishHandle == false)
            fetchEachPageData(dateFrom, dateTo, i)
        else {
            console.log(arrTransaction_FromTo)
            showTrans(arrTransaction_FromTo.sort(function (a, b) {
                if (a["when"] > b["when"])
                    return -1
                if (a["when"] < b["when"])
                    return 1
                return 0
            }))
            break
        }
    }

}

function numberPageTransaction(dateFrom, callback) {
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

function fetchEachPageData(dateFrom, dateTo, page) {
    var url = URL_ROOT + "transactions?fromDate=" + dateFrom + "&page=" + page
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
            //exportExcel(data["data"]["records"])
            handlerData_fromTo(dateTo, data["data"]["records"])
        })
        .catch(function (error) {
            console.log(error)
        })
}
// [{}{}{}{}{}]
function handlerData_fromTo(dateTo, records) {
    for (let i = 0; i < records.length; i++) {
        if (records[i]["when"] <= dateTo) {
            arrTransaction_FromTo.push(convertArray(records[i]))
            //console.log(arrTransaction_FromTo)
        }
        else {
            finishHandle = true
            break
        }
    }
}

function convertArray(obj) {
    return [obj["id"], obj["tid"], obj["amount"], obj["when"]]
}
