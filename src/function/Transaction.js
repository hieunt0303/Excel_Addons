import swal from "sweetalert"
import { add1Sheet, ClearAllData } from "../function/Excelfunction.js"
import {setLoading} from "../function/Loading.js"
export function getTransaction(dateFrom, typeChart) {
    numberPageTransaction(dateFrom, function (data) {
        getFullTransaction(dateFrom, data["data"]["totalRecords"], function (totalRecords) {
            //console.log(totalRecords["data"]["records"])
            var header =
                ["id", "ID giao dịch", "Số tiền", "Thời gian giao dịch"]
            showTrans(
                "Transaction",
                handlerData(totalRecords["data"]["records"]),
                header,
                "A1:D1",
                "tableTrans"
            )
            if (typeChart != "default") {
                add1Sheet("Chart")
                try {
                    // có trường hợp đã có data rồi và chưa tạo sheet
                    ClearAllData("Chart")
                } catch (error) {
                    console.log(error)
                }
                createTransForDate(typeChart)
            }
            setTimeout(setLoading(false),2000)
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

// [{}{}{}{}{}] --> [[],[],[],[]]
function handlerData(records) {
    var arrTrans = []
    for (let i = 0; i < records.length; i++) {
        arrTrans.push(convertArray(records[i]))
        //console.log(arrTransaction_FromTo)
    }
    console.warn(arrTrans)
    return arrTrans
}

function convertArray(obj) {
    //console.warn(obj)
    let arrWhen = obj["when"].split("-")
    let formatWhen = new Date(arrWhen[0], arrWhen[1] - 1, arrWhen[2])
    // console.warn(formatWhen.toISOString().split("T")[0])

    // var date = new Date(Math.round((JSDateToExcelDate(formatWhen) - (25567 + 2)) * 86400 * 1000));
    // var converted_date = date.toISOString().split('T')[0];
    // let converted_date = `${arrWhen[1]}/${arrWhen[2]}/${arrWhen[0]}`

    // console.warn(converted_date)
    return [obj["id"], obj["tid"], obj["amount"], obj["when"]]
}


function showTrans(nameSheet, arrData, arrHeader, positionHeader, nameTable) {
    console.log(arrData.length)
    console.log(typeof arrData)
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(nameSheet)
        var tableTrans = sheet.tables.add(positionHeader, true)
        tableTrans.name = nameTable
        tableTrans.getHeaderRowRange().values = [
            arrHeader
        ];

        tableTrans.rows.add(null, arrData)

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
function showLatestTrans(arrLatestTrans) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Transaction")
        var tableTrans = sheet.tables.getItem("tableTrans")

        tableTrans.rows.add(null, arrLatestTrans)

        if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
            sheet.getUsedRange().format.autofitColumns();
            sheet.getUsedRange().format.autofitRows();
        }
        return context.sync();
    });
}
export function checkLatestTrans() {
    getLastTrans(function (obj) {
        numberPageTransaction(obj["latestTime"], function (data) {
            getFullTransaction(obj["latestTime"], data["data"]["totalRecords"], function (totalRecords) {
                console.log(totalRecords["data"]["records"])

                totalRecords["data"]["records"].forEach(function (element, index) {
                    if (element["id"] == obj["latestId"]) {
                        var arrLatestTrans = totalRecords["data"]["records"].slice(index + 1)
                        console.log(arrLatestTrans)
                        setLoading(false) // stop SHOW LOADING
                        if (arrLatestTrans.length != 0) {
                            swal("Hey, you just got new transaction do you want to sync it ?", {
                                buttons: {
                                    catch: "Sync Trans",
                                    defeat: "Sync Trans, Chart",
                                    cancel: "No",
                                },
                            })
                                .then(function (result) {
                                    if (result) {
                                        switch (result) {
                                            case "defeat": {
                                                try {
                                                    deleteChart("Chart") // xóa lun cái chart của giao dịch theo ngày
                                                } catch (error) {
                                                    deleteTable("Chart", "tableTransDay") // xóa cái bảng giao dịch theo ngày
                                                    console.error("delete table success")
                                                }
                                                finally {
                                                    showLatestTrans(handlerData(totalRecords["data"]["records"].slice(index + 1)))
                                                    typeChart(function(){
                                                        console.log(this.returnValue)
                                                        createTransForDate(this.returnValue) // tạo lại cái giao dịch theo ngày
                                                      })
                                                }
                                                break;
                                            }
                                            case "catch": {
                                                showLatestTrans(handlerData(totalRecords["data"]["records"].slice(index + 1)))
                                                swal("Success!", "Update Transaction successfull!", "success");
                                                break;
                                            }
                                            default:
                                                swal("The latest Trans is not sync.");
                                        }
                                    }
                                })
                        }else{
                            swal("You have no new transaction.")
                        }
                    }
                })
            })
        })
    })
}
// lấy id và ngày của giao dịch cuối cùng có trong table 
function getLastTrans(callback) {
    Excel.run(function (context) {
        try {
            var sheet = context.workbook.worksheets.getItem("Transaction");

        } catch (error) {
            console.log(error)
        }
        var expensesTable = sheet.tables.getItem("tableTrans");

        // Get data from a single column
        var timeTrans = expensesTable.columns.getItem("Thời gian giao dịch").getDataBodyRange().load("values");
        var idTrans = expensesTable.columns.getItem("id").getDataBodyRange().load("values");
        return context.sync()
            .then(function () {
                var arrId = idTrans.values.map(function (index) {
                    return index[0]
                })
                var arrTime = timeTrans.values.map(function (index) {
                    var date = new Date(Math.round((index[0] - (25567 + 2)) * 86400 * 1000));
                    var converted_date = date.toISOString().split('T')[0];
                    return converted_date;
                })
                var objOutput = {
                    "latestId": arrId[arrId.length - 1],
                    //"latestId": "360331",
                    "latestTime": arrTime[arrTime.length - 1]
                }
                console.log(objOutput)
                // Sync to update the sheet in Excel
                return objOutput
            })
            .then(callback)
    }).catch(function (error) {
        console.log(error)
    });
}

function showLoadingTrans(bool) {
    if (bool) {
        document.getElementsByClassName("loader")[0].style.display = "block"
        document.getElementById("main_group").style.display = "none"
    }
    else {
        document.getElementsByClassName("loader")[0].style.display = "none"
        document.getElementById("main_group").style.display = "block"
    }
}

// hàm dùng để tạo 1 bảng có doanh thu theo tháng từ cái bảng trong trans
// chartData = {allMoney : [[],[],[],...], allTime : [[],[],[],...}
export function createTransForDate(typeChart) {
    console.log(typeChart)
    getAllTimeAndID(function (chartData) {
        var obj = formatTypeChart(chartData, typeChart)

        console.warn(obj)
        var tempTime = obj["allTime"][0]
        var startIndex = 0
        var arrOutput = []
        obj["allTime"].forEach(function (element, index) {
            if (element != tempTime) {
                arrOutput.push([
                    obj["allTime"][index - 1],
                    sumMoneyPerDay(obj["allMoney"], startIndex, index)
                ])
                tempTime = element
                startIndex = index
            }
        })
        arrOutput.push([
            obj["allTime"][obj["allTime"].length - 1],
            sumMoneyPerDay(obj["allMoney"], startIndex, obj["allTime"].length)
        ])
        console.log(arrOutput)
        var header = ["Ngày giao dịch", "Chi", "Thu"]
        showTrans("Chart", formatArrTransPerDay(arrOutput), header, "A1:C1", "tableTransDay")
        showChart("Chart", `A1:C${arrOutput.length}`)
    })
}

function formatArrTransPerDay(arr) {
    var output = []
    arr.forEach(function (element) {
        if (element[1] <= 0) {
            output.push([
                element[0],
                Math.abs(element[1]),
                ""
            ])
        }
        else {
            output.push([
                element[0],
                "",
                element[1],
            ])
        }
    })
    return output
}
function sumMoneyPerDay(arrMoney, startIndex, lastIndex) {
    var sum = 0
    for (var i = startIndex; i < lastIndex; i++) {
        sum += arrMoney[i]
    }
    return sum
}
function createTransForMonth() {

}
function getAllTimeAndID(callback) {
    Excel.run(function (context) {
        try {
            var sheet = context.workbook.worksheets.getItem("Transaction");

        } catch (error) {
            console.log(error)
        }
        var expensesTable = sheet.tables.getItem("tableTrans");

        // Get data from a single column
        var timeTrans = expensesTable.columns.getItem("Thời gian giao dịch").getDataBodyRange().load("values");
        var moneyTrans = expensesTable.columns.getItem("Số tiền").getDataBodyRange().load("values");
        return context.sync()
            .then(function () {
                console.warn(timeTrans.values)

                var arrMoney = moneyTrans.values.map(function (index) {
                    return index[0]
                })
                var arrTime = timeTrans.values.map(function (index) {
                    var date = new Date(Math.round((index[0] - (25567 + 2)) * 86400 * 1000));
                    var converted_date = date.toISOString().split('T')[0];
                    return converted_date;
                })
                var objOutput = {
                    "allMoney": arrMoney,
                    //"latestId": "360331",
                    "allTime": arrTime
                }
                console.log(objOutput)
                // Sync to update the sheet in Excel
                return objOutput
            })
            .then(callback)
    }).catch(function (error) {
        console.log(error)
    });

}
//<<================================================== Chart=======================================>>
export function showChart(nameSheet, position) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(nameSheet);
        var dataRange = sheet.getRange(position);
        var chart = sheet.charts.add("ColumnClustered", dataRange, "auto");

        chart.title.text = "Sales Data";
        chart.legend.position = "right"
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 7;
        chart.dataLabels.format.font.color = "black";

        return context.sync();
    }).catch(errorHandlerFunction);
}
//<<=========================================== Delete =========================================
export function deleteChart(nameSheet) {
    Excel.run(function (context) {
        console.warn("delete chart success")
        var sheet = context.workbook.worksheets.getItem(nameSheet);
        sheet.charts.getItemAt(0).delete();
        return context.sync();
    }).catch(errorHandlerFunction);
}

function deleteTable(nameSheet,nameTable) {
    Excel.run(function (context) {
        console.warn("delete table success")
        var sheet = context.workbook.worksheets.getItem(nameSheet);
        var expensesTable = sheet.tables.getItem(nameTable);

        // Resize the table.
        expensesTable.delete()

        return context.sync();
    }).catch(errorHandlerFunction);
}


//#region ============================================= Dialog chart  ===========================
export function typeChart(callback) {

    const favDialog = document.getElementById('favDialog');
    const selectEl = document.querySelector('select');
    const confirmBtn = document.getElementById('confirmBtn');
    if (typeof favDialog.showModal === "function") {
        favDialog.showModal();
    } else {
        alert("The <dialog> API is not supported by this browser");
    }

    selectEl.addEventListener('change', function onSelect(e) {
        confirmBtn.value = selectEl.value;
    });

    // favDialog.addEventListener('close', function onClose() {
    //   return favDialog.returnValue 
    // });
    favDialog.addEventListener('close', callback, { once: true });
}

function formatTypeChart(obj, typeChart) {
    var arrFormatTime = []
    switch (typeChart) {
        case "Trans By Day": {
            arrFormatTime.push(obj["allTime"])
            break
        }
        case "Trans By Month": {
            arrFormatTime.push(
                obj["allTime"].map(function (element) {
                    return element.split("-")[0] + "-" + element.split("-")[1]
                })
            )
            break
        }
        case "Trans By Year": {
            arrFormatTime.push(
                obj["allTime"].map(function (element) {
                    return element.split("-")[0]
                })
            )
            break
        }
    }

    return {
        allMoney: obj["allMoney"],
        allTime: arrFormatTime[0]
    }
}

function JSDateToExcelDate(inDate) {
    var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
    return returnDateTime.toString().substr(0, 5);
}