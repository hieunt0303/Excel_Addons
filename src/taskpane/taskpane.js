import { get } from "../../src/function/getAccessToken.js"
import { getUserInfo } from "../../src/function/UserInfo.js"
import { getTransaction, filterPage } from "../../src/function/Transaction.js"
import swal from 'sweetalert';
import { ClearAllData, addInitialSheet, deleteSheet, add1Sheet, renameCurrentSheet, getAPIKey_fromTable, handleAccessToken } from "../function/Excelfunction.js"
import { addContent, getInformationFromAPIKEY } from "../function/HandleAPI.js"
/* global console, document, Excel, Office */

//#region something not necessary
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

  checkSheetAPIKey_active()

  }
});
export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
//#endregion



// HÀM ĐỂ GÁN API KEY TỪ SHEET "HANDLE API" ĐỂ SAU NÀY TÍNH TOÁN
// NẾU CHƯA CÓ SHEET HOẶC ĐÃ XÓA THÌ API_KEY =="null" --> SẼ TẠO MỚI LẠI CÁC SHEET






// console.log(API_KEY)
// CHẠY HÀM NÀY ĐỂ LOAD RA CÁI COMBOBOX HIỂN THỊ CHỌN NGÀY CHO GIAO DỊCH 


// DÙNG CHO NGƯỜI MỚI : CHECK XEM ĐÃ CÓ TAB HANDLE APIKEY CHƯA( ĐÃ CÓ APIKEY ĐC NHẬP CHƯA ), NẾU CHƯA THÌ TẠO
//checkSheetAPIKey_active()

//<<==================================================== API KEY ===============================================>>


var elementExpired = document.getElementById("header_accesstoken_expired")
var elementExist = document.getElementById("header_accesstoken_exist")


document.getElementById("button_reload_APIKey").onclick = function () {
  getAPIKey_fromTable()
  handleAccessToken(API_KEY)
}

//<<==================================================== USER INFO ===============================================>>

document.getElementById("button_getUserInfo").onclick = function () {
  //getUserInfo()
  //console.log( "2021-02-03" < "2021-02-02")
  addContent(API_KEY, ACCESS_TOKEN, "null")
}

//#region <<============================================== TRANSACTION ==========================================>>

document.getElementById("button_getTransaction").onclick = function () {

  var txtDate = document.getElementById("txtDate")

  if (!txtDate.value)
    swal("Error", "Please enter information about  combobox");
  else {
    //console.log(formatDate(txtDate.value))
    ClearAllData("Transaction")
    getTransaction(formatDate(txtDate.value))
  }

}

//#endregion

//#region <<============================================ HANDLE APIKEY ==========================================>>
var text_apiKey = document.getElementById("text_apiKey")
document.getElementById("button_submit_apiKey").onclick = function (e) {
  if (text_apiKey.value == "")
    swal("Please enter Api Key");
  else {
    loadingPage()

    document.getElementById("api_key").style.display = "none"
    document.getElementById("main").style.display = "none"
    showLoading()
  }
}
//#endregion
//getAPIKey_fromTable()
//#region <<=========================================== ANOTHER FUNCTION ========================================>>
function loadingPage() {
  // api key == null nghĩa là mới dùng lần đầu 

  if (API_KEY == "null") {
    document.getElementById("api_key").style.display = "block"
    document.getElementById("main").style.display = "none"
    var h = new Promise(function (resolve) {
      resolve()
    })
    h
      .then(function () {
        console.log('create Transaction')
        add1Sheet("Transaction")
      })
      .then(function () {
        console.log('create UserInfo')
        add1Sheet("UserInfo")
      })
      .then(function () {
        console.log('create Handle API')
        add1Sheet("Handle API")
      })

  }
  else if (API_KEY == "" && ACCESS_TOKEN == "") {
    document.getElementById("api_key").style.display = "block"
    document.getElementById("main").style.display = "none"
  }
  else {
    document.getElementById("api_key").style.display = "none"
    document.getElementById("main").style.display = "block"
  }


}


// change api key
document.getElementById("handle_changeApiKey").onclick = function () {
  swal("Are you sure you want to DELETE ALL DATA?", {
    buttons: ["No", "Yes"],
  })
    .then(function (result) {
      if (result) {
        API_KEY = ""
        ACCESS_TOKEN = ""
        ClearAllData("Handle API")
        ClearAllData("Transaction")
        ClearAllData("UserInfo")

        loadingPage()

      }
    })
}
// Lấy ra tháng và năm 2021-08-04
function convertMonthDate(lastDate) {
  return {
    month: parseInt(lastDate.split("-")[1]),
    year: parseInt(lastDate.split("-")[0]),
  }
}

function createDataCombo(objDate) {
  var arrData = []
  for (let i = 2020; i <= objDate.year; ++i) {
    for (let j = 1; j <= 12; ++j) {
      if (j > objDate.month && i == objDate.year)
        break;
      arrData.push({
        month: j,
        year: i
      })
    }
  }
  console.log(arrData)
  for (let i = 0; i < arrData.length; ++i) {
    var opt = document.createElement("option")

    opt.value = `${arrData[i].month}/${arrData[i].year}`
    opt.text = `${arrData[i].month}/${arrData[i].year}`
    optionItem.appendChild(opt)
  }
}

function formatDate(date) {
  // 28/8/2020 --> 2020-08-28
  var format = date.split("/")
  return `${format[2]}-${format[1]}-${format[0]}`
}


function showLoading() {
  document.getElementsByClassName("loader")[0].style.display = "block"
  setTimeout(getInformationFromAPIKEY(text_apiKey.value), 3000)

}
//checkSheetAPIKey_active()

function checkSheetAPIKey_active() {
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
        ACCESS_TOKEN = merchantColumnValues[1][0]
        console.log(merchantColumnValues[0][0])

        document.getElementById("api_key").style.display = "none"
        document.getElementById("main").style.display = "block"
        // Sync to update the sheet in Excel
        return context.sync();
      });
  }).catch(function (error) {
    console.log(error)
    document.getElementById("api_key").style.display = "block"
    document.getElementById("main").style.display = "none"
  });
}
//#endregion
