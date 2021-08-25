import { get } from "../../src/function/getAccessToken.js"
import { getUserInfo, showExcel } from "../../src/function/UserInfo.js"
import { getTransaction, checkLatestTrans, typeChart } from "../../src/function/Transaction.js"
import swal from 'sweetalert';
import { ClearAllData, addInitialSheet, deleteSheet, addChartSheet, add1Sheet, renameCurrentSheet, getAPIKey_fromTable, handleAccessToken } from "../function/Excelfunction.js"
import { addContent, getInformationFromAPIKEY } from "../function/HandleAPI.js"
import { setLoading } from "../function/Loading.js"
/* global console, document, Excel, Office */

//#region something not necessary
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    checkSheetAPIKey_active()

    // setInterval(function(){
    //   checkLatestTrans()
    // },5000)

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

// DÙNG CHO NGƯỜI MỚI : CHECK XEM ĐÃ CÓ TAB HANDLE APIKEY CHƯA( ĐÃ CÓ APIKEY ĐC NHẬP CHƯA ), NẾU CHƯA THÌ TẠO
//checkSheetAPIKey_active()

//<<==================================================== API KEY ===============================================>>


var elementExpired = document.getElementById("header_accesstoken_expired")
var elementExist = document.getElementById("header_accesstoken_exist")


document.getElementById("button_reload_APIKey").onclick = function () {
  getAPIKey_fromTable()
  handleAccessToken(API_KEY)

  // show lại chức năng đề phòng bị ẩn do hết hạn access token
  document.getElementsByClassName("gr-accessTokenNotExpired").forEach(element => {
    element.style.display = 'block'
  });
  document.getElementsByClassName("gr-accessTokenExpired").forEach(element => {
    element.style.display = 'none'
  });
}

//<<==================================================== USER INFO ===============================================>>

document.getElementById("button_getUserInfo").onclick = function () {
  // check 401
  checkError401(function (checked) {
    if (!checked) {
      setLoading(true)
      getUserInfo()
    }
  })

}

//#region <<============================================== TRANSACTION ==========================================>>
globalThis.Check = 0
async function delay() {
  for (let i = 0; i <= 100; i++) {


    // console.log(this.returnValue)


  }
}

document.getElementById("button_getTransaction").onclick = function () {
  //check 401
  checkError401(function (checked) {
    if (!checked) {
      var txtDate = document.getElementById("txtDate")
      console.log(txtDate.value)
      if (!txtDate.value)
        swal("Error", "Please enter information about  combobox");
      else {
        ClearAllData("Transaction")
        typeChart(function () {
          if (this.returnValue != "default") {
            let type = this.returnValue
            addChartSheet(function (value) {
              console.warn(value)
              getTransaction(txtDate.value, type, value)
              setLoading(true) // SHOW LOADING
            })
          }
          else {
            console.log(this.returnValue)
            getTransaction(txtDate.value, this.returnValue, "default")
            setLoading(true) // SHOW LOADING
          }
        })
      }
    }
  })
}

document.getElementById("button_getLatestTrans").onclick = function () {
  //check 401
  checkError401(function (checked) {
    if (!checked) {
      setLoading(true) // SHOW LOADING
      checkLatestTrans()
    }
  })
}

//#endregion

//#region <<============================================ HANDLE APIKEY ==========================================>>
var text_apiKey = document.getElementById("text_apiKey")
document.getElementById("button_submit_apiKey").onclick = function (e) {
  if (text_apiKey.value == "")
    swal("Please enter Api Key");
  else {
    //API_KEY = text_apiKey.value
    loadingPage()

    document.getElementById("api_key").style.display = "none"
    document.getElementById("main").style.display = "none"
    showLoading()
  }
}
//#endregion

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
  document.getElementById("loader").style.display = "block"
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


//#region =========================================== CHECK 401 ERROR ==================================
function checkError401(callback) {
  var url = URL_ROOT + "transactions"
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
      if (respond.status === 401) {
        console.warn('hieu 401')
        // display: block 2 --> 2 chắc năng và hiển thị đoạn text hết hạn access token
        document.getElementsByClassName("gr-accessTokenNotExpired").forEach(element => {
          element.style.display = 'none'
        });
        document.getElementsByClassName("gr-accessTokenExpired").forEach(element => {
          element.style.display = 'block'
        });

        return true
      }
      return false
    }).then(callback)

}
//#endregion
