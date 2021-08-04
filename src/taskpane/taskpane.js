import { API_KEY, URL_ROOT, ACCESS_TOKEN } from "../valueConst.js"
import { get } from "../../src/function/getAccessToken.js"
import { getUserInfo } from "../../src/function/UserInfo.js"
import { getTransaction, lastTransPage, lastTransDate } from "../../src/function/Transaction.js"
import swal from 'sweetalert';



/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

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
      getAPI()
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}


function getAPI() {
  const url = URL_ROOT + "userInfo";
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
    .then((resp) => resp.json())
    .then(function (data) {
      console.log(data);
      if (data.error == "401")
        // run function to get APIKey

        get(URL_ROOT + "token", API_KEY)
    })
    .catch(function (error) {
      console.log(error);
    });
}


var elementExpired = document.getElementById("header_accesstoken_expired")
var elementExist = document.getElementById("header_accesstoken_exist")


document.getElementById("button_reload_APIKey").onclick = function () {

}
document.getElementById("button_getUserInfo").onclick = function () {
  getUserInfo()
  //console.log('hieu')
}
document.getElementById("button_getTransaction").onclick = function () {
  //getTransaction()
  swal("Do you want do delete old data?", {
    buttons: ["No", "Yes"],
  })
    .then(function (result) {
      if (result) {
        // Yes
        console.log("1")
      }
      else {
        //No
        console.log("2")
      }
    })
  
}

var optionItem = document.getElementById("selectItem")

document.getElementById("selectItem").onchange = function () {
  console.log(document.getElementById("selectItem").value)
}

// CHẠY HÀM NÀY ĐỂ LOAD RA CÁI COMBOBOX HIỂN THỊ CHỌN NGÀY CHO GIAO DỊCH 
loadComboboxTrans()

function loadComboboxTrans() {

  lastTransPage(function (lastPage) {
    lastTransDate(lastPage, function (lastDate) {
      createDataCombo(convertMonthDate(lastDate))
    })
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