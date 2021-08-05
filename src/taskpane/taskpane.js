import { get } from "../../src/function/getAccessToken.js"
import { getUserInfo } from "../../src/function/UserInfo.js"
import { getTransaction, lastTransPage, lastTransDate, filterPage } from "../../src/function/Transaction.js"
import swal from 'sweetalert';
import { ClearAllData, addInitialSheet } from "../function/Excelfunction.js"

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



// CHẠY HÀM NÀY ĐỂ CÓ 2 SHEET BAN ĐẦU ĐỂ THAO TÁC
//addInitialSheet()

// radio check selection
var check_selectTime = document.getElementById("check_selectTime")
var check_DateToDate = document.getElementById("check_DateToDate")
// check_DateToDate.onclick = check_transaction()
// check_DateToDate.onchange = check_transaction()
// check_selectTime.onclick = check_transaction()
// check_selectTime.onchange = check_transaction()

var rad = document.myForm.myRadios;
var prev = null;
for (var i = 0; i < rad.length; i++) {
    rad[i].addEventListener('change', check_transaction());
}
function check_transaction() {
  if (check_DateToDate.checked) {
    document.getElementById("gr_radio_fromDatetoDate").style.display = "block"
    document.getElementById("gr_radio_selectDate").style.display = "none"

  }
  else if (check_selectTime.checked) {
    document.getElementById("gr_radio_fromDatetoDate").style.display = "none"
    document.getElementById("gr_radio_selectDate").style.display = "block"
  }
}

var elementExpired = document.getElementById("header_accesstoken_expired")
var elementExist = document.getElementById("header_accesstoken_exist")

document.getElementById("button_reload_APIKey").onclick = function () {

  console.log(filterPage(2, {}))
}
document.getElementById("button_getUserInfo").onclick = function () {
  getUserInfo()
  //console.log('hieu')
}

document.getElementById("button_getTransaction").onclick = function () {

  //getTransaction()
  if (optionItem.value == "0")
    text_alert.style.display = "block"
  else {
    text_alert.style.display = "none"

    swal("Do you want do delete old data?", {
      buttons: ["No", "Yes"],
    })
      .then(function (result) {
        if (result) {
          // Yes
          ClearAllData("Transaction")

        }
        else {
          //No
          console.log("2")

        }
      })

  }

}

var optionItem = document.getElementById("selectItem")
var text_alert = document.getElementById("text_alert")

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
