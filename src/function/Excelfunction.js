export function ClearAllData(SHEET){
    if(!SHEET){
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            sheet.getRange().clear();
            return ctx.sync();
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    else{
        Excel.run(function (ctx) {
            var sheet = ctx.workbook.worksheets.getItem(SHEET);
            sheet.getRange().clear();
            return ctx.sync();
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
}
// Xóa các sheet tồn tại và thêm vào 2 sheet cần thiết trans và userinfo
export function addInitialSheet(){
    renameCurrentSheet("Transaction")
    
}

function renameCurrentSheet(nameSheet){
    Excel.run(function (context) {
        var currentSheet = context.workbook.worksheets.getActiveWorksheet();
        currentSheet.name = nameSheet;
    
        return context.sync();
    })
    .catch(function(error){
        add1Sheet("UserInfo")
        add1Sheet("Handle API")
    })
    .then(function(e){
        add1Sheet("UserInfo")
    })

}

function add1Sheet(nameSheet){
    Excel.run(function (context) {
        var sheets = context.workbook.worksheets;
    
        var sheet = sheets.add(nameSheet);
        sheet.load("name, position");
    
        return context.sync()
            .then(function () {
                console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
            });
    }).catch(errorHandlerFunction);
}

// Handler date --> show transaction
// records = [[],[],[],[]]
export function showTrans(TOTAL_TRANSACTION){
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("Transaction")
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

