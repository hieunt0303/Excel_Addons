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