var XLSX = require('xlsx');
var excelbuilder = require('msexcel-builder');
var program = require('commander');

program
    .version('0.0.1')
    .option('-f, --infile [type]', 'Define the input file', 'sample')
    .option('-o, --outfile [type]', 'Define the output file', 'output')
    .parse(process.argv);

/*Excel sheet parsing*/
var inFileName =  'input/' + program.infile + ".xlsx";
var outFileName =  'output/' + program.outfile + ".xlsx";
var workbook = XLSX.readFile(inFileName);
var sheet_name_list = workbook.SheetNames;

var empReportObj = [];
sheet_name_list.forEach(function(sheetName) {
    var worksheet = workbook.Sheets[sheetName];

    var valuesPerMemArray = [];
    var valuesPerMem = {};
    var reportIndex = 0;
    var arr = [];

    var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
    for(var i = 0; i<roa.length; i++){
        var row = roa[i];
        var rowMemNo = row["MEMNO"];
        var rowConNo = row["con"];
        var rowTotalCount = row["CONT"];

        if(rowMemNo == "Grand Total"){
            var reportObjWraper = {};
            reportObjWraper.grandTotal = roa[i+1]["con"];
            reportObjWraper.pages = valuesPerMemArray;
            empReportObj[reportIndex++] = reportObjWraper;
            valuesPerMemArray = [];
            valuesPerMem = {};
        } else {
            if(rowMemNo != undefined && rowMemNo.indexOf('page') != -1){
                var pageObj = {};
                pageObj.name = rowMemNo;
                pageObj.values = arr;

                if(roa[i+1]["MEMNO"] == "Grand Total"){
                    pageObj.totalCount = roa[i-1]["CONT"];
                } else {
                    pageObj.totalCount = rowTotalCount;
                }
                valuesPerMem.pageKey = rowMemNo;
                valuesPerMem.pageValue = pageObj;
                valuesPerMemArray.push(valuesPerMem);
                valuesPerMem = {};
                arr = [];
            } else if(rowMemNo != undefined && rowConNo != undefined && rowMemNo != "MEMNO"){
                var memValue = {};
                memValue.key = rowMemNo;
                memValue.value = rowConNo;
                arr.push(memValue);
            }
        }
    }
});

// Create a new workbook file in current working-path
var workbook = excelbuilder.createWorkbook('./', outFileName)

/*CHANGE NUM OF ROWS AND COLUMNS ACCORDINGLY*/
// Create a new worksheet with 10 columns and 12 rows
var sheet1 = workbook.createSheet('sheet1', 20, 1000);

var colIndex = 0;
var rowIndex = 0;

for (var j = 0; j < empReportObj.length; j++) {

    var singleEmpReportObject = empReportObj[j];
    var pages = singleEmpReportObject.pages;
    var grandTotal = singleEmpReportObject.grandTotal;

    for (var k = 0; k < pages.length; k++) {
        var page = pages[k];

        var pageValue = page.pageValue;
        var pageKey = page.pageKey;

        var pageValues = pageValue.values;
        var pageName = pageValue.name;

        var pageTotalCount = pageValue.totalCount;

        for (var i = 0; i < pageValues.length; i++) {
            if((i % 5) == 0){
                rowIndex++;
                colIndex = 0;
            }
            var pageValue = pageValues[i];
            sheet1.set(++colIndex, rowIndex, pageValue.key);
            sheet1.set(++colIndex, rowIndex, pageValue.value);
        }
        rowIndex = rowIndex + 2;
        sheet1.set(1, rowIndex, pageName);
        sheet1.set(3, rowIndex, pageTotalCount);

        rowIndex = rowIndex + 2;
    }

    sheet1.set(1, rowIndex, "Grand Total");
    sheet1.set(3, rowIndex, grandTotal);

    rowIndex = rowIndex + 4;
}

// Save it
workbook.save(function(ok){
    if (!ok)
        workbook.cancel();
    else
        console.log('congratulations, your workbook created');
});

module.exports = {};
