var XLSX = require('./xlsx');
var OUTFILE = 'tmp/file-styling.xlsx';
var INFILE = 'tmp/file.xlsx';
var filedata = XLSX.readFile(INFILE);
var sheetData = JSON.stringify(filedata.Sheets.Sheet1);
var x = filedata.Sheets.Sheet1;
console.log("sheet data" + sheetData);

for (var key in x) {
    if (x.hasOwnProperty(key)) {
        x[key].s = {
            "font": {
              "color": {
                "rgb": "FFC6EFCE"
              }
            }
          };
        console.log(key + " -> " + JSON.stringify( x[key].s));
    }
}