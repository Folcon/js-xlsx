var XLSX = require('./xlsx');
var OUTFILE = 'tmp/file-styling.xlsx';
var INFILE = 'tmp/file.xlsx';
var filedata = XLSX.readFile(INFILE);
var sheetData = filedata.Sheets.Sheet1;
console.log("sheet data" + sheetData);

for (var key in sheetData) {
    if (sheetData.hasOwnProperty(key)) {
        sheetData[key].s = {
            "font": {
              "color": {
                "rgb": "FF0000FF"
              }
            }
          };
        console.log(key + " -> " + JSON.stringify( sheetData[key].s));
    }
}
var defaultCellStyle = { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}};
var wopts = { bookType:'xlsx', bookSST:false, type:'binary', defaultCellStyle: defaultCellStyle, showGridLines: true};
XLSX.writeFile(filedata, OUTFILE, wopts);