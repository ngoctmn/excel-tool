var xlsx = require('xlsx');
var fs = require('fs');
var path = require('path');

var inputFolder = 'input'; //Input folder

var rootFolder = path.dirname(__dirname);
var inputDir = path.join(rootFolder, inputFolder);
var inputFiles = fs.readdirSync(inputDir);

var combinedData = [];
var sheetName = [];

function readFileToJson(fileName) {
	var wb = xlsx.readFile(fileName, { cellDates: true });
	for (var sheet in Object.keys(wb.Sheets)) {
		var ws = Object.values(wb.Sheets)[sheet];
		var data = xlsx.utils.sheet_to_json(ws);
		combinedData.push(data);
	}
	for (var sheetTab in wb.SheetNames) {
		sheetName.push(wb.SheetNames[sheetTab]);
	}
}

(function pushDataToArray() {
	inputFiles.forEach(function(file) {
		var fileExtension = path.parse(file).ext;
		if (fileExtension === '.xlsx' && file[0] !== '~') {
			var fullFilePath = path.join(inputDir, file);
			readFileToJson(fullFilePath);
			return combinedData;
		}
	});
})();

var newWB = xlsx.utils.book_new();

(function writeDataToSheets() {
	for (var i = 0; i < combinedData.length; i++) {
		var newWS = xlsx.utils.json_to_sheet(combinedData[i]);
		xlsx.utils.book_append_sheet(newWB, newWS, sheetName[i]); //New sheet name
	}
})();

(function saveFile() {
	xlsx.writeFile(newWB, '3Combined.xlsx'); //New combined file name
	fs.renameSync('3Combined.xlsx', path.join(rootFolder, '3Combined.xlsx'), function(err) {
		if (err) throw err;
	});
	console.log('Done');
})();
