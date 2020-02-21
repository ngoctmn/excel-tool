var xlsx = require('xlsx');
var fs = require('fs');
var path = require('path');

var inputFolder = 'input'; //Input folder

var rootFolder = path.dirname(__dirname);
var inputDir = path.join(rootFolder, inputFolder);
var inputFiles = fs.readdirSync(inputDir);

var combinedData = [];

function readFileToJson(fileName) {
	var wb = xlsx.readFile(fileName, { cellDates: true });
	for (var sheet in Object.keys(wb.Sheets)) {
		var ws = Object.values(wb.Sheets)[sheet];
		var data = xlsx.utils.sheet_to_json(ws);
		combinedData.push(data);
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

totalData = [];
(function writeDataToOneSheet() {
	for (var i = 0; i < combinedData.length; i++) {
		totalData = totalData.concat(combinedData[i]);
	}
	var newWS = xlsx.utils.json_to_sheet(totalData);
	xlsx.utils.book_append_sheet(newWB, newWS, 'Combined Data'); //New sheet name
})();

(function saveFile() {
	xlsx.writeFile(newWB, '2Combined.xlsx'); //New combined file name
	fs.renameSync('2Combined.xlsx', path.join(rootFolder, '2Combined.xlsx'), function(err) {
		if (err) throw err;
	});
	console.log('Done');
})();
