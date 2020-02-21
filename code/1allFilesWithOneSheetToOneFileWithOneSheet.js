var xlsx = require('xlsx');
var fs = require('fs');
var path = require('path');

var inputFolder = 'input'; //Input folder

var rootFolder = path.dirname(__dirname);
var inputDir = path.join(rootFolder, inputFolder);
var inputFiles = fs.readdirSync(inputDir);

function readFileToJson(fileName) {
	var wb = xlsx.readFile(fileName, { cellDates: true });
	var firstTabName = wb.SheetNames[0];
	var ws = wb.Sheets[firstTabName];
	var data = xlsx.utils.sheet_to_json(ws);
	return data;
}

var combinedData = [];

(function putDataToArray() {
	inputFiles.forEach(function(file) {
		var fileExtension = path.parse(file).ext;
		if (fileExtension === '.xlsx' && file[0] !== '~') {
			var fullFilePath = path.join(inputDir, file);
			var data = readFileToJson(fullFilePath);
			combinedData = combinedData.concat(data);
		}
	});
})();

var newWB = xlsx.utils.book_new();

(function writeDataToOneSheet() {
	var newWS = xlsx.utils.json_to_sheet(combinedData);
	xlsx.utils.book_append_sheet(newWB, newWS, 'Combined Data'); //New sheet name
})();

(function saveFile() {
	xlsx.writeFile(newWB, '1Combined.xlsx'); //New combined file name
	fs.renameSync('1Combined.xlsx', path.join(rootFolder, '1Combined.xlsx'), function(err) {
		if (err) throw err;
	});
	console.log('Done');
})();
