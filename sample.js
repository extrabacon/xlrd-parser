var xlrd = require('xlrd-parser');

xlrd.parse('myfile.xlsx', function (err, workbook) {
	// Iterate on sheets
	workbook.sheets.forEach(function (sheet) {
		console.log('sheet: ' + sheet.name);
		// Iterate on rows
		sheet.rows.forEach(function (row) {
			// Iterate on cells
			row.forEach(function (cell) {
				console.log(cell.address + ': ' + cell.value);
			});
		});
	});
});

xlrd.stream('myfile.xlsx').on('open', function (workbook) {
	console.log('successfully opened ' + workbook.file);
}).on('data', function (data) {

	var currentWorkbook = data.workbook,
		currentSheet = data.sheet,
		batchOrRows = data.rows;

	// TODO: handle streaming logic here

}).on('error', function (err) {
	// TODO: handle error here
}).on('close', function () {
	// TODO: finishing logic here
});
