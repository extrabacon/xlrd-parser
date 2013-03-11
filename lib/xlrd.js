var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util'),
	spawn = require('child_process').spawn;

function XlrdParser(path, options) {

	var self = this,
		runxlrd = spawn('python', ['deps/python-excel-xlrd/runxlrd-json.py', 'show', path]),
		workbook,
		sheet,
		row,
		rowindex = -1,
		rows = [],
		remaining = '';

	options = options || {
		debug: false
	};

	function flushRows() {
		if (rows.length > 0) {
			// Emit an event with the accumulated rows
			self.emit('data', {
				workbook: workbook,
				sheet: sheet,
				rows: rows.slice(0)
			});
			// Reset rows for the next iteration
			rows = [];
		}
	}

	EventEmitter.call(this);

	runxlrd.stdout.setEncoding('utf8');
	runxlrd.stdout.on('data', function (data) {

		var lines = data.split(/\n/g),
			lastLine = _.last(lines);

		if (options.debug) {
			console.log(data);
		}

		// Fix the first line with the remaining from the previous iteration of 'data'
		lines[0] = remaining + lines[0];
		// Keep the remaining for the next iteration of 'data'
		remaining = lastLine;

		_.initial(lines).forEach(function (line) {
			var record = JSON.parse(line),
				value = record[1];
			switch (record[0]) {
				case 'workbook':
					workbook = {
						file: value.file,
						meta: {
							user: value.user,
							sheets: {
								count: value.sheets.count,
								names: value.sheets.names
							}
						}
					};
					self.emit('open', workbook);
					break;
				case 'sheet':
					if (sheet != null) {
						// if we are changing sheets, flush rows immediately
						flushRows();
					}
					sheet = {
						index: value.index,
						name: value.name,
						bounds: {
							rows: value.rows,
							columns: value.cols
						},
						visibility: (function () {
							switch (value.visibility) {
								case 1: return 'hidden';
								case 2: return 'very hidden';
								default: return 'visible';
							}
						})()
					};
					break;
				case 'cell':
					if (record[1].r !== rowindex) {
						// switch rows
						row = [];
						rowindex = value.r;
						rows.push(row);
					}
					// append cell to current row
					row.push({
						row: value.r,
						column: value.c,
						address: value.cn + (value.r + 1),
						value: (function () {
							if (value.t == 3) { // XL_CELL_DATE
								return new Date(
									value.v[0], value.v[1] - 1, value.v[2],
									value.v[3], value.v[4], value.v[5]
								);
							} else if (value.t == 5) { // XL_CELL_ERROR
								return new Error(value.v);
							} else if (value.t == 6) { // XL_CELL_BLANK
								return undefined;
							} else {
								return value.v;
							}
						})()
					});
					break;
				case 'error':
					self.emit('error', _.extend(new Error(value.type + ': ' + value.message), value));
					break;
			}
		});

		flushRows();
	});

	runxlrd.on('exit', function (code) {

		if (code != 0) {
			self.emit('error', new Error('exit code ' + code + ' returned from runxlrd-json'));
		}

		self.emit('close');
	});
}

util.inherits(XlrdParser, EventEmitter);

exports.stream = function (path) {
	return new XlrdParser(path);
};

exports.parse = function (path, callback) {

	// with parse, we assemble a workbook/sheet/row/cell structure

	var reader = new XlrdParser(path),
		workbooks = [],
		errors;

	reader.on('data', function (data) {

		var workbook = _.find(workbooks, function (w) {
			return w.file === data.workbook.file;
		});

		if (!workbook) {
			workbook = _.extend({ sheets: [] }, data.workbook);
			workbooks.push(workbook);
		}

		var sheet = _.find(workbook.sheets, function (s) {
			return s.index === data.sheet.index;
		});

		if (!sheet) {
			sheet = _.extend({ rows: [] }, data.sheet);
			workbook.sheets.push(sheet);
			workbook.sheets[sheet.name] = sheet;
		}

		data.rows.forEach(function (row) {
			sheet.rows.push(row);
		});

	});

	reader.on('error', function (err) {
		if (!errors) {
			errors = err;
		} else {
			errors = [errors];
			errors.push(err);
		}
	});

	reader.on('close', function () {

		callback = callback || function () {};

		if (!errors && !workbooks.length) {
			var message = 'file not found: ' + path;
			errors = _.extend(new Error(message), { type: 'file_not_found', message: message });
		}

		if (workbooks.length === 0) {
			callback(errors, null);
		} else if (workbooks.length === 1) {
			callback(errors, workbooks[0])
		} else {
			callback(errors, workbooks);
		}
	});

};
