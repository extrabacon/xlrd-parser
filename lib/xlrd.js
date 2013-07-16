var _ = require('underscore'),
	EventEmitter = require('events').EventEmitter,
	util = require('util'),
	spawn = require('child_process').spawn;

function XlrdParser(path, options) {

	var self = this,
		runxlrd = spawn('python', formatArgs()),
		workbook,
		sheet,
		row,
		rowindex = -1,
		rows = [],
		remaining = '';

	function formatArgs() {

		var args = ['deps/python-excel-xlrd/xlrd-parser.py'];
		options = options || { debug: false };

		if (options.meta) {
			args.push('-m');
		}
		if (options.sheets) {
			if (_.isArray(options.sheets)) {
				args.push(_.map(options.sheets, function (s) { return ['-s', s] }));
			} else {
				args.push(['-s', options.sheets]);
			}
		}
		if (options.maxRows) {
			args.push(['-r', options.maxRows]);
		}

		args.push(path);
		return _.flatten(args);
	}

	function parseValue(value) {

		if (_.isArray(value)) {
			// parse non-native data types
			if (value[0] == 'date') {
				return new Date(
					value[1], value[2] - 1, value[3],
					value[4], value[5], value[6]
				);
			} else if (value[0] == 'error') {
				return new Error(value[1]);
			} else if (value[0] == 'empty') {
				return null;
			}
		}

		return value;
	}

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
				case 'w':
					workbook = {
						file: value.file,
						meta: {
							user: value.user,
							sheets: value.sheets
						}
					};
					self.emit('open', workbook);
					break;
				case 's':
					if (sheet != null) {
						// if we are changing sheets, flush rows immediately
						flushRows();
					}
					sheet = {
						index: value.index,
						name: value.name,
						bounds: {
							rows: value.rows,
							columns: value.columns
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
				case 'c':
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
						address: value.a,
						value: parseValue(value.v)
					});
					break;
				case 'err':
					self.emit('error', _.extend(
						new Error('[' + value.exception + '] ' + value.type + ': ' + value.details),
						value)
					);
					break;
			}
		});

		flushRows();
	});

	runxlrd.on('exit', function (code) {

		if (code != 0) {
			self.emit('error', _.extend(
				new Error('exit code ' + code + ' returned from xlrd-parser.py'),
				{ code: code })
			);
		}

		self.emit('close');
	});
}

util.inherits(XlrdParser, EventEmitter);

/**
 * Parses a workbook, returning an EventEmitter instance for loading the results
 *
 * @param {String|Array} path The path of the file(s) to load
 * @param {Object} options The options object (optional)
 * @returns {EventEmitter} The instance to use to stream the results
 */
exports.stream = function (path, options) {
	return new XlrdParser(path, options);
};

/**
 * Parses a workbook, returning a workbook/sheet/row/cell structure
 *
 * @param {String|Array} path The path of the file(s) to load
 * @param {Object} options The options object (optional)
 * @param {requestCallback} callback The callback method to invoke with the results
 */
exports.parse = function (path, options, callback) {

	if (_.isFunction(options)) {
		callback = options;
		options = null;
	}

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
			errors = _.extend(new Error('file not found: ' + path), { id: 'file_not_found' });
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
