# xlrd-parser

High performance Excel file parser based on the xlrd library from [www.python-excel.org](www.python-excel.org) for
reading Excel files in XLS or XLSX formats.

This module interfaces with a Python shell to stream JSON from stdout using a child process. It is not a port of xlrd
from Python to Javascript (which is surely possible).

## Features

+ Much faster and more memory efficient than most alternatives
+ Support for both XLS and XLSX formats
+ Can read multiple sheets
+ Can load data selectively
+ Can stream large files (tested with 250K+ rows)

## Documentation

### Installation
```bash
npm install xlrd-parser
```

### Parsing a file

Parsing a file loads the entire file into an object structure composed of a workbook, sheets, rows and cells.

```javascript
var xlrd = require('xlrd');

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
```

Cell values are parsed as strings, numbers and dates depending on the cell format.

For more details on the API, see the included unit tests.

### Streaming a large file

For large files, you may want to stream the data. The stream method returns a familiar EventEmitter instance.

```javascript
var xlrd = require('xlrd');

xlrd.stream('myfile.xlsx').on('open', function (workbook) {
	console.log('successfully opened ' + workbook.file);
}).on('data', function (data) {

	var currentWorkbook = data.workbook,
		currentSheet = data.sheet,
		batchOfRows = data.rows;

	// TODO: handle streaming logic here

}).on('error', function (err) {
	// TODO: handle error here
}).on('close', function () {
	// TODO: finishing logic here
});
```

### Options

An object can be passed to the `parse` and `stream` methods to define additional options.

* `meta` - loads only workbook metadata, without iterating on rows - `Boolean`
* `sheet` || `sheets` - loads sheets selectively, either by name or by index - `String`, `Number` or `Array`
* `maxRows` - the maximum number of rows to load per sheet - `Number`
* `debug` - outputs JSON from the xlrd-parser child process - `Boolean`

#### Examples:

Output sheet names without loading any data

```javascript
xlrd.parse('myfile.xlsx', { meta: true }, function (err, workbook) {
  console.log(workbook.meta.sheets);
});
```

Load only the first 10 rows from the first sheet

```javascript
xlrd.parse('myfile.xlsx', { sheets: 0, maxRows: 10 }, function (err, workbook) {
  // workbook will contain only the first sheet
});
```

Load only a sheet named 'products'

```javascript
var stream = xlrd.stream('myfile.xlsx', { sheets: 'products' });
```

## Compatibility

+ Tested with Node 0.10.x
+ Tested on Mac OS X 10.8
+ Tested on Ubuntu Linux 12.04 (requires prior installation of curl: apt-get install curl)

## Dependencies

+ Python version 2.6+
+ xlrd version 0.7.4+
+ underscore.js
+ bash (installation script)
+ curl (installation script)

Windows platform is not yet supported, but it is only a matter of converting the installation script to PowerShell.
A Python shell also needs to be available from the command line, which could be installed via
[chocolatey](http://chocolatey.org/packages/python).

## Changelog

### 0.1.0

+ can probe a workbook for metadata, without iterating on rows (options.meta)
+ can specify which sheet to load, either by name or by index (options.sheets)
+ can specify maximum number of rows to load (options.maxRows)
+ parsing of errors, such as #DIV/0!, #NAME? or #VALUE!
+ added JSDoc comments
+ simplified and optimized Python script, emitted JSON is more compact which should further reduce memory footprint
+ sheets are now loaded on-demand
+ better error handling

### 0.0.1

+ initial release


## Limitations

+ Does not parse formatting info (performance impact)

## Thanks

Many thanks to the authors of the xlrd library ([here](http://github.com/python-excel/xlrd)). It is the best and most
efficient open-source library I could find.

## License

The package itself is MIT licenced.

License from xlrd library:

	Portions copyright © 2005-2009, Stephen John Machin, Lingfo Pty Ltd
	All rights reserved.

	Redistribution and use in source and binary forms, with or without
	modification, are permitted provided that the following conditions are met:

	1. Redistributions of source code must retain the above copyright notice,
	this list of conditions and the following disclaimer.

	2. Redistributions in binary form must reproduce the above copyright notice,
	this list of conditions and the following disclaimer in the documentation
	and/or other materials provided with the distribution.

	3. None of the names of Stephen John Machin, Lingfo Pty Ltd and any
	contributors may be used to endorse or promote products derived from this
	software without specific prior written permission.

	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
	AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
	THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
	PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS
	BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
	CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
	SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
	INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
	CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
	ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF
	THE POSSIBILITY OF SUCH DAMAGE.

	/*-
	 * Copyright (c) 2001 David Giffin.
	 * All rights reserved.
	 *
	 * Based on the the Java version: Andrew Khan Copyright (c) 2000.
	 *
	 *
	 * Redistribution and use in source and binary forms, with or without
	 * modification, are permitted provided that the following conditions
	 * are met:
	 *
	 * 1. Redistributions of source code must retain the above copyright
	 *    notice, this list of conditions and the following disclaimer.
	 *
	 * 2. Redistributions in binary form must reproduce the above copyright
	 *    notice, this list of conditions and the following disclaimer in
	 *    the documentation and/or other materials provided with the
	 *    distribution.
	 *
	 * 3. All advertising materials mentioning features or use of this
	 *    software must display the following acknowledgment:
	 *    "This product includes software developed by
	 *     David Giffin <david@giffin.org>."
	 *
	 * 4. Redistributions of any form whatsoever must retain the following
	 *    acknowledgment:
	 *    "This product includes software developed by
	 *     David Giffin <david@giffin.org>."
	 *
	 * THIS SOFTWARE IS PROVIDED BY DAVID GIFFIN ``AS IS'' AND ANY
	 * EXPRESSED OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
	 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
	 * PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL DAVID GIFFIN OR
	 * ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	 * SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
	 * NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
	 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
	 * HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
	 * STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
	 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
	 * OF THE POSSIBILITY OF SUCH DAMAGE.
	 */
