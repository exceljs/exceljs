/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var fs = require('fs');
var csv = require('fast-csv');
var moment = require('moment');
var PromishLib = require('../utils/promish');
var StreamBuf = require('../utils/stream-buf');

var utils = require('../utils/utils');

var CSV = module.exports = function (workbook) {
  this.workbook = workbook;
  this.worksheet = null;
};

/* eslint-disable quote-props */
var SpecialValues = {
  'true': true,
  'false': false,
  '#N/A': { error: '#N/A' },
  '#REF!': { error: '#REF!' },
  '#NAME?': { error: '#NAME?' },
  '#DIV/0!': { error: '#DIV/0!' },
  '#NULL!': { error: '#NULL!' },
  '#VALUE!': { error: '#VALUE!' },
  '#NUM!': { error: '#NUM!' }
};
/* eslint-ensable quote-props */

CSV.prototype = {
  readFile: function readFile(filename, options) {
    var self = this;
    options = options || {};
    var stream;
    return utils.fs.exists(filename).then(function (exists) {
      if (!exists) {
        throw new Error('File not found: ' + filename);
      }
      stream = fs.createReadStream(filename);
      return self.read(stream, options);
    }).then(function (worksheet) {
      stream.close();
      return worksheet;
    });
  },
  read: function read(stream, options) {
    var _this = this;

    options = options || {};
    return new PromishLib.Promish(function (resolve, reject) {
      var csvStream = _this.createInputStream(options).on('worksheet', resolve).on('error', reject);

      stream.pipe(csvStream);
    });
  },
  createInputStream: function createInputStream(options) {
    options = options || {};
    var worksheet = this.workbook.addWorksheet(options.sheetName);

    var dateFormats = options.dateFormats || [moment.ISO_8601, 'MM-DD-YYYY', 'YYYY-MM-DD'];
    var map = options.map || function (datum) {
      if (datum === '') {
        return null;
      }
      if (!isNaN(datum)) {
        return parseFloat(datum);
      }
      var dt = moment(datum, dateFormats, true);
      if (dt.isValid()) {
        return new Date(dt.valueOf());
      }
      var special = SpecialValues[datum];
      if (special !== undefined) {
        return special;
      }
      return datum;
    };

    var csvStream = csv(options).on('data', function (data) {
      worksheet.addRow(data.map(map));
    }).on('end', function () {
      csvStream.emit('worksheet', worksheet);
    });
    return csvStream;
  },

  write: function write(stream, options) {
    var _this2 = this;

    return new PromishLib.Promish(function (resolve, reject) {
      options = options || {};
      // var encoding = options.encoding || 'utf8';
      // var separator = options.separator || ',';
      // var quoteChar = options.quoteChar || '\'';

      var worksheet = _this2.workbook.getWorksheet(options.sheetName || options.sheetId);

      var csvStream = csv.createWriteStream(options);
      stream.on('finish', function () {
        resolve();
      });
      csvStream.on('error', reject);
      csvStream.pipe(stream);

      var dateFormat = options.dateFormat;
      var dateUTC = options.dateUTC;
      var map = options.map || function (value) {
        if (value) {
          if (value.text || value.hyperlink) {
            return value.hyperlink || value.text || '';
          }
          if (value.formula || value.result) {
            return value.result || '';
          }
          if (value instanceof Date) {
            if (dateFormat) {
              dateUTC ? moment.utc(value).format(dateFormat) : moment(value).format(dateFormat);
            }
            return dateUTC ? moment.utc(value).format() : moment(value).format();
          }
          if (value.error) {
            return value.error;
          }
          if ((typeof value === 'undefined' ? 'undefined' : _typeof(value)) === 'object') {
            return JSON.stringify(value);
          }
        }
        return value;
      };

      var includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
      var lastRow = 1;
      if (worksheet) {
        worksheet.eachRow(function (row, rowNumber) {
          if (includeEmptyRows) {
            while (lastRow++ < rowNumber - 1) {
              csvStream.write([]);
            }
          }
          var values = row.values;
          values.shift();
          csvStream.write(values.map(map));
          lastRow = rowNumber;
        });
      }
      csvStream.end();
    });
  },
  writeFile: function writeFile(filename, options) {
    options = options || {};

    var streamOptions = {
      encoding: options.encoding || 'utf8'
    };
    var stream = fs.createWriteStream(filename, streamOptions);

    return this.write(stream, options);
  },
  writeBuffer: function writeBuffer(options) {
    var self = this;
    var stream = new StreamBuf();
    return self.write(stream, options).then(function () {
      return stream.read();
    });
  }
};
//# sourceMappingURL=csv.js.map
