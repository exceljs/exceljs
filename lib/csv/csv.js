/**
 * Copyright (c) 2015 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';


var fs = require('fs');
var csv = require('fast-csv');
var moment = require('moment');

var utils = require('../utils/utils');

var CSV = module.exports = function(workbook) {
  this.workbook = workbook;
  this.worksheet = null;
};

CSV.prototype = {
  readFile: function(filename, options) {
    var self = this;
    options = options || {};
    var stream;
    return utils.fs.exists(filename)
      .then(function(exists) {
        if (!exists) {
          throw new Error('File not found: ' + filename);
        }
        stream = fs.createReadStream(filename);
        return self.read(stream, options);
      })
      .then(function(worksheet) {
        stream.close();
        return worksheet;
      });
  },
  read: function(stream, options) {
    options = options || {};
    var self = this;
    return new Promise(function(resolve, reject) {

      var csvStream = self.createInputStream(options)
        .on('worksheet', function(worksheet) {
          resolve(worksheet);
        })
        .on('error', function(error) {
          reject(error);
        });

      stream.pipe(csvStream);

    });

  },
  createInputStream: function(options) {
    options = options || {};
    var worksheet = this.workbook.addWorksheet(options.sheetName);

    var dateFormats = options.dateFormats || [
        moment.ISO_8601,
        'MM-DD-YYYY',
        'YYYY-MM-DD'
      ];
    var map = options.map || function(datum) {
        if (datum == '') {
          return null;
        }
        if (!isNaN(datum)) {
          return parseFloat(datum);
        }
        var dt = moment(datum, dateFormats, true);
        if (dt.isValid()) {
          return new Date(dt.valueOf());
        }
        return datum;
      };

    var csvStream = csv(options)
      .on('data', function(data){
        worksheet.addRow(data.map(map));
      })
      .on('end', function() {
        csvStream.emit('worksheet', worksheet);
      });
    return csvStream;
  },

  write: function(stream, options) {

    var self = this;

    return new Promise(function(resolve, reject) {
      options = options || {};
      //var encoding = options.encoding || 'utf8';
      //var separator = options.separator || ',';
      //var quoteChar = options.quoteChar || '\'';

      var worksheet = self.workbook.getWorksheet(options.sheetName || options.sheetId);

      var csvStream = csv.createWriteStream(options);
      stream.on('finish', function() {
        resolve();
      });
      csvStream.on('error', function(error) {
        reject(error);
      });
      csvStream.pipe(stream);

      var dateFormat = options.dateFormat;
      var map = options.map || function(value) {
          if (value) {
            if (value.text || value.hyperlink) {
              return value.hyperlink || value.text || '';
            }
            if (value.formula || value.result) {
              return value.result || '';
            }
            if (value instanceof Date) {
              return dateFormat ? moment(value).format(dateFormat) : moment(value).format();
            }
            if (typeof value === 'object') {
              return JSON.stringify(value);
            }
          }
          return value;
        };

      var includeEmptyRows = (options.includeEmptyRows === undefined) || options.includeEmptyRows;
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
  writeFile: function(filename, options) {
    options = options || {};

    var streamOptions = {
      encoding: options.encoding || 'utf8'
    };
    var stream = fs.createWriteStream(filename, options);

    return this.write(stream, options);
  }
};
