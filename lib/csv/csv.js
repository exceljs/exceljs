'use strict';

const fs = require('fs');
const csv = require('fast-csv');
const moment = require('moment');
const PromiseLib = require('../utils/promise');
const StreamBuf = require('../utils/stream-buf');

const utils = require('../utils/utils');

const CSV = (module.exports = function(workbook) {
  this.workbook = workbook;
  this.worksheet = null;
});

/* eslint-disable quote-props */
const SpecialValues = {
  true: true,
  false: false,
  '#N/A': { error: '#N/A' },
  '#REF!': { error: '#REF!' },
  '#NAME?': { error: '#NAME?' },
  '#DIV/0!': { error: '#DIV/0!' },
  '#NULL!': { error: '#NULL!' },
  '#VALUE!': { error: '#VALUE!' },
  '#NUM!': { error: '#NUM!' },
};
/* eslint-ensable quote-props */

CSV.prototype = {
  readFile(filename, options) {
    const self = this;
    options = options || {};
    let stream;
    return utils.fs
      .exists(filename)
      .then(exists => {
        if (!exists) {
          throw new Error(`File not found: ${filename}`);
        }
        stream = fs.createReadStream(filename);
        return self.read(stream, options);
      })
      .then(worksheet => {
        stream.close();
        return worksheet;
      });
  },
  read(stream, options) {
    options = options || {};
    return new PromiseLib.Promise((resolve, reject) => {
      const csvStream = this.createInputStream(options)
        .on('worksheet', resolve)
        .on('error', reject);

      stream.pipe(csvStream);
    });
  },
  createInputStream(options) {
    options = options || {};
    const worksheet = this.workbook.addWorksheet(options.sheetName);

    const dateFormats = options.dateFormats || [moment.ISO_8601, 'MM-DD-YYYY', 'YYYY-MM-DD'];
    const map =
      options.map ||
      function(datum) {
        if (datum === '') {
          return null;
        }
        const datumNumber = Number(datum);
        if (!Number.isNaN(datumNumber)) {
          return datumNumber;
        }
        const dt = moment(datum, dateFormats, true);
        if (dt.isValid()) {
          return new Date(dt.valueOf());
        }
        const special = SpecialValues[datum];
        if (special !== undefined) {
          return special;
        }
        return datum;
      };

    const csvStream = csv(options)
      .on('data', data => {
        worksheet.addRow(data.map(map));
      })
      .on('end', () => {
        csvStream.emit('worksheet', worksheet);
      });
    return csvStream;
  },

  write(stream, options) {
    return new PromiseLib.Promise((resolve, reject) => {
      options = options || {};
      // const encoding = options.encoding || 'utf8';
      // const separator = options.separator || ',';
      // const quoteChar = options.quoteChar || '\'';

      const worksheet = this.workbook.getWorksheet(options.sheetName || options.sheetId);

      const csvStream = csv.createWriteStream(options);
      stream.on('finish', () => {
        resolve();
      });
      csvStream.on('error', reject);
      csvStream.pipe(stream);

      const { dateFormat, dateUTC } = options;
      const map =
        options.map ||
        (value => {
          if (value) {
            if (value.text || value.hyperlink) {
              return value.hyperlink || value.text || '';
            }
            if (value.formula || value.result) {
              return value.result || '';
            }
            if (value instanceof Date) {
              if (dateFormat) {
                return dateUTC ? moment.utc(value).format(dateFormat) : moment(value).format(dateFormat);
              }
              return dateUTC ? moment.utc(value).format() : moment(value).format();
            }
            if (value.error) {
              return value.error;
            }
            if (typeof value === 'object') {
              return JSON.stringify(value);
            }
          }
          return value;
        });

      const includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
      let lastRow = 1;
      if (worksheet) {
        worksheet.eachRow((row, rowNumber) => {
          if (includeEmptyRows) {
            while (lastRow++ < rowNumber - 1) {
              csvStream.write([]);
            }
          }
          const { values } = row;
          values.shift();
          csvStream.write(values.map(map));
          lastRow = rowNumber;
        });
      }
      csvStream.end();
    });
  },
  writeFile(filename, options) {
    options = options || {};

    const streamOptions = {
      encoding: options.encoding || 'utf8',
    };
    const stream = fs.createWriteStream(filename, streamOptions);

    return this.write(stream, options);
  },
  writeBuffer(options) {
    const self = this;
    const stream = new StreamBuf();
    return self.write(stream, options).then(() => stream.read());
  },
};
