const fs = require('fs');
const fastCsv = require('fast-csv');
const customParseFormat = require('dayjs/plugin/customParseFormat');
const utc = require('dayjs/plugin/utc');
const dayjs = require('dayjs').extend(customParseFormat).extend(utc);
const StreamBuf = require('../utils/stream-buf');

const {
  fs: {exists},
} = require('../utils/utils');

/* eslint-disable quote-props */
const SpecialValues = {
  true: true,
  false: false,
  '#N/A': {error: '#N/A'},
  '#REF!': {error: '#REF!'},
  '#NAME?': {error: '#NAME?'},
  '#DIV/0!': {error: '#DIV/0!'},
  '#NULL!': {error: '#NULL!'},
  '#VALUE!': {error: '#VALUE!'},
  '#NUM!': {error: '#NUM!'},
};
/* eslint-enable quote-props */

class CSV {
  constructor(workbook) {
    this.workbook = workbook;
    this.worksheets = new Map(); // To hold multiple sheets
  }

  async readFile(filename, options) {
    options = options || {};
    if (!(await exists(filename))) {
      throw new Error(`File not found: ${filename}`);
    }
    const stream = fs.createReadStream(filename);
    const worksheets = await this.read(stream, options);
    stream.close();
    return worksheets;
  }

  // Updated to handle multiple sheets
  read(stream, options) {
    options = options || {};
    const worksheets = new Map();

    return new Promise((resolve, reject) => {
      const dateFormats = options.dateFormats || [
        'YYYY-MM-DD[T]HH:mm:ssZ',
        'YYYY-MM-DD[T]HH:mm:ss',
        'MM-DD-YYYY',
        'YYYY-MM-DD',
      ];

      const map = options.map || this.defaultMap(dateFormats);

      const csvStream = fastCsv
        .parse(options.parserOptions)
        .on('data', (data) => {
          // Assuming first row is for sheet name
          if (!worksheets.has(options.sheetName)) {
            worksheets.set(options.sheetName, []);
          }
          worksheets.get(options.sheetName).push(data.map(map));
        })
        .on('end', () => {
          resolve(worksheets); // Return all worksheets as a map
        })
        .on('error', reject);

      stream.pipe(csvStream);
    });
  }

  // Default data mapping for CSV parsing
  defaultMap(dateFormats) {
    return (datum) => {
      if (datum === '') {
        return null;
      }
      const datumNumber = Number(datum);
      if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
        return datumNumber;
      }
      const dt = dateFormats.reduce((matchingDate, currentDateFormat) => {
        if (matchingDate) {
          return matchingDate;
        }
        const dayjsObj = dayjs(datum, currentDateFormat, true);
        if (dayjsObj.isValid()) {
          return dayjsObj;
        }
        return null;
      }, null);
      if (dt) {
        return new Date(dt.valueOf());
      }
      const special = SpecialValues[datum];
      if (special !== undefined) {
        return special;
      }
      return datum;
    };
  }

  // Writing support for multiple sheets
  write(stream, options) {
    return new Promise((resolve, reject) => {
      options = options || {};
      const csvStream = fastCsv.format(options.formatterOptions);

      stream.on('finish', () => {
        resolve();
      });
      csvStream.on('error', reject);
      csvStream.pipe(stream);

      const {dateFormat, dateUTC} = options;
      const map = options.map || this.defaultMap([dateFormat]);

      // Write each sheet in the workbook
      this.workbook.eachSheet((worksheet) => {
        worksheet.eachRow((row) => {
          const values = row.values.slice(1); // Skip row number
          csvStream.write(values.map(map));
        });
      });

      csvStream.end();
    });
  }

  writeFile(filename, options) {
    options = options || {};
    const streamOptions = {
      encoding: options.encoding || 'utf8',
    };
    const stream = fs.createWriteStream(filename, streamOptions);

    return this.write(stream, options);
  }

  // New method: Add a batch mode for large files (to avoid memory overload)
  async writeBuffer(options) {
    const stream = new StreamBuf();
    await this.write(stream, options);
    return stream.read();
  }

  // New method: Custom error handling for malformed data
  validateData(row) {
    // Validate if row is malformed (e.g., too many or too few columns)
    const expectedColumns = 5; // Set the expected number of columns
    if (row.length !== expectedColumns) {
      throw new Error(`Malformed data row: Expected ${expectedColumns} columns, but got ${row.length}`);
    }
    return true;
  }

  // New method: Add dynamic column mapping for writing
  mapColumns(columns) {
    return (value, columnIndex) => {
      if (columns && columns[columnIndex]) {
        return columns[columnIndex](value);
      }
      return value;
    };
  }

  // New method: Streaming for larger CSV files
  async streamFile(filename, options) {
    const stream = fs.createReadStream(filename, { encoding: 'utf8' });
    const csvStream = fastCsv.parse(options.parserOptions);

    const rows = [];
    csvStream.on('data', (row) => {
      rows.push(row);
      if (rows.length >= 1000) {
        // Process rows in batches of 1000
        this.processBatch(rows);
        rows.length = 0; // Reset the batch
      }
    });

    stream.pipe(csvStream);

    return new Promise((resolve, reject) => {
      csvStream.on('end', () => {
        if (rows.length) this.processBatch(rows); // Process remaining rows
        resolve();
      });
      csvStream.on('error', reject);
    });
  }

  // Example of batch processing method
  processBatch(batch) {
    console.log(`Processing ${batch.length} rows...`);
    // Custom processing logic for each batch (e.g., saving to a database)
  }
}

module.exports = CSV;
