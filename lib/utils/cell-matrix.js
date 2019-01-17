/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const _ = require('./under-dash');
const colCache = require('./col-cache');

const CellMatrix = function(template) {
  this.template = template;
  this.sheets = {};
};

CellMatrix.prototype = {
  addCell(addressStr) {
    this.addCellEx(colCache.decodeEx(addressStr));
  },
  getCell(addressStr) {
    return this.findCellEx(colCache.decodeEx(addressStr), true);
  },
  findCell(addressStr) {
    return this.findCellEx(colCache.decodeEx(addressStr), false);
  },

  findCellAt(sheetName, rowNumber, colNumber) {
    const sheet = this.sheets[sheetName];
    const row = sheet && sheet[rowNumber];
    return row && row[colNumber];
  },
  addCellEx(address) {
    if (address.top) {
      for (let row = address.top; row <= address.bottom; row++) {
        for (let col = address.left; col <= address.right; col++) {
          this.getCellAt(address.sheetName, row, col);
        }
      }
    } else {
      this.findCellEx(address, true);
    }
  },
  getCellEx(address) {
    return this.findCellEx(address, true);
  },
  findCellEx(address, create) {
    const sheet = this.findSheet(address, create);
    const row = this.findSheetRow(sheet, address, create);
    return this.findRowCell(row, address, create);
  },
  getCellAt(sheetName, rowNumber, colNumber) {
    const sheet = this.sheets[sheetName] || (this.sheets[sheetName] = []);
    const row = sheet[rowNumber] || (sheet[rowNumber] = []);
    return (
      row[colNumber] ||
      (row[colNumber] = {
        sheetName,
        address: colCache.n2l(colNumber) + rowNumber,
        row: rowNumber,
        col: colNumber,
      })
    );
  },

  removeCellEx(address) {
    const sheet = this.findSheet(address);
    if (!sheet) {
      return;
    }
    const row = this.findSheetRow(sheet, address);
    if (!row) {
      return;
    }
    delete row[address.col];
  },

  forEach(callback) {
    _.each(this.sheets, sheet => {
      if (sheet) {
        sheet.forEach(row => {
          if (row) {
            row.forEach(cell => {
              if (cell) {
                callback(cell);
              }
            });
          }
        });
      }
    });
  },
  map(callback) {
    const results = [];
    this.forEach(cell => {
      results.push(callback(cell));
    });
    return results;
  },

  findSheet(address, create) {
    const name = address.sheetName;
    if (this.sheets[name]) {
      return this.sheets[name];
    }
    if (create) {
      return (this.sheets[name] = []);
    }
    return undefined;
  },
  findSheetRow(sheet, address, create) {
    const row = address.row;
    if (sheet && sheet[row]) {
      return sheet[row];
    }
    if (create) {
      return (sheet[row] = []);
    }
    return undefined;
  },
  findRowCell(row, address, create) {
    const col = address.col;
    if (row && row[col]) {
      return row[col];
    }
    if (create) {
      return (row[col] = this.template ? Object.assign(address, JSON.parse(JSON.stringify(this.template))) : address);
    }
    return undefined;
  },
};

module.exports = CellMatrix;
