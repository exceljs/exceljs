/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software"), to deal
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

var _ = require('./under-dash');

var utils = require('./utils');
var colCache = require('./col-cache');

var CellMatrix = function(template) {
  this.template = template;
  this.sheets = {};
};

CellMatrix.prototype = {
  addCell: function(addressStr) {
    this.addCellEx(colCache.decodeEx(addressStr));
  },
  getCell: function(addressStr) {
    return this.findCellEx(colCache.decodeEx(addressStr), true);
  },
  findCell: function(addressStr) {
    return this.findCellEx(colCache.decodeEx(addressStr), false);
  },

  findCellAt: function(sheetName, rowNumber, colNumber) {
    var sheet = this.sheets[sheetName];
    var row = sheet && sheet[rowNumber];
    return row && row[colNumber];
  },
  addCellEx: function(address) {
    if (address.top) {
      for (var row = address.top; row <= address.bottom; row++) {
        for (var col = address.left; col <= address.right; col++) {
          this.getCellAt(address.sheetName, row, col);
        }
      }
    } else {
      this.findCellEx(address, true);
    }
  },
  getCellEx: function(address) {
    return this.findCellEx(address, true);
  },
  findCellEx: function(address, create) {
    var sheet = this.findSheet(address, create);
    var row = this.findSheetRow(sheet, address, create);
    return this.findRowCell(row, address, create);
  },
  getCellAt: function(sheetName, rowNumber, colNumber) {
    var sheet = this.sheets[sheetName] || (this.sheets[sheetName] = []);
    var row = sheet[rowNumber] || (sheet[rowNumber] = []);
    return row[colNumber] || (row[colNumber] = {
      sheetName: sheetName,
      address: colCache.n2l(colNumber) + rowNumber,
      row: rowNumber,
      col: colNumber
    });
  },

  removeCellEx: function(address) {
    var sheet = this.findSheet(address);
    if (!sheet) { return; }
    var row = this.findSheetRow(sheet, address);
    if (!row) { return; }
    delete row[address.col];
  },

  forEach: function(callback) {
    _.each(this.sheets, function(sheet) {
      if (sheet) sheet.forEach(function(row) {
        if (row) row.forEach(function(cell) {
          if (cell) callback(cell);
        });
      });
    });
  },
  map: function(callback) {
    var results = [];
    this.forEach(function(cell) { results.push(callback(cell)); });
    return results;
  },

  findSheet: function(address, create) {
    var name = address.sheetName;
    if (this.sheets[name]) {
      return this.sheets[name];
    }
    if (create) {
      return (this.sheets[name] = []);
    }
  },
  findSheetRow: function(sheet, address, create) {
    var row = address.row;
    if (sheet && sheet[row]) {
      return sheet[row];
    }
    if (create) {
      return (sheet[row] = []);
    }
  },
  findRowCell: function(row, address, create) {
    var col = address.col;
    if (row && row[col]) {
      return row[col];
    }
    if (create) {
      return (row[col] = this.template ? Object.assign(address, JSON.parse(JSON.stringify(this.template))) : address);
    }
  }
};

module.exports = CellMatrix;