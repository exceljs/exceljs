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

var _ = require('underscore');
var utils = require('../utils/utils');
var colCache = require('../utils/col-cache');

var CellMatrix = function(template) {
  this.template = template;
  this.sheets = {};
};

CellMatrix.prototype = {
  getCell: function(addressStr) {
    return this.findCellEx(colCache.decodeEx(addressStr), true);
  },
  findCell: function(addressStr) {
    return this.findCellEx(colCache.decodeEx(addressStr), false);
  },

  findCellAt: function(name, rowNumber, colNumber) {
    var sheet = this.sheets[name];
    var row = sheet && sheet[rowNumber];
    return row && row[colNumber];
  },
  getCellEx: function(address) {
    var cell = this.findCellEx(address, true);
    //console.log('getCellEx', JSON.stringify(address), JSON.stringify(cell));
    return cell;
  },
  findCellEx: function(address, create) {
    var sheet = this.findSheet(address, create);
    var row = this.findSheetRow(sheet, address, create);
    return this.findRowCell(row, address, create);
  },

  forEach: function(callback) {
    _.each(this.sheets, function(sheet) {
      _.each(sheet, function(row) {
        _.each(row, function(cell) {
          callback(cell);
        });
      });
    });
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
      return (row[col] = this.template ? _.extend(address, JSON.parse(JSON.stringify(this.template))) : address);
    }
  }
};

module.exports = CellMatrix;