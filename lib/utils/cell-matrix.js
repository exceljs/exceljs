/**
 * Copyright (c) 2016 Guyon Roche
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
"use strict";
var _ = require('underscore');
var utils = require("../utils/utils");
var colCache = require("../utils/col-cache");

var Matrix = function() {
  this.sheets = {};
};

Matrix.prototype = {
  getCell: function(address, template) {
    return this.findCell(address, true, template);
  },
  findCell: function(address, create, template) {
    address = colCache.decodeEx(address);
    var sheet = this.findSheet(address.sheetName, true);
    var row = this.findSheetRow(sheet, address.row, true);
    return this.findRowCell(row, address.col, true, template);
  },

  findSheet: function(name, create) {
    if (this.sheets[name]) {
      return this.sheets[name];
    }
    if (create) {
      return this.sheets[name] = [];
    }
  },
  findSheetRow: function(sheet, row, create) {
    if (sheet && sheet[row]) {
      return sheet[row];
    }
    if (create) {
      return sheet[row] = [];
    }
  },
  findRowCell: function(row, col, create, template) {
    if (row && row[col]) {
      return row[col];
    }
    if (create) {
      return row[col] = template;
    }
  }
};
