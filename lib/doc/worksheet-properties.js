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
'use strict';
var _ = require('underscore');
var utils = require('../utils/utils');
var colCache = require('../utils/col-cache');
var CellMatrix = require('../utils/cell-matrix');
var Range = require('./range');

var WorksheetFormatProperties = module.exports = function(worksheet, model) {
  model = model || {};
  this._worksheet = worksheet;
  this.model = {
    defaultRowHeight: model.defaultRowHeight || 15,
    dyDescent: model.dyDescent || 55,
    outlineLevelCol: model.outlineLevelCol || 0,
    outlineLevelRow: model.outlineLevelRow || 0
  };
};

WorksheetFormatProperties.prototype = {
  get defaultRowHeight() {
    return this.model.defaultRowHeight;
  },
  set defaultRowHeight(value) {
    return this.model.defaultRowHeight = value;
  },
  get dyDescent() {
    return this.model.dyDescent;
  },
  set dyDescent(value) {
    return this.model.dyDescent = value;
  },
  get outlineLevelCol() {
    return this.model.outlineLevelCol;
  },
  set outlineLevelCol(value) {
    console.log('set outlineLevelCol', value)
    this.model.outlineLevelCol = value;
    _.each(this._worksheet._columns, function(column) {
      column.collapsed = column.outlineLevel >= value;
    });
    return value;
  },
  get outlineLevelRow() {
    return this.model.outlineLevelRow;
  },
  set outlineLevelRow(value) {
    console.log('set outlineLevelRow', value)
    this.model.outlineLevelRow = value;
    _.each(this._worksheet._rows, function(row) {
      row.collapsed = row.outlineLevel >= value;
    });
    return value;
  }
};

