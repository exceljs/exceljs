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
var CellMatrix = require("../utils/cell-matrix");

var DefinedNames = module.exports = function() {
  this._names = {};
  this._matrix = null;
};

DefinedNames.prototype = {
  // add a named cell. fullAddress in the form SheetName!$col$row
  add: function(name, fullAddress) {
    this._matrix = null;
    var entry = this._names[name];
    if (!entry) {
      entry = this._names[name] = [];
    }
    entry.push(fullAddress);
  },
  // add a named cell range. fullRange in the form SheetName!$col1$row1:$col2$row2
  addRange: function(name, fullRange) {
    var range = colCache.decodeEx(fullRange);
    for (var col = range.left; col <= range.right; col++) {
      for (var row = range.top; row <= range.bottom; row++) {
        this.add(name, range.sheetName + '!$' + colCache.n2l(col) + '$' + row);
      }
    }
  },

  getMatrix: function() {
    // build sparse array of names
    if (!this._matrix) {
      var matrix = this._matrix = new CellMatrix();
      _.each(this._names, function (addresses, name) {
        _.each(addresses, function (address) {
          var template = {address: address, names: []};
          matrix.getCell(address, template).names.push(name);
        });
      });
    }
    return this._matrix;
  },

  // get all the names of a cell
  getNames: function(fullAddress) {
    var matrix = this.getMatrix();
    var cell = matrix.findCell(fullAddress);
    return cell && cell.names;
  },

  getNameRanges: function(name) {
    // reduce named cells to ranges
  },

  get xml() {

  },
  set xml(value) {

  }
};

//<definedNames>
//  <definedName name="Corners">Sheet1!$H$1,Sheet1!$H$3,Sheet1!$J$3,Sheet1!$J$1</definedName>
//  <definedName name="CrossSheet">Sheet2!$A$1,Sheet1!$A$4,Sheet2!$C$1</definedName>
//  <definedName name="Ducks">Sheet1!$A$1:$A$3</definedName>
//  <definedName name="Intersection">Sheet1!$H$1,Sheet1!$L$1</definedName>
//</definedNames>
