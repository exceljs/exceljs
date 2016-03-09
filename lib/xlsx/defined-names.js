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
var Dimensions = require('../utils/dimensions');

var DefinedNames = module.exports = function() {
  this.matrix = new CellMatrix({names: {}});
};

DefinedNames.prototype = {
  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add: function(locStr, name) {
    var location = colCache.decodeEx(locStr);
    this.addEx(location, name);
  },
  addEx: function(location, name) {
    if (location.top) {
      for (var col = location.left; col <= location.right; col++) {
        for (var row = location.top; row <= location.bottom; row++) {
          var address = {
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row: row,
            col: col
          };

          this.matrix.getCellEx(address).names[name] = true;
        }
      }
    } else {
      this.matrix.getCellEx(location).names[name] = true;
    }
  },

  forEach: function(callback) {
    this.matrix.forEach(function(cell) {
      callback(cell, _.map(cell.names, function(value, name) { return name; }));
    });
  },

  // get all the names of a cell - address is decoded struct
  getNames: function(address) {
    var cell = this.matrix.findCellEx(address);
    if (cell && cell.names) {
      return _.map(cell.names, function(value, name) { return name; });
    } else {
      return null;
    }
  },

  _explore: function(cell, name) {
    cell.names[name] = false;
    var sheetName = cell.sheetName;

    var dimensions = new Dimensions(cell.row, cell.col, cell.row, cell.col);
    var x, y;
    var matrix = this.matrix;

    // grow vertical - only one col to worry about
    function vGrow(y, edge) {
      var c = matrix.findCellAt(sheetName, y, cell.col);
      if (!c || !c.names[name]) { return false; }
      dimensions[edge] = y;
      c.names[name] = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, 'top'); y--) {}
    for (y = cell.row + 1; vGrow(y, 'bottom'); y++) {}

    // grow horizontal - ensure all rows can grow
    function hGrow(x, edge) {
      var c, cells = [];
      for (y = dimensions.top; y <= dimensions.bottom; y++) {
        c = matrix.findCellAt(sheetName, y, x);
        if (c && c.names[name]) {
          cells.push(c);
        } else {
          return false;
        }
      }
      dimensions[edge] = x;
      for (var i = 0; i < cells.length; i++) {
        cells[i].names[name] = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, 'left'); x--) {}
    for (x = cell.col + 1; hGrow(x, 'right'); x++) {}

    return dimensions;
  },

  getNameRanges: function() {
    var self = this;

    var names = {};

    // iterate cells - grow ranges using mark & sweep - prioritise vertical
    this.matrix.forEach(function(cell) {
      _.each(cell.names, function(value, name) {
        cell.names[name] = true;
      });
    });
    this.matrix.forEach(function(cell) {
      _.each(cell.names, function(value, name) {
        if (cell.names[name]) {
          var dimensions = self._explore(cell, name);
          if (dimensions) {
            var ranges = names[name] || (names[name] = []);
            var range = dimensions.count > 1 ? dimensions.$range : dimensions.$t$l;
            ranges.push(cell.sheetName + '!' + range);
          }
        }
      });
    });

    return names;
  },

  get xml() {
    var self = this;
    var xml = [];
    var names = this.getNameRanges();
    _.each(names, function(ranges, name) {
      if (ranges.length) {
        xml.push('<definedName name="' + name + '">');
        xml.push(ranges.join(','));
        xml.push('</definedName>');
      }
    });
    return xml.length ?
      '<definedNames>' + xml.join('') + '</definedNames>' :
      '';
  },
  parse: function(node) {
    // new definedName
    this._parseName = node.attributes.name;
    this._parseText = [];
  },
  parseText: function(text) {
    // some text of definedName
    this._parseText.push(text);
  },
  parseClose: function() {
    var self = this;
    // definedName finished. Pull out ranges
    var addresses = this._parseText.join('').split(',');
    _.each(addresses, function(address) {
      self.add(address, self._parseName);
    });
  }
};

//<definedNames>
//  <definedName name="Corners">Sheet1!$H$1,Sheet1!$H$3,Sheet1!$J$3,Sheet1!$J$1</definedName>
//  <definedName name="CrossSheet">Sheet2!$A$1,Sheet1!$A$4,Sheet2!$C$1</definedName>
//  <definedName name="Ducks">Sheet1!$A$1:$A$3</definedName>
//  <definedName name="Intersection">Sheet1!$H$1,Sheet1!$L$1</definedName>
//</definedNames>
