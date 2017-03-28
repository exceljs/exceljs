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

var _ = require('../utils/under-dash');
var colCache = require('../utils/col-cache');
var CellMatrix = require('../utils/cell-matrix');
var Range = require('./range');
var rangeRegexp = /\$(\w+)\$(\d+)(:\$(\w+)\$(\d+))?/;

var DefinedNames = module.exports = function() {
  this.matrixMap = {};
};

DefinedNames.prototype = {
  getMatrix: function(name) {
    return this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
  },

  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add: function(locStr, name) {
    var location = colCache.decodeEx(locStr);
    this.addEx(location, name);
  },
  addEx: function(location, name) {
    var matrix = this.getMatrix(name);
    if (location.top) {
      for (var col = location.left; col <= location.right; col++) {
        for (var row = location.top; row <= location.bottom; row++) {
          var address = {
            sheetName: location.sheetName,
            address: colCache.n2l(col) + row,
            row: row,
            col: col
          };

          matrix.addCellEx(address);
        }
      }
    } else {
      matrix.addCellEx(location);
    }
  },

  remove: function(locStr, name) {
    var location = colCache.decodeEx(locStr);
    this.removeEx(location, name);
  },
  removeEx: function(location, name) {
    var matrix = this.getMatrix(name);
    matrix.removeCellEx(location);
  },
  removeAllNames: function(location) {
    _.each(this.matrixMap, function(matrix) {
      matrix.removeCellEx(location);
    });
  },

  forEach: function(callback) {
    _.each(this.matrixMap, function(matrix, name) {
      matrix.forEach(function(cell) {
        callback(name, cell);
      });
    });
  },

  // get all the names of a cell
  getNames: function(addressStr) {
    return this.getNamesEx(colCache.decodeEx(addressStr));
  },
  getNamesEx: function(address) {
    return _.map(this.matrixMap, function(matrix, name) {
      if (matrix.findCellEx(address)) {
        return name;
      }
    }).filter(function(name) { return name; });
  },

  _explore: function(matrix, cell) {
    cell.mark = false;
    var sheetName = cell.sheetName;

    var range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
    var x, y;

    // grow vertical - only one col to worry about
    function vGrow(y, edge) {
      var c = matrix.findCellAt(sheetName, y, cell.col);
      if (!c || !c.mark) { return false; }
      range[edge] = y;
      c.mark = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, 'top'); y--) {}
    for (y = cell.row + 1; vGrow(y, 'bottom'); y++) {}

    // grow horizontal - ensure all rows can grow
    function hGrow(x, edge) {
      var c, cells = [];
      for (y = range.top; y <= range.bottom; y++) {
        c = matrix.findCellAt(sheetName, y, x);
        if (c && c.mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      range[edge] = x;
      for (var i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, 'left'); x--) {}
    for (x = cell.col + 1; hGrow(x, 'right'); x++) {}

    return range;
  },

  getRanges: function(name, matrix) {
    var self = this;
    matrix = matrix || this.matrixMap[name];

    if (!matrix) {
      return { name: name, ranges: [] };
    }

    // mark and sweep!
    matrix.forEach(function(cell) { cell.mark = true; });
    var ranges = matrix.map(function(cell) {
        return cell.mark && self._explore(matrix, cell);
      })
      .filter(function(range) { return range; })
      .map(function(range) { return range.$shortRange; });

    return {
      name: name, ranges: ranges
    };
  },

  get model() {
    var self = this;
    // To get names per cell - just iterate over all names finding cells if they exist
    return _.map(this.matrixMap, function(matrix, name) {
        return self.getRanges(name, matrix);
      })
      .filter(function(definedName) {
        return definedName.ranges.length;
      });
  },
  set model(value) {
    // value is [ { name, ranges }, ... ]
    var matrixMap = this.matrixMap = {};
    value.forEach(function(definedName) {
      var matrix = matrixMap[definedName.name] = new CellMatrix();
      definedName.ranges.forEach(function(rangeStr) {
        if(rangeRegexp.test(rangeStr.split('!').pop() || '')) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
};

//<definedNames>
//  <definedName name="Corners">Sheet1!$H$1,Sheet1!$H$3,Sheet1!$J$3,Sheet1!$J$1</definedName>
//  <definedName name="CrossSheet">Sheet2!$A$1,Sheet1!$A$4,Sheet2!$C$1</definedName>
//  <definedName name="Ducks">Sheet1!$A$1:$A$3</definedName>
//  <definedName name="Intersection">Sheet1!$H$1,Sheet1!$L$1</definedName>
//</definedNames>
