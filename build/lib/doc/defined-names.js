/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../utils/under-dash');
var colCache = require('../utils/col-cache');
var CellMatrix = require('../utils/cell-matrix');
var Range = require('./range');

var rangeRegexp = /[$](\w+)[$](\d+)(:[$](\w+)[$](\d+))?/;

var DefinedNames = module.exports = function () {
  this.matrixMap = {};
};

DefinedNames.prototype = {
  getMatrix: function getMatrix(name) {
    return this.matrixMap[name] || (this.matrixMap[name] = new CellMatrix());
  },

  // add a name to a cell. locStr in the form SheetName!$col$row or SheetName!$c1$r1:$c2:$r2
  add: function add(locStr, name) {
    var location = colCache.decodeEx(locStr);
    this.addEx(location, name);
  },
  addEx: function addEx(location, name) {
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

  remove: function remove(locStr, name) {
    var location = colCache.decodeEx(locStr);
    this.removeEx(location, name);
  },
  removeEx: function removeEx(location, name) {
    var matrix = this.getMatrix(name);
    matrix.removeCellEx(location);
  },
  removeAllNames: function removeAllNames(location) {
    _.each(this.matrixMap, function (matrix) {
      matrix.removeCellEx(location);
    });
  },

  forEach: function forEach(callback) {
    _.each(this.matrixMap, function (matrix, name) {
      matrix.forEach(function (cell) {
        callback(name, cell);
      });
    });
  },

  // get all the names of a cell
  getNames: function getNames(addressStr) {
    return this.getNamesEx(colCache.decodeEx(addressStr));
  },
  getNamesEx: function getNamesEx(address) {
    return _.map(this.matrixMap, function (matrix, name) {
      return matrix.findCellEx(address) && name;
    }).filter(Boolean);
  },

  _explore: function _explore(matrix, cell) {
    cell.mark = false;
    var sheetName = cell.sheetName;

    var range = new Range(cell.row, cell.col, cell.row, cell.col, sheetName);
    var x, y;

    // grow vertical - only one col to worry about
    function vGrow(yy, edge) {
      var c = matrix.findCellAt(sheetName, yy, cell.col);
      if (!c || !c.mark) {
        return false;
      }
      range[edge] = yy;
      c.mark = false;
      return true;
    }
    for (y = cell.row - 1; vGrow(y, 'top'); y--) {}
    for (y = cell.row + 1; vGrow(y, 'bottom'); y++) {}

    // grow horizontal - ensure all rows can grow
    function hGrow(xx, edge) {
      var c,
          cells = [];
      for (y = range.top; y <= range.bottom; y++) {
        c = matrix.findCellAt(sheetName, y, xx);
        if (c && c.mark) {
          cells.push(c);
        } else {
          return false;
        }
      }
      range[edge] = xx;
      for (var i = 0; i < cells.length; i++) {
        cells[i].mark = false;
      }
      return true;
    }
    for (x = cell.col - 1; hGrow(x, 'left'); x--) {}
    for (x = cell.col + 1; hGrow(x, 'right'); x++) {}

    return range;
  },

  getRanges: function getRanges(name, matrix) {
    var _this = this;

    matrix = matrix || this.matrixMap[name];

    if (!matrix) {
      return { name: name, ranges: [] };
    }

    // mark and sweep!
    matrix.forEach(function (cell) {
      cell.mark = true;
    });
    var ranges = matrix.map(function (cell) {
      return cell.mark && _this._explore(matrix, cell);
    }).filter(Boolean).map(function (range) {
      return range.$shortRange;
    });

    return {
      name: name, ranges: ranges
    };
  },

  get model() {
    var self = this;
    // To get names per cell - just iterate over all names finding cells if they exist
    return _.map(this.matrixMap, function (matrix, name) {
      return self.getRanges(name, matrix);
    }).filter(function (definedName) {
      return definedName.ranges.length;
    });
  },
  set model(value) {
    // value is [ { name, ranges }, ... ]
    var matrixMap = this.matrixMap = {};
    value.forEach(function (definedName) {
      var matrix = matrixMap[definedName.name] = new CellMatrix();
      definedName.ranges.forEach(function (rangeStr) {
        if (rangeRegexp.test(rangeStr.split('!').pop() || '')) {
          matrix.addCell(rangeStr);
        }
      });
    });
  }
};
//# sourceMappingURL=defined-names.js.map
