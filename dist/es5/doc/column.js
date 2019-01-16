/**
 * Copyright (c) 2014-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../utils/under-dash');

var Enums = require('./enums');
var colCache = require('../utils/col-cache');
// Column defines the column properties for 1 column.
// This includes header rows, widths, key, (style), etc.
// Worksheet will condense the columns as appropriate during serialization
var Column = module.exports = function (worksheet, number, defn) {
  this._worksheet = worksheet;
  this._number = number;
  if (defn !== false) {
    // sometimes defn will follow
    this.defn = defn;
  }
};

Column.prototype = {
  get number() {
    return this._number;
  },
  get worksheet() {
    return this._worksheet;
  },
  get letter() {
    return colCache.n2l(this._number);
  },
  get isCustomWidth() {
    return this.width !== undefined && this.width !== 8;
  },
  get defn() {
    return {
      header: this._header,
      key: this.key,
      width: this.width,
      style: this.style,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel
    };
  },
  set defn(value) {
    if (value) {
      this.key = value.key;
      this.width = value.width;
      this.outlineLevel = value.outlineLevel;
      if (value.style) {
        this.style = value.style;
      } else {
        this.style = {};
      }

      // headers must be set after style
      this.header = value.header;
      this._hidden = !!value.hidden;
    } else {
      delete this._header;
      delete this.key;
      delete this.width;
      this.style = {};
      this.outlineLevel = 0;
    }
  },
  get headers() {
    return this._header && this._header instanceof Array ? this._header : [this._header];
  },
  get header() {
    return this._header;
  },
  set header(value) {
    var _this = this;

    if (value !== undefined) {
      this._header = value;
      this.headers.forEach(function (text, index) {
        _this._worksheet.getCell(index + 1, _this.number).value = text;
      });
    } else {
      this._header = undefined;
    }
  },
  get key() {
    return this._key;
  },
  set key(value) {
    var column = this._key && this._worksheet.getColumnKey(this._key);
    if (column === this) {
      this._worksheet.deleteColumnKey(this._key);
    }

    this._key = value;
    if (value) {
      this._worksheet.setColumnKey(this._key, this);
    }
  },
  get hidden() {
    return !!this._hidden;
  },
  set hidden(value) {
    this._hidden = value;
  },
  get outlineLevel() {
    return this._outlineLevel || 0;
  },
  set outlineLevel(value) {
    this._outlineLevel = value;
  },
  get collapsed() {
    return !!(this._outlineLevel && this._outlineLevel >= this._worksheet.properties.outlineLevelCol);
  },

  toString: function toString() {
    return JSON.stringify({
      key: this.key,
      width: this.width,
      headers: this.headers.length ? this.headers : undefined
    });
  },
  equivalentTo: function equivalentTo(other) {
    return this.width === other.width && this.hidden === other.hidden && this.outlineLevel === other.outlineLevel && _.isEqual(this.style, other.style);
  },
  get isDefault() {
    if (this.isCustomWidth) {
      return false;
    }
    if (this.hidden) {
      return false;
    }
    if (this.outlineLevel) {
      return false;
    }
    var s = this.style;
    if (s && (s.font || s.numFmt || s.alignment || s.border || s.fill)) {
      return false;
    }
    return true;
  },
  get headerCount() {
    return this.headers.length;
  },

  eachCell: function eachCell(options, iteratee) {
    var colNumber = this.number;
    if (!iteratee) {
      iteratee = options;
      options = null;
    }
    this._worksheet.eachRow(options, function (row, rowNumber) {
      iteratee(row.getCell(colNumber), rowNumber);
    });
  },

  get values() {
    var v = [];
    this.eachCell(function (cell, rowNumber) {
      if (cell && cell.type !== Enums.ValueType.Null) {
        v[rowNumber] = cell.value;
      }
    });
    return v;
  },

  set values(v) {
    var _this2 = this;

    if (!v) {
      return;
    }
    var colNumber = this.number;
    var offset = 0;
    if (v.hasOwnProperty('0')) {
      // assume contiguous array, start at row 1
      offset = 1;
    }
    v.forEach(function (value, index) {
      _this2._worksheet.getCell(index + offset, colNumber).value = value;
    });
  },

  // =========================================================================
  // styles
  _applyStyle: function _applyStyle(name, value) {
    this.style[name] = value;
    this.eachCell(function (cell) {
      cell[name] = value;
    });
    return value;
  },

  get numFmt() {
    return this.style.numFmt;
  },
  set numFmt(value) {
    this._applyStyle('numFmt', value);
  },
  get font() {
    return this.style.font;
  },
  set font(value) {
    this._applyStyle('font', value);
  },
  get alignment() {
    return this.style.alignment;
  },
  set alignment(value) {
    this._applyStyle('alignment', value);
  },
  get border() {
    return this.style.border;
  },
  set border(value) {
    this._applyStyle('border', value);
  },
  get fill() {
    return this.style.fill;
  },
  set fill(value) {
    this._applyStyle('fill', value);
  }
};

// =============================================================================
// static functions

Column.toModel = function (columns) {
  // Convert array of Column into compressed list cols
  var cols = [];
  var col = null;
  if (columns) {
    columns.forEach(function (column, index) {
      if (column.isDefault) {
        if (col) {
          col = null;
        }
      } else if (!col || !column.equivalentTo(col)) {
        col = {
          min: index + 1,
          max: index + 1,
          width: column.width,
          style: column.style,
          isCustomWidth: column.isCustomWidth,
          hidden: column.hidden,
          outlineLevel: column.outlineLevel,
          collapsed: column.collapsed
        };
        cols.push(col);
      } else {
        col.max = index + 1;
      }
    });
  }
  return cols.length ? cols : undefined;
};

Column.fromModel = function (worksheet, cols) {
  cols = cols || [];
  var columns = [];
  var count = 1;
  var index = 0;
  while (index < cols.length) {
    var col = cols[index++];
    while (count < col.min) {
      columns.push(new Column(worksheet, count++));
    }
    while (count <= col.max) {
      columns.push(new Column(worksheet, count++, col));
    }
  }
  return columns.length ? columns : null;
};
//# sourceMappingURL=column.js.map
