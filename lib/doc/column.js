/**
 * Copyright (c) 2014 Guyon Roche
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
    return (this.width !== undefined) && (this.width !== 8);
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
      }
      else {
        this.style = {}
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
    return this._header && (this._header instanceof Array) ? this._header : [this._header];
  },
  get header() {
    return this._header;
  },
  set header(value) {
    if (value != undefined) {
      var self = this;
      this._header = value;
      this.headers.forEach(function (text, index) {
        self._worksheet.getCell(index + 1, self.number).value = text;
      });
    } else {
      this._header = [];
    }
    return value;
  },
  get key() {
    return this._key;
  },
  set key(value) {
    if (this._key && (this._worksheet._keys[this._key] === this)) {
      delete this._worksheet._keys[this._key];
    }
    this._key = value;
    if (value) {
      this._worksheet._keys[this._key] = this;
    }
    return value;
  },
  get hidden() {
    return !!this._hidden;
  },
  set hidden(value) {
    return this._hidden = value;
  },
  get outlineLevel() {
    return this._outlineLevel || 0;
  },
  set outlineLevel(value) {
    this._outlineLevel = value;
    return value;
  },
  get collapsed() {
    return !!(this._outlineLevel && (this._outlineLevel >= this._worksheet.properties.outlineLevelCol));
  },

  toString: function () {
    return JSON.stringify({
      key: this.key,
      width: this.width,
      headers: this.headers.length ? this.headers : undefined
    });
  },
  equivalentTo: function (other) {
    return (this.width == other.width) &&
            (this.hidden == other.hidden) &&
            (this.outlineLevel == other.outlineLevel) &&
            _.isEqual(this.style, other.style);
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

  eachCell: function (options, iteratee) {
    var colNumber = this.number;
    if (!iteratee) {
      iteratee = options;
      options = null;
    }
    if (options && options.includeEmpty) {
      this._worksheet.eachRow(options, function (row, rowNumber) {
        iteratee(row.getCell(colNumber), rowNumber);
      });
    } else {
      this._worksheet.eachRow(function (row, rowNumber) {
        var cell = row.findCell(colNumber);
        if (cell) {
          iteratee(cell, rowNumber);
        }
      });
    }
  },

  // =========================================================================
  // styles
  _applyStyle: function (name, value) {
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
    return this._applyStyle('numFmt', value);
  },
  get font() {
    return this.style.font;
  },
  set font(value) {
    return this._applyStyle('font', value);
  },
  get alignment() {
    return this.style.alignment;
  },
  set alignment(value) {
    return this._applyStyle('alignment', value);
  },
  get border() {
    return this.style.border;
  },
  set border(value) {
    return this._applyStyle('border', value);
  },
  get fill() {
    return this.style.fill;
  },
  set fill(value) {
    return this._applyStyle('fill', value);
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
      } else {
        if (!col || !column.equivalentTo(col)) {
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