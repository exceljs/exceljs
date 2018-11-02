/**
 * Copyright (c) 2014-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var colCache = require('./../utils/col-cache');

// used by worksheet to calculate sheet dimensions
var Range = module.exports = function () {
  this.decode(arguments);
};

Range.prototype = {
  _set_tlbr: function _set_tlbr(t, l, b, r, s) {
    this.model = {
      top: Math.min(t, b),
      left: Math.min(l, r),
      bottom: Math.max(t, b),
      right: Math.max(l, r),
      sheetName: s
    };
  },
  _set_tl_br: function _set_tl_br(tl, br, s) {
    tl = colCache.decodeAddress(tl);
    br = colCache.decodeAddress(br);
    this._set_tlbr(tl.row, tl.col, br.row, br.col, s);
  },
  decode: function decode(argv) {
    switch (argv.length) {
      case 5:
        // [t,l,b,r,s]
        this._set_tlbr(argv[0], argv[1], argv[2], argv[3], argv[4]);
        break;
      case 4:
        // [t,l,b,r]
        this._set_tlbr(argv[0], argv[1], argv[2], argv[3]);
        break;

      case 3:
        // [tl,br,s]
        this._set_tl_br(argv[0], argv[1], argv[2]);
        break;
      case 2:
        // [tl,br]
        this._set_tl_br(argv[0], argv[1]);
        break;

      case 1:
        var value = argv[0];
        if (value instanceof Range) {
          // copy constructor
          this.model = {
            top: value.model.top,
            left: value.model.left,
            bottom: value.model.bottom,
            right: value.model.right,
            sheetName: value.sheetName
          };
        } else if (value instanceof Array) {
          // an arguments array
          this.decode(value);
        } else if (value.top && value.left && value.bottom && value.right) {
          // a model
          this.model = {
            top: value.top,
            left: value.left,
            bottom: value.bottom,
            right: value.right,
            sheetName: value.sheetName
          };
        } else {
          // [sheetName!]tl:br
          var tlbr = colCache.decodeEx(value);
          if (tlbr.top) {
            this.model = {
              top: tlbr.top,
              left: tlbr.left,
              bottom: tlbr.bottom,
              right: tlbr.right,
              sheetName: tlbr.sheetName
            };
          } else {
            this.model = {
              top: tlbr.row,
              left: tlbr.col,
              bottom: tlbr.row,
              right: tlbr.col,
              sheetName: tlbr.sheetName
            };
          }
        }
        break;

      case 0:
        this.model = {
          top: 0,
          left: 0,
          bottom: 0,
          right: 0
        };
        break;

      default:
        throw new Error('Invalid number of arguments to _getDimensions() - ' + argv.length);
    }
  },

  get top() {
    return this.model.top || 1;
  },
  set top(value) {
    this.model.top = value;
  },
  get left() {
    return this.model.left || 1;
  },
  set left(value) {
    this.model.left = value;
  },
  get bottom() {
    return this.model.bottom || 1;
  },
  set bottom(value) {
    this.model.bottom = value;
  },
  get right() {
    return this.model.right || 1;
  },
  set right(value) {
    this.model.right = value;
  },
  get sheetName() {
    return this.model.sheetName;
  },
  set sheetName(value) {
    this.model.sheetName = value;
  },

  get _serialisedSheetName() {
    var sheetName = this.model.sheetName;
    if (sheetName) {
      if (/^[a-zA-Z0-9]*$/.test(sheetName)) {
        return sheetName + '!';
      }
      return '\'' + sheetName + '\'!';
    }
    return '';
  },

  expand: function expand(top, left, bottom, right) {
    if (!this.model.top || top < this.top) this.top = top;
    if (!this.model.left || left < this.left) this.left = left;
    if (!this.model.bottom || bottom > this.bottom) this.bottom = bottom;
    if (!this.model.right || right > this.right) this.right = right;
  },
  expandRow: function expandRow(row) {
    if (row) {
      var dimensions = row.dimensions;
      if (dimensions) {
        this.expand(row.number, dimensions.min, row.number, dimensions.max);
      }
    }
  },
  expandToAddress: function expandToAddress(addressStr) {
    var address = colCache.decodeEx(addressStr);
    this.expand(address.row, address.col, address.row, address.col);
  },

  get tl() {
    return colCache.n2l(this.left) + this.top;
  },
  get $t$l() {
    return '$' + colCache.n2l(this.left) + '$' + this.top;
  },
  get br() {
    return colCache.n2l(this.right) + this.bottom;
  },
  get $b$r() {
    return '$' + colCache.n2l(this.right) + '$' + this.bottom;
  },

  get range() {
    return this._serialisedSheetName + this.tl + ':' + this.br;
  },
  get $range() {
    return this._serialisedSheetName + this.$t$l + ':' + this.$b$r;
  },
  get shortRange() {
    return this.count > 1 ? this.range : this._serialisedSheetName + this.tl;
  },
  get $shortRange() {
    return this.count > 1 ? this.$range : this._serialisedSheetName + this.$t$l;
  },
  get count() {
    return (1 + this.bottom - this.top) * (1 + this.right - this.left);
  },

  toString: function toString() {
    return this.range;
  },

  intersects: function intersects(other) {
    if (other.sheetName && this.sheetName && other.sheetName !== this.sheetName) return false;
    if (other.bottom < this.top) return false;
    if (other.top > this.bottom) return false;
    if (other.right < this.left) return false;
    if (other.left > this.right) return false;
    return true;
  },
  contains: function contains(addressStr) {
    var address = colCache.decodeEx(addressStr);
    return this.containsEx(address);
  },
  containsEx: function containsEx(address) {
    if (address.sheetName && this.sheetName && address.sheetName !== this.sheetName) return false;
    return address.row >= this.top && address.row <= this.bottom && address.col >= this.left && address.col <= this.right;
  }
};
//# sourceMappingURL=range.js.map
