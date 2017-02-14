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
var _ = require('../../../utils/under-dash');

var Range = require('../../../doc/range');
var colCache = require('../../../utils/col-cache');
var Enums = require('../../../doc/enums');

var Merges = module.exports = function() {
  // optional mergeCells is array of ranges (like the xml)
  this.merges = {};
};

Merges.prototype = {
  add:  function(merge) {
    // merge is {address, master}
    if (this.merges[merge.master]) {
      this.merges[merge.master].expandToAddress(merge.address);
    } else {
      var range = merge.master + ':' + merge.address;
      this.merges[merge.master] = new Range(range);
    }
  },
  get mergeCells() {
    return _.map(this.merges, function(merge) {
      return merge.range;
    });
  },
  reconcile: function(mergeCells, rows) {
    // reconcile merge list with merge cells
    _.each(mergeCells, function (merge) {
      var dimensions = colCache.decode(merge);
      for (var i = dimensions.top; i <= dimensions.bottom; i++) {
        var row = rows[i-1];
        for (var j = dimensions.left; j <= dimensions.right; j++) {
          var cell = row.cells[j-1];
          if (!cell) {
            // nulls are not included in document - so if master cell has no value - add a null one here
            row.cells[j] = {
              type: Enums.ValueType.Null,
              address: colCache.encodeAddress(i, j)
            };
          } else if (cell.type === Enums.ValueType.Merge) {
            cell.master = dimensions.tl;
          }
        }
      }
    });
  },
  getMasterAddress:  function(address) {
    // if address has been merged, return its master's address. Assumes reconcile has been called
    var range = this.hash[address];
    if (range) {
      return range.tl;
    }
  }
};