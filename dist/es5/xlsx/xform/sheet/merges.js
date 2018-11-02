/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _ = require('../../../utils/under-dash');

var Range = require('../../../doc/range');
var colCache = require('../../../utils/col-cache');
var Enums = require('../../../doc/enums');

var Merges = module.exports = function () {
  // optional mergeCells is array of ranges (like the xml)
  this.merges = {};
};

Merges.prototype = {
  add: function add(merge) {
    // merge is {address, master}
    if (this.merges[merge.master]) {
      this.merges[merge.master].expandToAddress(merge.address);
    } else {
      var range = merge.master + ':' + merge.address;
      this.merges[merge.master] = new Range(range);
    }
  },
  get mergeCells() {
    return _.map(this.merges, function (merge) {
      return merge.range;
    });
  },
  reconcile: function reconcile(mergeCells, rows) {
    // reconcile merge list with merge cells
    _.each(mergeCells, function (merge) {
      var dimensions = colCache.decode(merge);
      for (var i = dimensions.top; i <= dimensions.bottom; i++) {
        var row = rows[i - 1];
        for (var j = dimensions.left; j <= dimensions.right; j++) {
          var cell = row.cells[j - 1];
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
  getMasterAddress: function getMasterAddress(address) {
    // if address has been merged, return its master's address. Assumes reconcile has been called
    var range = this.hash[address];
    return range && range.tl;
  }
};
//# sourceMappingURL=merges.js.map
