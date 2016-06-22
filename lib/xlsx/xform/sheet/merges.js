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

var Range = require('../../../doc/range');
var colCache = require('../../../utils/col-cache');

var Merges = module.exports = function(mergeCells) {
  // optional mergeCells is array of ranges (like the xml)
  this.merges = {};
};

Merges.prototype = {
  add:  function(merge) {
    // merge is {address, master}
    if (this.merges[merge.master]) {
      this.merges[merge.master].expandToAddress(merge.address);
    } else {
      this.merges[merge.master] = new Range(merge.master + ':' + merge.address);
    }
  },
  get mergeCells() {
    return _.map(this.merges, function(merge) {
      return merge.range.range;
    });
  },
  reconcile: function(mergeCells) {
    // for each range - build hash for each affected address
    var merges = this.merges = {};
    var hash = this.hash = {};
    mergeCells.forEach(function(merge) {
      var range = new Range(merge);
      merges[range.tl] = range;
      for (var r = range.top; r <= range.bottom; r++) {
        for (var c = range.left; c <= range.right; c++) {
          var address = colCache.n2l(c) + r;
          hash[address] = range;
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