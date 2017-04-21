/**
 * Copyright (c) 2015 Guyon Roche
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

var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var BaseXform = require('../base-xform');

var AutoFilterXform = module.exports = function() {
};

utils.inherits(AutoFilterXform, BaseXform, {

  get tag() { return 'autoFilter'; },

  render: function(xmlStream, model) {
    function getAddress (autoFilterObject) {
      if (autoFilterObject) {
        if (typeof autoFilterObject === 'string') {
            return colCache.getAddress(autoFilterObject);
        } else if (autoFilterObject.row && autoFilterObject.column) {
            return colCache.getAddress(autoFilterObject.row, autoFilterObject.column);
        }
      }
    }
    if (model) {
      var firstAddress = getAddress(model.from);
      var secondAddress = getAddress(model.to);
      if (firstAddress && secondAddress) {
        xmlStream.leafNode('autoFilter', {ref: firstAddress.address + ':' + secondAddress.address});
      }
    }
  },

  parseOpen: function(node) {
    if (node.name === 'autoFilter') {
      var cells = node.attributes.ref.split(':');
      if (cells.length === 2) {
        var firstAddress = colCache.getAddress(cells[0]);
        var secondAddress = colCache.getAddress(cells[1]);
        this.model = {
          from: {
            row: firstAddress.row,
            column: firstAddress.col
          },
          to: {
            row: secondAddress.row,
            column: secondAddress.col
          }
        }
      }
    }
  }

});