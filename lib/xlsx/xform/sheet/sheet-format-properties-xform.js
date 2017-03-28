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

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var SheetFormatPropertiesXform = module.exports = function() {
};

utils.inherits(SheetFormatPropertiesXform, BaseXform, {

  get tag() { return 'sheetFormatPr'; },

  render: function(xmlStream, model) {
    if (model) {
      var attributes = {
        defaultRowHeight: model.defaultRowHeight,
        outlineLevelRow: model.outlineLevelRow,
        outlineLevelCol: model.outlineLevelCol,
        'x14ac:dyDescent': model.dyDescent
      };
      if (_.some(attributes, function(value) { return value !== undefined; })) {
        xmlStream.leafNode('sheetFormatPr', attributes);
      }
    }
  },
  
  parseOpen: function(node) {
    if (node.name === 'sheetFormatPr') {
      this.model = {
        defaultRowHeight:  parseFloat(node.attributes.defaultRowHeight || 0),
        dyDescent: parseFloat(node.attributes['x14ac:dyDescent'] || 0),
        outlineLevelRow: parseInt(node.attributes.outlineLevelRow || 0),
        outlineLevelCol: parseInt(node.attributes.outlineLevelCol || 0)
      };
      return true;
    } else {
      return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
