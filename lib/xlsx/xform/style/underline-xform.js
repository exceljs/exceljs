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
var BaseXform = require('../base-xform');

// | underline | Font <u>underline</u> style | true, false, 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting' |

var UnderlineXform = module.exports = function(model) {
  this.model = model;
};

UnderlineXform.Attributes = {
  single: {},
  double: { val: 'double' },
  singleAccounting: { val: 'singleAccounting' },
  doubleAccounting: { val: 'doubleAccounting' }
};

utils.inherits(UnderlineXform, BaseXform, {

  get tag() { return 'u'; },

  render: function(xmlStream, model) {
    model = model || this.model;

    if (model === true) {
      xmlStream.leafNode('u')
    } else {
      var attr = UnderlineXform.Attributes[model];
      if (attr) {
        xmlStream.leafNode('u', attr);
      }
    }
  },

  parseOpen: function(node) {
    if (node.name === 'u') {
      this.model = node.attributes.val || true;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
