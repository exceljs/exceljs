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
var defaultNumFormats = require('../../defaultnumformats');

var BaseXform = require('../base-xform');


function hashDefaultFormats() {
  var hash = {};
  _.each(defaultNumFormats, function(dnf, id) {
    if (dnf.f) {
      hash[dnf.f] = parseInt(id);
    }
    // at some point, add the other cultures here...
  });
  return hash;
}
var defaultFmtHash = hashDefaultFormats();


// NumFmt encapsulates translation between number format and xlsx
var NumFmtXform = module.exports = function(id, formatCode) {
  this.id = id;
  this.formatCode = formatCode;
};


utils.inherits(NumFmtXform, BaseXform, {

  get tag() { return 'numFmt'; },

  getDefaultFmtId: function(formatCode) {
    return defaultFmtHash[formatCode];
  },
  getDefaultFmtCode: function(numFmtId) {
    return defaultNumFormats[numFmtId] && defaultNumFormats[numFmtId].f;
  }
},{
  render: function(xmlStream, model) {
    //var formatCode = model.formatCode.replace(/([\s\-\(\)])/g, '\\$1');
    xmlStream.leafNode('numFmt', {numFmtId: model.id, formatCode: model.formatCode});
  },
  parseOpen:  function(node) {
    switch(node.name) {
      case 'numFmt':
        this.model = {
          id: parseInt(node.attributes.numFmtId),
          formatCode: node.attributes.formatCode.replace(/[\\](.)/g, '$1')
        };
        return true;
      default:
        return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
