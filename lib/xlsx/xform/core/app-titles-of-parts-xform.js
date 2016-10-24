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

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var AppTitlesOfPartsXform = module.exports = function() {
};

utils.inherits(AppTitlesOfPartsXform, BaseXform, {
  render: function(xmlStream, model) {

    xmlStream.openNode('TitlesOfParts');
    xmlStream.openNode('vt:vector', {size: model.length, baseType: 'lpstr'});

    model.forEach(function(sheet) {
      xmlStream.leafNode('vt:lpstr', undefined, sheet.name);
    });
    
    xmlStream.closeNode();
    xmlStream.closeNode();
  },
  
  parseOpen: function(node) {
    // no parsing
    return node.name === 'TitlesOfParts';
  },
  parseText: function() {
  },
  parseClose: function(name) {
    return name !== 'TitlesOfParts';
  }
});
