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

var DefinedNamesXform = module.exports = function() {
};

utils.inherits(DefinedNamesXform, BaseXform, {
  render: function(xmlStream, model) {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode('definedName');
    xmlStream.addAttribute('name', model.name);
    xmlStream.writeText(model.ranges.join(','));
    xmlStream.closeNode();
  },
  parseOpen: function(node) {
    switch(node.name) {
      case 'definedName':
        this._parsedName = node.attributes.name;
        this._parsedText = [];
        return true;
      default:
        return false;
    }
  },
  parseText: function(text) {
    this._parsedText.push(text);
  },
  parseClose: function() {
    this.model = {
      name: this._parsedName,
      ranges: this._parsedText.join('')
        .split(',')
        .filter(function (range) {
          return range;
        })
    };
    return false;
  }
});