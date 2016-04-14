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

var XmlStream = require('../utils/xml-stream');

var DefinedNamesXform = module.exports = function(model) {
  this.model = model || [];
};

DefinedNamesXform.prototype = {
  get xml() {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    // </definedNames>

    if (this.model && this.model.length) {
      var xml = new XmlStream();
      xml.openNode('definedNames');
      _.each(this.model, function(definedName) {
        xml.openNode('definedName');
        xml.addAttribute('name', definedName.name);
        xml.writeText(definedName.ranges.join(','));
        xml.closeNode();
      });
      return xml.xml;

      // var xml = ['<definedNames>'];
      // _.each(this.model, function(definedName) {
      //   var definedNameXml = '<definedName name="' + definedName.name + '">' + definedName.ranges.join(',') + '</definedName>';
      //   xml.push(definedNameXml);
      // });
      // xml.push('</definedNames>');
      // return xml.join('');
    } else {
      return '';
    }
  },
  parseOpen: function(node) {
    this._parsedName = node.attributes.name;
    this._parsedText = [];
  },
  parseText: function(text) {
    this._parsedText.push(text);
  },
  parseClose: function(name) {
    this.model.push({
      name: this._parsedName,
      ranges: this._parsedText.join('')
        .split(',')
        .filter(function (range) {
          return range;
        })
    });
    this._parsedName = undefined;
    this._parsedText = undefined;
  }
};