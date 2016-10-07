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

var utils = require('../../utils/utils');
var BaseXform = require('./base-xform');
var XmlStream = require('../../utils/xml-stream');

// var model = {
//   tag: 'name',
//   $: {attr: 'value'},
//   c: [
//     { tag: 'child' }
//   ],
//   t: 'some text'
// };

function build(xmlStream, model) {
  xmlStream.openNode(model.tag, model.$);
  if (model.c) {
    model.c.forEach(function (child) {
      build(xmlStream, child);
    });
  }
  if (model.t) {
    xmlStream.writeText(model.t);
  }
  xmlStream.closeNode();
}

var StaticXform = module.exports = function(model) {
  // This class is an optimisation for static (unimportant and unchanging) xml
  // It is stateless - apart from its static model and so can be used as a singleton
  // Being stateless - it will only track entry to and exit from it's root xml tag during parsing and nothing else
  // Known issues:
  //    since stateless - parseOpen always returns true. Parent xform must know when to start using this xform
  //    if the root tag is recursive, the parsing will behave unpredictably
  this._model = model;
};

utils.inherits(StaticXform, BaseXform, {
  render: function(xmlStream) {
    if (!this._xml) {
      var stream = new XmlStream();
      build(stream, this._model);
      this._xml = stream.xml;
    }
    xmlStream.writeXml(this._xml);
  },

  parseOpen: function() {
    return true;
  },
  parseText: function() {
  },
  parseClose: function(name) {
    switch(name) {
      case this._model.tag:
        return false;
      default:
        return true;
    }
  }
});
