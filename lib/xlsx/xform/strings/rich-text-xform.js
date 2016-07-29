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

var TextXform = require('./text-xform');
var FontXform = require('../style/font-xform');

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

// <r>
//   <rPr>
//     <sz val="11"/>
//     <color theme="1" tint="5"/>
//     <rFont val="Calibri"/>
//     <family val="2"/>
//     <scheme val="minor"/>
//   </rPr>
//   <t xml:space="preserve"> is </t>
// </r>

var RichTextXform = module.exports = function(model) {
  this.model = model;
};

RichTextXform.FONT_OPTIONS = {
  tagName: 'rPr',
  fontNameTag:  'rFont'
};

utils.inherits(RichTextXform, BaseXform, {

  get tag() { return 'r'; },

  get textXform() { return this._textXform || (this._textXform = new TextXform()); },
  get fontXform() { return this._fontXform || (this._fontXform = new FontXform(RichTextXform.FONT_OPTIONS)); },

  render: function(xmlStream, model) {
    model = model || this.model;

    xmlStream.openNode('r');
    if (model.font) {
      this.fontXform.render(xmlStream, model.font);
    }
    this.textXform.render(xmlStream, model.text);
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'r':
          this.model = {};
          return true;
        case 't':
          this.parser = this.textXform;
          this.parser.parseOpen(node);
          return true;
        case 'rPr':
          this.parser = this.fontXform;
          this.parser.parseOpen(node);
          return true;
        default:
          return false;
      }
    }
  },
  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function(name) {
    switch(name) {
      case 'r':
        return false;
      case 't':
        this.model.text = this.parser.model;
        this.parser = undefined;
        return true;
      case 'rPr':
        this.model.font = this.parser.model;
        this.parser = undefined;
        return true;
      default:
        if (this.parser) {
          this.parser.parseClose(name);
        }
        return true;
    }
  }
});
