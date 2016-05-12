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

var XmlStream = require('../../utils/xml-stream');

var ColorXform = require('./color-xform');
var SizeXform = require('./size-xform');

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

// Size encapsulates translation from size model to/from xlsx
var RichTextXform = module.exports = function(model, name) {
  // this.name controls the xm node name
  this.model = model;
};

RichTextXform.prototype = {

  get colorXform() { return this._colorXform || (this._colorXform = new ColorXform()); },
  get sizeXform() { return this._sizeXform || (this._sizeXform = new SizeXform()); },


  write: function(xmlStream, model) {
    model = model || this.model;

    xmlStream.openNode('r');
    _.each(model, function(richText) {
      if (richText.style) {
        xmlStream.openNode('rPr');
        this.sizeXform.write(xmlStream, richText.style.size);
        this.colorXform.write(xmlStream, richText.style.color);
        xmlStream.closeNode();
      }
      xmlStream.openNode('t');
      xmlStream.writeText(richText.text);
      xmlStream.closeNode();
    });
    xmlStream.closeNode();
  },

  parse: function(node) {
    this.model = node.attributes.val;
  }
};
