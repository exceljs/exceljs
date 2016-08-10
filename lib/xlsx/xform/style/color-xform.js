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

//<color auto="1"/>
//<color rgb="FFFF0000"/>
//<color theme="0" tint="-0.499984740745262"/>
//<color indexed="64"/>

// var modelSchema = {
//   // argb style
//   argb: 'AARRGGBB',
//
//   // theme style
//   theme: 0, // any valid theme integer
//   tint: 0.5, // optional. float value from -1 to 1
//
//   // indexed style
//   indexed: 64, // a valid color index colour
//
//   // auto
//   auto: true // optional - is the default state
// };

// Color encapsulates translation from color model to/from xlsx
var ColorXform = module.exports = function(name) {
  // this.name controls the xm node name
  this.name = name || 'color';
};

utils.inherits(ColorXform, BaseXform, {

  get tag() { return this.name; },

  render: function(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.name);
      if (model.argb) {
        xmlStream.addAttribute('rgb', model.argb);
      } else if (model.theme) {
        xmlStream.addAttribute('theme', model.theme);
        if (model.tint) {
          xmlStream.addAttribute('tint', model.tint);
        }
      } else if (model.indexed) {
        xmlStream.addAttribute('indexed', model.indexed)
      } else {
        xmlStream.addAttribute('auto', '1');
      }
      xmlStream.closeNode();
      return true;
    }
  },
  
  parseOpen: function(node) {
    if (node.name === this.name) {
      if (node.attributes.rgb) {
        this.model = {argb: node.attributes.rgb};
      } else if (node.attributes.theme) {
        this.model = {theme: parseInt(node.attributes.theme)};
        if (node.attributes.tint) {
          this.model.tint = parseFloat(node.attributes.tint);
        }
      } else if (node.attributes.indexed) {
        this.model = {indexed: parseInt(node.attributes.indexed)};
      } else {
        this.model = undefined;
      }
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
