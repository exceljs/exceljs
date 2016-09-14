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

var AlignmentXform = require('./alignment-xform');

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
// </xf>

// Style assists translation from style model to/from xlsx
var StyleXform = module.exports = function(options) {
  this.xfId = !!(options && options.xfId);
  this.map = {
    alignment: new AlignmentXform()
  };
};

utils.inherits(StyleXform, BaseXform, {

  get tag() { return 'xf'; },

  render: function(xmlStream, model) {
    xmlStream.openNode('xf', {
      numFmtId: model.numFmtId || 0,
      fontId: model.fontId || 0,
      fillId: model.fillId || 0,
      borderId: model.borderId || 0
    });
    if (this.xfId) {
      xmlStream.addAttribute('xfId', model.xfId || 0);
    }

    if (model.numFmtId) {
      xmlStream.addAttribute('applyNumberFormat', '1');
    }
    if (model.fontId) {
      xmlStream.addAttribute('applyFont', '1');
    }
    if (model.fillId) {
      xmlStream.addAttribute('applyFill', '1');
    }
    if (model.borderId) {
      xmlStream.addAttribute('applyBorder', '1');
    }

    if (model.alignment) {
      xmlStream.addAttribute('applyAlignment', '1');
      this.map.alignment.render(xmlStream, model.alignment);
    }

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    // console.log('style-xform',JSON.stringify(node));

    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      // used during sax parsing of xml to build font object
      switch (node.name) {
        case 'xf':
          this.model = {
            numFmtId: parseInt(node.attributes.numFmtId),
            fontId: parseInt(node.attributes.fontId),
            fillId: parseInt(node.attributes.fillId),
            borderId: parseInt(node.attributes.borderId)
          };
          if (this.xfId) {
            this.model.xfId = parseInt(node.attributes.xfId);
          }
          return true;
        case 'alignment':
          this.parser = this.map.alignment;
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
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.alignment = this.parser.model;
        this.parser = undefined;
      }
      return true;
    } else {
      return name !== 'xf';
    }
  }
});
