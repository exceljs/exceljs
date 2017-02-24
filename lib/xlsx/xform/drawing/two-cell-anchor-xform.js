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

var fs =  require('fs');
var _ = require('lodash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var TemplateXform = require('../template-xform');

var CellPositionXform = require('./cell-position-xform');
var PicXform = require('./pic-xform');

var TwoCellAnchorXform = module.exports = function() {
  this.map = {
    'xdr:from':  new CellPositionXform({tag: 'xdr:from'}),
    'xdr:to':  new CellPositionXform({tag: 'xdr:to'}),
    'xdr:pic': new PicXform(),
    'xdr:clientData': new TemplateXform('<xdr:clientData/>')
  };
};

utils.inherits(TwoCellAnchorXform, BaseXform, {

  get tag() { return 'xdr:twoCellAnchor'; },

  prepare: function(model, options) {
    this.map['xdr:pic'].prepare(model.picture, options);
  },

  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:from'].render(xmlStream, model.tl);
    this.map['xdr:to'].render(xmlStream, model.br);
    this.map['xdr:pic'].render(xmlStream, model.picture);
    this.map['xdr:clientData'].render(xmlStream, {});

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch (node.name) {
        case this.tag:
          _.each(this.map, function(xform) {
            xform.reset();
          });
          break;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
          break;
      }
      return true;
    }
  },

  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose: function() {
    return false;
  }
});
