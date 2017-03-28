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
var StaticXform = require('../static-xform');

var BlipFillXform = require('./blip-fill-xform');

var spPrJSON = require('./sp-pr.json');

var templates = {
  nvPicPr: fs.readFileSync(__dirname + '/nv-pic-pr.xml'),
  spPr: fs.readFileSync(__dirname + '/sp-pr.xml')
};

var PicXform = module.exports = function(options) {
  this.map = {
    'xdr:nvPicPr': new TemplateXform(templates.nvPicPr),
    'xdr:blipFill': new BlipFillXform(),
    'xdr:spPr': new StaticXform(spPrJSON)
  };
};

utils.inherits(PicXform, BaseXform, {

  get tag() { return 'xdr:pic'; },

  prepare: function(model, options) {
    model.index = options.index +  1;
  },

  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:nvPicPr'].render(xmlStream, model);
    this.map['xdr:blipFill'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);

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

  parseText: function() {
  },

  parseClose: function() {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    } else {
      switch(name) {
        case this.tag:
          this.model = {
            rId: this.map['a:blip'].model
          };
          return false;
        default:
          // not quite sure how we get here!
          return true;
      }
    }
  }
});
