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

var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');

var CellPositionXform = require('./cell-position-xform');
var PicXform = require('./pic-xform');

var TwoCellAnchorXform = module.exports = function() {
  this.map = {
    'xdr:from':  new CellPositionXform({tag: 'xdr:from'}),
    'xdr:to':  new CellPositionXform({tag: 'xdr:to'}),
    'xdr:pic': new PicXform(),
    'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
  };
};

utils.inherits(TwoCellAnchorXform, BaseXform, {
  get tag() { return 'xdr:twoCellAnchor'; },

  prepare: function(model, options) {
    model.picture = {
      rId: model.rId,
    };
    this.map['xdr:pic'].prepare(model.picture, options);

    // convert model.range into tl, br
    if (typeof model.range === 'string') {
      var range = colCache.decode(model.range);
      // Note - zero based
      model.tl = {
        col: range.left - 1,
        row: range.top - 1,
      };
      // add one here to include to bottom-right of br
      model.br = {
        col: range.right,
        row: range.bottom
      }
    } else {
      model.tl = model.range.tl;
      model.br = model.range.br;
    }
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
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  },

  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch(name) {
      case this.tag:
        this.model = {
          tl: this.map['xdr:from'].model,
          br: this.map['xdr:to'].model,
          picture: this.map['xdr:pic'].model,
        }
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  },

  reconcile: function(model, options) {
    console.log('twoCellAnchorXform.reconcile', model);
    var rel = options.rels[model.picture.rId];
    console.log('rel', rel);
    console.log('options.media', options.media);
    var match = rel.Target.match(/.*\/media\/(.+\.[a-z]{3,4})/);
    if (match) {
      var name = match[1];
      var mediaId = options.mediaIndex[name];
      model.medium = options.media[mediaId];
    }
    model.range = colCache.encode(model.tl.row, model.tl.col, model.br.row, model.br.col);
  }
});
