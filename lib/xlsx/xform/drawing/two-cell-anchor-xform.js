/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');

var CellPositionXform = require('./cell-position-xform');
var PicXform = require('./pic-xform');

var TwoCellAnchorXform = module.exports = function() {
  this.map = {
    'from': new CellPositionXform({tag: 'from', ns: 'xdr'}),
    'to': new CellPositionXform({tag: 'to', ns: 'xdr'}),
    'pic': new PicXform(),
    'clientData': new StaticXform({tag: 'clientData', ns: 'xdr'}),
  };
};

utils.inherits(TwoCellAnchorXform, BaseXform, {
  get tag() { return 'twoCellAnchor'; },

  prepare: function(model, options) {
    this.map['pic'].prepare(model.picture, options);

    // convert model.range into tl, br
    if (typeof model.range === 'string') {
      var range = colCache.decode(model.range);
      // Note - zero based
      model.tl = {
        col: range.left - 1,
        row: range.top - 1
      };
      // zero based but also +1 to cover to bottom right of br cell
      model.br = {
        col: range.right,
        row: range.bottom
      };
    } else {
      model.tl = model.range.tl;
      model.br = model.range.br;
    }
  },

  render: function(xmlStream, model) {
    if (model.range.editAs) {
      xmlStream.openNode('xdr:' + this.tag, {editAs: model.range.editAs});
    } else {
      xmlStream.openNode('xdr:' + this.tag);
    }

    this.map['from'].render(xmlStream, model.tl);
    this.map['to'].render(xmlStream, model.br);
    this.map['pic'].render(xmlStream, model.picture);
    this.map['clientData'].render(xmlStream, {});

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
        this.model = {
          editAs: node.attributes.editAs
        };
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
    switch (name) {
      case this.tag:
        this.model = this.model || {};
        this.model.tl = this.map['from'].model;
        this.model.br = this.map['to'].model;
        this.model.picture = this.map['pic'].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  },

  reconcile: function(model, options) {
    if (model.picture && model.picture.rId) {
      var rel = options.rels[model.picture.rId];
      var match = rel.Target.match(/.*\/media\/(.+[.][a-z]{3,4})/);
      if (match) {
        var name = match[1];
        var mediaId = options.mediaIndex[name];
        model.medium = options.media[mediaId];
      }
    }
    if (model.tl && model.br && Number.isInteger(model.tl.row) && Number.isInteger(model.tl.col) && Number.isInteger(model.br.row) && Number.isInteger(model.br.col)) {
      model.range = colCache.encode(model.tl.row + 1, model.tl.col + 1, model.br.row, model.br.col);
    } else {
      model.range = {
        tl: model.tl,
        br: model.br,
      };
    }
    if (model.editAs) {
      model.range.editAs = model.editAs;
      delete model.editAs;
    }
    delete model.tl;
    delete model.br;
  }
});
