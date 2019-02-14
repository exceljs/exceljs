/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');

var CellPositionXform = require('./cell-position-xform');
var ExtXform = require('./ext-xform');
var PicXform = require('./pic-xform');

var OneCellAnchorXform = module.exports = function() {
  this.map = {
    'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
    'xdr:ext': new ExtXform({tag: 'xdr:ext'}),
    'xdr:pic': new PicXform(),
    'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
  };
};

utils.inherits(OneCellAnchorXform, BaseXform, {
  get tag() { return 'xdr:oneCellAnchor'; },

  prepare: function(model, options) {
    this.map['xdr:pic'].prepare(model.picture, options);

    model.tl = model.range.tl;
    model.ext = model.range.ext;
  },

  render: function(xmlStream, model) {
    if (model.range.editAs) {
      xmlStream.openNode(this.tag, {editAs: model.range.editAs});
    } else {
      xmlStream.openNode(this.tag);
    }

    this.map['xdr:from'].render(xmlStream, model.tl);
    this.map['xdr:ext'].render(xmlStream, model.ext);
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
        this.model.tl = this.map['xdr:from'].model;
        this.model.ext = this.map['xdr:ext'].model;
        this.model.picture = this.map['xdr:pic'].model;
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
    model.range = {
      tl: model.tl,
      ext: model.ext,
    };
    if (model.editAs) {
      model.range.editAs = model.editAs;
      delete model.editAs;
    }
    delete model.tl;
    delete model.ext;
  }
});
