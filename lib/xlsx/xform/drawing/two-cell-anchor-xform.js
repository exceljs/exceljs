/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

const utils = require('../../../utils/utils');
const colCache = require('../../../utils/col-cache');
const BaseXform = require('../base-xform');
const StaticXform = require('../static-xform');

const CellPositionXform = require('./cell-position-xform');
const PicXform = require('./pic-xform');

const TwoCellAnchorXform = (module.exports = function() {
  this.map = {
    'xdr:from': new CellPositionXform({ tag: 'xdr:from' }),
    'xdr:to': new CellPositionXform({ tag: 'xdr:to' }),
    'xdr:pic': new PicXform(),
    'xdr:clientData': new StaticXform({ tag: 'xdr:clientData' }),
  };
});

utils.inherits(TwoCellAnchorXform, BaseXform, {
  get tag() {
    return 'xdr:twoCellAnchor';
  },

  prepare(model, options) {
    this.map['xdr:pic'].prepare(model.picture, options);

    // convert model.range into tl, br
    if (typeof model.range === 'string') {
      const range = colCache.decode(model.range);
      // Note - zero based
      model.tl = {
        col: range.left - 1,
        row: range.top - 1,
      };
      // zero based but also +1 to cover to bottom right of br cell
      model.br = {
        col: range.right,
        row: range.bottom,
      };
    } else {
      model.tl = model.range.tl;
      model.br = model.range.br;
    }
  },

  render(xmlStream, model) {
    if (model.range.editAs) {
      xmlStream.openNode(this.tag, { editAs: model.range.editAs });
    } else {
      xmlStream.openNode(this.tag);
    }

    this.map['xdr:from'].render(xmlStream, model.tl);
    this.map['xdr:to'].render(xmlStream, model.br);
    this.map['xdr:pic'].render(xmlStream, model.picture);
    this.map['xdr:clientData'].render(xmlStream, {});

    xmlStream.closeNode();
  },

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          editAs: node.attributes.editAs,
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

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose(name) {
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
        this.model.br = this.map['xdr:to'].model;
        this.model.picture = this.map['xdr:pic'].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  },

  reconcile(model, options) {
    if (model.picture && model.picture.rId) {
      const rel = options.rels[model.picture.rId];
      const match = rel.Target.match(/.*\/media\/(.+[.][a-z]{3,4})/);
      if (match) {
        const name = match[1];
        const mediaId = options.mediaIndex[name];
        model.medium = options.media[mediaId];
      }
    }
    if (
      model.tl &&
      model.br &&
      Number.isInteger(model.tl.row) &&
      Number.isInteger(model.tl.col) &&
      Number.isInteger(model.br.row) &&
      Number.isInteger(model.br.col)
    ) {
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
  },
});
