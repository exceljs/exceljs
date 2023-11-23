const BaseCellAnchorXform = require('./base-cell-anchor-xform');
const StaticXform = require('../static-xform');

const CellPositionXform = require('./cell-position-xform');
const ExtXform = require('./ext-xform');
const PicXform = require('./picture/pic-xform');
const SpXform = require('./shape/sp-xform');

class OneCellAnchorXform extends BaseCellAnchorXform {
  constructor() {
    super();

    this.map = {
      'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
      'xdr:ext': new ExtXform({tag: 'xdr:ext'}),
      'xdr:pic': new PicXform(),
      'xdr:sp': new SpXform(),
      'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
    };
  }

  get tag() {
    return 'xdr:oneCellAnchor';
  }

  prepare(model, options) {
    if (model.picture) {
      this.map['xdr:pic'].prepare(model.picture, options);
    }
    if (model.shape) {
      this.map['xdr:sp'].prepare(model.shape, options);
    }
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {editAs: model.range.editAs || 'oneCell'});

    this.map['xdr:from'].render(xmlStream, model.range.tl);
    this.map['xdr:ext'].render(xmlStream, model.range.ext);
    if (model.picture) {
      this.map['xdr:pic'].render(xmlStream, model.picture);
    }
    if (model.shape) {
      this.map['xdr:sp'].render(xmlStream, model.shape);
    }
    this.map['xdr:clientData'].render(xmlStream, {});

    xmlStream.closeNode();
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model.range.tl = this.map['xdr:from'].model;
        this.model.range.ext = this.map['xdr:ext'].model;
        this.model.picture = this.map['xdr:pic'].model;
        this.model.shape = this.map['xdr:sp'].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model, options) {
    model.medium = this.reconcilePicture(model.picture, options);
  }
}

module.exports = OneCellAnchorXform;
