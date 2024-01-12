const BaseXform = require('../../base-xform');
const XfrmXform = require('../xfrm-xform');
const PrstGeomXform = require('./prst-geom-xform');
const SolidFillXform = require('./solid-fill-xform');
const LnXform = require('./ln-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties
class SpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:xfrm': new XfrmXform(),
      'a:prstGeom': new PrstGeomXform(),
      'a:solidFill': new SolidFillXform(),
      'a:ln': new LnXform(),
    };
  }

  get tag() {
    return 'xdr:spPr';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag);
    this.map['a:xfrm'].render(xmlStream, shape);
    this.map['a:prstGeom'].render(xmlStream, shape);
    if (shape.fill && shape.fill.type === 'solid') {
      this.map['a:solidFill'].render(xmlStream, shape.fill.color);
    } else {
      xmlStream.leafNode('a:noFill');
    }
    if (shape.outline) {
      this.map['a:ln'].render(xmlStream, shape.outline);
    }
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {};
        this.noFill = false;
        break;
      case 'a:noFill':
        this.noFill = true;
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

  parseText() {}

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        if (this.map['a:prstGeom'].model) {
          this.model.type = this.map['a:prstGeom'].model.type;
        }
        if (this.map['a:solidFill'].model) {
          this.model.fill = {
            type: 'solid',
            color: this.map['a:solidFill'].model,
          };
        }
        if (this.map['a:ln'].model) {
          this.model.outline = this.map['a:ln'].model;
        }
        if (this.map['a:xfrm'].model) {
          this.mergeModel(this.map['a:xfrm'].model);
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpPrXform;
