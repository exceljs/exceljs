const BaseXform = require('../../base-xform');
const XfrmXform = require('../xfrm-xform');
const PrstGeomXform = require('./prst-geom-xform');
const SolidFillXform = require('./solid-fill-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties
class SpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:xfrm': new XfrmXform(),
      'a:prstGeom': new PrstGeomXform(),
      'a:solidFill': new SolidFillXform(),
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
        this.mergeModel({
          ...this.map['a:xfrm'].model,
          ...this.map['a:prstGeom'].model,
        });
        if (this.map['a:solidFill'].model) {
          this.mergeModel({
            fill: {
              type: 'solid',
              color: this.map['a:solidFill'].model,
            },
          });
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpPrXform;
