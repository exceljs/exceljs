const BaseXform = require('../../base-xform');
const XfrmXform = require('../xfrm-xform');
const PrstGeomXform = require('./prst-geom-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties
class SpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'a:xfrm': new XfrmXform(),
      'a:prstGeom': new PrstGeomXform(),
    };
  }

  get tag() {
    return 'xdr:spPr';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag);
    this.map['a:xfrm'].render(xmlStream, shape);
    this.map['a:prstGeom'].render(xmlStream, shape);
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
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpPrXform;
