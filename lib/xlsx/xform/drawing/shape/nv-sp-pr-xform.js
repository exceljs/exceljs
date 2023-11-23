const BaseXform = require('../../base-xform');
const CNvPrXform = require('../c-nv-pr-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeProperties
class NvSpPrXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:cNvPr': new CNvPrXform(false),
    };
  }

  get tag() {
    return 'xdr:nvSpPr';
  }

  render(xmlStream, shape) {
    xmlStream.openNode(this.tag);
    this.map['xdr:cNvPr'].render(xmlStream, shape);
    xmlStream.leafNode('xdr:cNvSpPr');
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
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
        this.model = this.map['xdr:cNvPr'].model;
        return false;
      default:
        return true;
    }
  }
}

module.exports = NvSpPrXform;
