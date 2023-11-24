const BaseXform = require('../../base-xform');

const NvSpPrXform = require('./nv-sp-pr-xform');
const SpPrXform = require('./sp-pr-xform');
const StyleXform = require('./style-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape
class SpXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:nvSpPr': new NvSpPrXform(),
      'xdr:spPr': new SpPrXform(),
      'xdr:style': new StyleXform(),
    };
  }

  get tag() {
    return 'xdr:sp';
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {macro: '', textlink: ''});

    this.map['xdr:nvSpPr'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);
    this.map['xdr:style'].render(xmlStream, model);

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

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
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
        this.mergeModel({
          ...(this.map['xdr:style'].model.fill ? {fill: this.map['xdr:style'].model.fill} : {}),
          ...(this.map['xdr:style'].model.outline ? {outline: this.map['xdr:style'].model.outline} : {}),
          ...this.map['xdr:spPr'].model,
        });
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpXform;
