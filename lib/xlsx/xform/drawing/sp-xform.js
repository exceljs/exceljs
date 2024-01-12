const BaseXform = require('../base-xform');
const StaticXform = require('../static-xform');
const NvSpPrXform = require('./nv-sp-pr-xform');
const SpPrXform = require('./sp-pr-xform');
const StyleXform = require('./style-xform');

const txBodyJSON = require('./tx-body');

class SpXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:nvSpPr': new NvSpPrXform(),
      'xdr:spPr': new SpPrXform(),
      'xdr:style': new StyleXform(),
      'xdr:txBody': new StaticXform(txBodyJSON),
    };
  }

  get tag() {
    return 'xdr:sp';
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    this.map['xdr:nvSpPr'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);
    this.map['xdr:style'].render(xmlStream, model);
    this.map['xdr:txBody'].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node) {
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
        const
          fill = Object.assign({}, this.map['xdr:style'].model.fill, this.map['xdr:spPr'].model.fill),
          stroke = Object.assign({}, this.map['xdr:style'].model.stroke, this.map['xdr:spPr'].model.stroke)

        this.model = Object.assign(
          {},
          this.map['xdr:spPr'].model,
          Object.keys(fill).length ? {fill} : null,
          Object.keys(stroke).length ? {stroke} : null,
          this.map['xdr:nvSpPr'].model,
        );
        return false;
      default:
        return true;
    }
  }
}

module.exports = SpXform;
