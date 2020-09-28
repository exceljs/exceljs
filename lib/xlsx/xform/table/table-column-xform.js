const BaseXform = require('../base-xform');
const CalculatedColumnFormulaXform = require('./calculated-column-formula-xform');

class TableColumnXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      calculatedColumnFormula: new CalculatedColumnFormulaXform(),
    };
  }

  get tag() {
    return 'tableColumn';
  }

  prepare(model, options) {
    model.id = options.index + 1;
  }

  render(xmlStream, model) {
    const nodeAttributes = {
      id: model.id.toString(),
      name: model.name,
      totalsRowLabel: model.totalsRowLabel,
      totalsRowFunction: model.totalsRowFunction,
      dxfId: model.dxfId,
    };
    if (!model.calculatedColumnFormula) {
      xmlStream.leafNode(this.tag, nodeAttributes);
      return true;
    }

    xmlStream.openNode(this.tag, nodeAttributes);
    this.map.calculatedColumnFormula.render(xmlStream, model.calculatedColumnFormula);
    xmlStream.closeNode();

    return true;
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag: {
        const {attributes} = node;
        this.model = {
          name: attributes.name,
          totalsRowLabel: attributes.totalsRowLabel,
          totalsRowFunction: attributes.totalsRowFunction,
          dxfId: attributes.dxfId,
        };
        return true;
      }
      default: {
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
        break;
      }
    }

    return false;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.calculatedColumnFormula = this.parser.model;
        this.parser = undefined;
      }
      return true;
    }

    return false;
  }
}

module.exports = TableColumnXform;
