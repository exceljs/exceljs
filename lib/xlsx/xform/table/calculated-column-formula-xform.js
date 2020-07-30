const BaseXform = require('../base-xform');

class CalculatedColumnFormulaXform extends BaseXform {
  get tag() {
    return 'calculatedColumnFormula';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, null, model);
    return true;
  }

  parseOpen(node) {
    if (node.name === this.tag) {
      this.model = null;
      return true;
    }

    return false;
  }

  parseText(text) {
    this.model = text;
  }

  parseClose() {
    return false;
  }
}

module.exports = CalculatedColumnFormulaXform;
