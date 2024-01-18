const CompositeXform = require('../../composite-xform');

const FExtXform = require('../cf-ext/f-ext-xform');

class FormulaExtXform extends CompositeXform {
  constructor() {
    super();

    this.map = {
      'xm:f': (this.fExtXform = new FExtXform()),
    };
  }

  get tag() {
    return 'x14:formula1';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {});
    if (model !== undefined && model.length === 1) {
      this.fExtXform.render(xmlStream, model[0]);
    }
    xmlStream.closeNode();
  }

  createNewModel(node) {
    return {};
  }
}

module.exports = FormulaExtXform;
