const BaseXform = require('../base-xform');

class LegacyDrawingFHXform extends BaseXform {
  get tag() {
    return 'legacyDrawingHF';
  }

  render(xmlStream, model) {
    if (model) {
      xmlStream.leafNode('legacyDrawingHF', {'r:id': model.rId});
    }
  }

  parseOpen(node) {
    if (node.name === 'legacyDrawingHF') {
      this.model = {
        rId: node.attributes['r:id'],
      };
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = LegacyDrawingFHXform;
