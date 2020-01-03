const BaseXform = require('../base-xform');

class HyperlinkXform extends BaseXform {
  get tag() {
    return 'hyperlink';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('hyperlink', {
      ref: model.address,
      'r:id': model.rId,
      tooltip: model.tooltip,
    });
  }

  parseOpen(node) {
    if (node.name === 'hyperlink') {
      this.model = {
        address: node.attributes.ref,
        rId: node.attributes['r:id'],
        tooltip: node.attributes.tooltip,
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

module.exports = HyperlinkXform;
