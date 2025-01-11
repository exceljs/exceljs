const BaseXform = require('../../base-xform');

class VmlFillform extends BaseXform {
  get tag() {
    return 'v:fill';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      focussize: model.focussize,
      on: model.on,
      color2: model.color2,
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          // 'f'
          on: node.attributes.on,
          focussize: node.attributes.focussize, // '0.0'
          color2: node.attributes.color2,
        };
        return true;
      default:
        return true;
    }
  }

  parseText() {}

  parseClose(name) {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes
        return true;
    }
  }
}

module.exports = VmlFillform;
