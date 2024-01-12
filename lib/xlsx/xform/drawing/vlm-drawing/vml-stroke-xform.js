const BaseXform = require('../../base-xform');

class VmlStrokeform extends BaseXform {
  get tag() {
    return 'v:stroke';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      on: model.on || 'f',
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          on: node.attributes.on,
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

module.exports = VmlStrokeform;
