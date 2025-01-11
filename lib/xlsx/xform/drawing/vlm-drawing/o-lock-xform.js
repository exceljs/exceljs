const BaseXform = require('../../base-xform');

class OLockform extends BaseXform {
  get tag() {
    return 'o:lock';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      'v:ext': 'edit',
      rotation: 't',
      aspectratio: 't',
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {};
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

module.exports = OLockform;
