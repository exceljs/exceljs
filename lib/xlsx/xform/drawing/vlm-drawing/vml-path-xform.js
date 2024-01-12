const BaseXform = require('../../base-xform');

class VmlPathform extends BaseXform {
  get tag() {
    return 'v:path';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      'o:connecttype': model.connecttype,
    });
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          connecttype: node.attributes['o:connecttype'],
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

module.exports = VmlPathform;
