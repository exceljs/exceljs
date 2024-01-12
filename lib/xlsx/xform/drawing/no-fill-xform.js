const BaseXform = require('../base-xform');

class NoFillXform extends BaseXform {
  constructor() {
    super();
    this.model = null;
  }

  get tag() {
    return 'a:noFill';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag);
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          color: '000000',
          opacity: 0,
        };
        return true;
      default:
        return true;
    }
  }

  parseClose(name) {
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

module.exports = NoFillXform;
