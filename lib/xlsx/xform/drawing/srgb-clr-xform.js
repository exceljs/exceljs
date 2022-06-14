const BaseXform = require('../base-xform');

class SrgbClrXform extends BaseXform {
  constructor() {
    super();
    this.model = null;
  }

  get tag() {
    return 'a:srgbClr';
  }

  render(xmlStream, model) {
    if(model.color) {
      xmlStream.openNode(this.tag, {val: model.color});
      if (model.opacity < 1) {
        xmlStream.leafNode('a:alpha', {val: model.opacity * 100000});
      }
      xmlStream.closeNode();
    }
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          color: node.attributes.val,
          opacity: 1,
        };
        return true;
      case 'a:alpha':
        this.model.opacity = parseInt(node.attributes.val || '100000', 10) / 100000;
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

module.exports = SrgbClrXform;
