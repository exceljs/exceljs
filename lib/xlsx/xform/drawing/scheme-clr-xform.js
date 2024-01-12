const BaseXform = require('../base-xform');

class SchemeClrXform extends BaseXform {
  constructor() {
    super();
    this.model = null;
  }

  get tag() {
    return 'a:schemeClr';
  }

  render(xmlStream, model) {
    if(model.schemeColor) {
      xmlStream.openNode(this.tag, {val: model.schemeColor});
      if (model.shade !== 0) {
        xmlStream.leafNode('a:shade', {val: model.shade * 100000});
      }
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
          schemeColor: node.attributes.val,
          shade: 0,
          opacity: 1,
        };
        return true;
      case 'a:alpha':
        this.model.opacity = parseInt(node.attributes.val || '100000', 10) / 100000;
        return true;
      case 'a:shade':
        this.model.shade = parseInt(node.attributes.val || '0', 10) / 100000;
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

module.exports = SchemeClrXform;
