const BaseXform = require('../../base-xform');

class VmlShadowform extends BaseXform {
  get tag() {
    return 'v:shadow';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, {
      on: model.on || 'f',
      focussize: model.focussize || '0,0',
      color2: model.color2 || 'infoBackground [80]',
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

module.exports = VmlShadowform;
