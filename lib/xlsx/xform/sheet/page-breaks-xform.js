const BaseXform = require('../base-xform');

class PageBreaksXform extends BaseXform {
  get tag() {
    return 'brk';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('brk', model);
  }

  parseOpen(node) {
    if (node.name === this.tag) {
      this.model = {
        id: node.attributes.id ? parseInt(node.attributes.id, 10) : undefined,
        max: node.attributes.max ? parseInt(node.attributes.max, 10) : undefined,
        min: node.attributes.min ? parseInt(node.attributes.min, 10) : undefined,
        man: node.attributes.man ? parseInt(node.attributes.man, 10) : undefined,
      };
      return true;
    }
    return false;
  }

  parseClose() {
    return false;
  }
}

module.exports = PageBreaksXform;
