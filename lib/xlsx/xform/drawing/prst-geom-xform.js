const BaseXform = require('../base-xform');

class PrstGeomXform extends BaseXform {
  get tag() {
    return 'a:prstGeom';
  }

  prepare(model, options) {}

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {prst: model.shape});
    xmlStream.leafNode('a:avLst', {});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          shape: node.attributes.prst
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

module.exports = PrstGeomXform;
