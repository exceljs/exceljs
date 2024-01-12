const BaseXform = require('../base-xform');

class XfrmXform extends BaseXform {
  get tag() {
    return 'a:xfrm';
  }

  prepare(model, options) {}

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {rot: model.rotation * 60000});
    xmlStream.leafNode('a:off', {x: 0, y: 0});
    xmlStream.leafNode('a:ext', {cx: 0, cy: 0});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {rotation: (node.attributes.rot || 0) / 60000};
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

module.exports = XfrmXform;
