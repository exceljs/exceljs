const BaseXform = require('../base-xform');

// DocumentFormat.OpenXml.Drawing.Transform2D
class XfrmXform extends BaseXform {
  constructor() {
    super();
    this.map = {};
  }

  get tag() {
    return 'a:xfrm';
  }

  render(xmlStream) {
    xmlStream.openNode(this.tag);
    xmlStream.leafNode('a:off', {x: 0, y: 0});
    xmlStream.leafNode('a:ext', {cx: 0, cy: 0});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText() {}

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

module.exports = XfrmXform;
