const BaseXform = require('../../base-xform');

// DocumentFormat.OpenXml.Drawing.PresetGeometry
class PrstGeomXform extends BaseXform {
  constructor() {
    super();

    this.map = {};
  }

  get tag() {
    return 'a:prstGeom';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {prst: model.type});
    xmlStream.leafNode('a:avLst', {});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {type: node.attributes.prst};
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
        // TOOD avLst
        return false;
      default:
        return true;
    }
  }
}

module.exports = PrstGeomXform;
