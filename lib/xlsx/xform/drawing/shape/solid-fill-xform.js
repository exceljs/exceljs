const BaseXform = require('../../base-xform');

// DocumentFormat.OpenXml.Drawing.SolidFill
class SolidFillXform extends BaseXform {
  constructor() {
    super();

    this.map = {};
  }

  get tag() {
    return 'a:solidFill';
  }

  render(xmlStream, color) {
    xmlStream.openNode(this.tag);
    if (color.theme) {
      xmlStream.leafNode('a:schemeClr', {val: color.theme});
    } else if (color.rgb) {
      xmlStream.leafNode('a:srgbClr', {val: color.rgb});
    }
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {};
        break;
      case 'a:schemeClr':
        this.model.theme = node.attributes.val;
        break;
      case 'a:srgbClr':
        this.model.rgb = node.attributes.val;
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

module.exports = SolidFillXform;
