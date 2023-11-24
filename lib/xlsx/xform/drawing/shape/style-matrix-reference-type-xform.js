const BaseXform = require('../../base-xform');

// DocumentFormat.OpenXml.Drawing.StyleMatrixReferenceType
class StyleMatrixReferenceTypeXform extends BaseXform {
  constructor(tagName) {
    super();

    this.map = {};
    this.tagName = tagName;
  }

  get tag() {
    return this.tagName;
  }

  idx() {
    switch (this.tagName) {
      case 'a:lnRef':
        return 2;
      case 'a:fillRef':
        return 1;
      default:
        // do not know when to come here
        return 0;
    }
  }

  render(xmlStream) {
    xmlStream.openNode(this.tag, {idx: this.idx()});
    xmlStream.leafNode('a:schemeClr', {val: 'accent1'});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
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

module.exports = StyleMatrixReferenceTypeXform;
