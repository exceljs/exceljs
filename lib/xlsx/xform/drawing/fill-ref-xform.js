const BaseXform = require('../base-xform');

const SchemeClrXform = require('./scheme-clr-xform');

class FillRefXform extends BaseXform {
  constructor() {
    super();
    this.model = null;
    this.map = {
      'a:schemeClr': new SchemeClrXform(),
    };
  }

  get tag() {
    return 'a:fillRef';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {idx: 1});
    this.map['a:schemeClr'].render(xmlStream, {schemeColor: 'accent1', shade: 0, opacity: 1});
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

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }

    switch (name) {
      case this.tag:
        this.model = this.map['a:schemeClr'].model;
        return false;
      default:
        return true;
    }
  }
}

module.exports = FillRefXform;
