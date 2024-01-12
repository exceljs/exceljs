const BaseXform = require('../base-xform');

const SchemeClrXform = require('./scheme-clr-xform');
const SrgbClrXform = require('./srgb-clr-xform');

class SolidFillXform extends BaseXform {
  constructor() {
    super();
    this.model = null;
    this.map = {
      'a:srgbClr': new SrgbClrXform(),
      'a:schemeClr': new SchemeClrXform(),
    };
  }

  get tag() {
    return 'a:solidFill';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);
    if(model.color) {
      this.map['a:srgbClr'].render(xmlStream, model);
    } else {
      this.map['a:schemeClr'].render(xmlStream, model);
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
        this.model = Object.assign(
          {},
          this.map['a:schemeClr'].model,
          this.map['a:srgbClr'].model,
        );
        if(!Object.keys(this.model).length) {
          this.model = null;
        }
        return false;
      default:
        return true;
    }
  }
}

module.exports = SolidFillXform;
