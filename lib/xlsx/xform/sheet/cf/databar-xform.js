const BaseXform = require('../../base-xform');

const ColorXform = require('../../style/color-xform');
const CfvoXform = require('./cfvo-xform');

class DatabarXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      cfvo: this.cfvoXform = new CfvoXform(),
      color: this.colorXform = new ColorXform(),
    };
  }

  get tag() {
    return 'dataBar';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
    });
    model.color.forEach(color => {
      this.colorXform.render(xmlStream, color);
    });

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {
          cfvo: [],
          color: [],
        };
        break;

      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
    }
    return false;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model[name].push(this.parser.model);
        this.parser = null;
      }
      return true;
    }
    return name !== this.tag;
  }
}

module.exports = DatabarXform;
