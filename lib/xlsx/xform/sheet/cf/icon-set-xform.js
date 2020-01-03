const BaseXform = require('../../base-xform');

const CfvoXform = require('./cfvo-xform');

class IconSetXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      cfvo: this.cfvoXform = new CfvoXform(),
    };
  }

  get tag() {
    return 'iconSet';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      iconSet: BaseXform.toStringAttribute(model.iconSet, '3TrafficLights'),
      reverse: BaseXform.toBoolAttribute(model.reverse, false),
      showValue: BaseXform.toBoolAttribute(model.showValue, true),
    });

    model.cfvo.forEach(cfvo => {
      this.cfvoXform.render(xmlStream, cfvo);
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
          iconSet: BaseXform.toStringValue(node.attributes.iconSet, '3TrafficLights'),
          reverse: BaseXform.toBoolValue(node.attributes.reverse),
          showValue: BaseXform.toBoolValue(node.attributes.showValue),
          cfvo: [],
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

module.exports = IconSetXform;
