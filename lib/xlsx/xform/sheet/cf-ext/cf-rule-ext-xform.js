const crypto = require('crypto');
const BaseXform = require('../../base-xform');
const CompositeXform = require('../../composite-xform');

const DatabarExtXform = require('./databar-ext-xform');
const IconSetExtXform = require('./icon-set-ext-xform');

const randomUUID = () => {
  // support browser globalThis is window
  // eslint-disable-next-line n/no-unsupported-features/node-builtins
  if (globalThis.crypto && typeof globalThis.crypto.randomUUID === 'function') {
    // eslint-disable-next-line n/no-unsupported-features/node-builtins
    return globalThis.crypto.randomUUID();
  }

  return crypto.randomUUID();
};

const extIcons = {
  '3Triangles': true,
  '3Stars': true,
  '5Boxes': true,
};

class CfRuleExtXform extends CompositeXform {
  constructor() {
    super();

    this.map = {
      'x14:dataBar': (this.databarXform = new DatabarExtXform()),
      'x14:iconSet': (this.iconSetXform = new IconSetExtXform()),
    };
  }

  get tag() {
    return 'x14:cfRule';
  }

  static isExt(rule) {
    // is this rule primitive?
    if (rule.type === 'dataBar') {
      return DatabarExtXform.isExt(rule);
    }
    if (rule.type === 'iconSet') {
      if (rule.custom || extIcons[rule.iconSet]) {
        return true;
      }
    }
    return false;
  }

  prepare(model) {
    if (CfRuleExtXform.isExt(model)) {
      model.x14Id = `{${randomUUID()}}`.toUpperCase();
    }
  }

  render(xmlStream, model) {
    if (!CfRuleExtXform.isExt(model)) {
      return;
    }

    switch (model.type) {
      case 'dataBar':
        this.renderDataBar(xmlStream, model);
        break;
      case 'iconSet':
        this.renderIconSet(xmlStream, model);
        break;
    }
  }

  renderDataBar(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'dataBar',
      id: model.x14Id,
    });

    this.databarXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderIconSet(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'iconSet',
      priority: model.priority,
      id: model.x14Id || `{${randomUUID()}}`,
    });

    this.iconSetXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  createNewModel({attributes}) {
    return {
      type: attributes.type,
      x14Id: attributes.id,
      priority: BaseXform.toIntValue(attributes.priority),
    };
  }

  onParserClose(name, parser) {
    Object.assign(this.model, parser.model);
  }
}

module.exports = CfRuleExtXform;
