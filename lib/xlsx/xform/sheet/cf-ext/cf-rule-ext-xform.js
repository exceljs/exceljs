const {v4: uuidv4} = require('uuid');
const BaseXform = require('../../base-xform');
const CompositeXform = require('../../composite-xform');

const DatabarExtXform = require('./databar-ext-xform');
const IconSetExtXform = require('./icon-set-ext-xform');
const DxfExtXform = require('../../style/dxf-ext-xform');
const FExtXform = require('./f-ext-xform');

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
      'x14:dxf': (this.dxfExtXform = new DxfExtXform()),
      'xm:f': (this.fXform = new FExtXform()),
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
    if (rule.type === 'expression') {
      if (rule.dxf) {
        return true;
      }
    }
    return false;
  }

  prepare(model) {
    if (CfRuleExtXform.isExt(model)) {
      model.x14Id = `{${uuidv4()}}`.toUpperCase();
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
      case 'expression':
        this.renderExpression(xmlStream, model);
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
      id: model.x14Id || `{${uuidv4()}}`,
    });

    this.iconSetXform.render(xmlStream, model);

    xmlStream.closeNode();
  }

  renderExpression(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      type: 'expression',
      priority: model.priority,
      id: model.x14Id || `{${uuidv4()}}`,
    });

    this.fXform.render(xmlStream, model.formulae[0]);

    if (model.dxf) {
      this.dxfExtXform.render(xmlStream, model.dxf);
    }

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
    switch (name) {
      case 'xm:f':
        this.model.formulae = this.model.formulae || [];
        this.model.formulae.push(parser.model);
        break;

      case 'x14:dxf':
        this.model.dxf = parser.model;
        break;

      default:
        Object.assign(this.model, parser.model);
        break;
    }
  }
}

module.exports = CfRuleExtXform;
