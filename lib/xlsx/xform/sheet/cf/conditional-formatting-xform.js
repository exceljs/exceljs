const BaseXform = require('../../base-xform');

const CfRuleXform = require('./cf-rule-xform');

class ConditionalFormattingXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      cfRule: new CfRuleXform(),
    };
  }

  get tag() {
    return 'conditionalFormatting';
  }

  render(xmlStream, model) {
    // if there are no primitive rules, exit now
    if (!model.rules.some(CfRuleXform.isPrimitive)) {
      return;
    }

    xmlStream.openNode(this.tag, {sqref: model.ref});

    model.rules.forEach(rule => {
      if (CfRuleXform.isPrimitive(rule)) {
        rule.ref = model.ref;
        this.map.cfRule.render(xmlStream, rule);
      }
    });

    xmlStream.closeNode();
  }

  parseOpen(node) {
    this.parser = this.parser || this.map[node.name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    if (node.name === this.tag) {
      this.model = {
        ref: node.attributes.sqref,
        rules: [],
      };
      return true;
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
        this.model.rules.push(this.parser.model);
        this.parser = null;
      }
      return true;
    }
    return name !== this.tag;
  }
}

module.exports = ConditionalFormattingXform;
