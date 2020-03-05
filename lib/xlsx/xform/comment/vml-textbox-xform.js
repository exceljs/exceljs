const BaseXform = require('../base-xform');
const _ = require('../../../utils/under-dash');

class VmlTextboxXform extends BaseXform {
  get tag() {
    return 'v:textbox';
  }

  conversionUnit(value, multiple, unit) {
    return `${parseFloat(value) * multiple.toFixed(2)}${unit}`;
  }

  reverseConversionUnit(margins) {
    return (margins || '').split(',').map(margin => {
      return Number(parseFloat(this.conversionUnit(parseFloat(margin), 0.1, '')).toFixed(2));
    });
  }

  render(xmlStream, model) {
    const attributes = {
      style: 'mso-direction-alt:auto',
    };
    if (model && model.note) {
      let {margins} = model.note;
      if (_.isArray(margins)) {
        margins = margins
          .map(margin => {
            return this.conversionUnit(margin, 10, 'mm');
          })
          .join(',');
      }
      if (margins) {
        attributes.inset = margins;
      }
    }
    xmlStream.openNode('v:textbox', attributes);
    xmlStream.leafNode('div', {style: 'text-align:left'});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          margins: this.reverseConversionUnit(node.attributes.inset),
        };
        return true;
      default:
        return true;
    }
  }

  parseText() {}

  parseClose(name) {
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

module.exports = VmlTextboxXform;
