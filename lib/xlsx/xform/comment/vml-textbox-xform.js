const BaseXform = require('../base-xform');

class VmlTextboxXform extends BaseXform {
  get tag() {
    return 'v:textbox';
  }

  conversionUnit(value, multiple, unit) {
    return `${parseFloat(value) * multiple.toFixed(2)}${unit}`;
  }

  reverseConversionUnit(inset) {
    return (inset || '').split(',').map(margin => {
      return Number(parseFloat(this.conversionUnit(parseFloat(margin), 0.1, '')).toFixed(2));
    });
  }

  render(xmlStream, model) {
    const styleList = ['mso-direction-alt:auto'];
    const attributes = {};
    if (model && model.note) {
      let {inset} = model.note && model.note.margins;
      if (Array.isArray(inset)) {
        inset = inset
          .map(margin => {
            return this.conversionUnit(margin, 10, 'mm');
          })
          .join(',');
      }
      if (inset) {
        attributes.inset = inset;
      }
      if (model.note.autoShape) {
        styleList.push('mso-fit-shape-to-text:t');
      }
    }
    attributes.style = styleList.join(';');
    xmlStream.openNode('v:textbox', attributes);
    xmlStream.leafNode('div', {style: 'text-align:left'});
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {
          inset: this.reverseConversionUnit(node.attributes.inset),
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
