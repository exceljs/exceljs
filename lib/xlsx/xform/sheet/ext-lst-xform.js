const BaseXform = require('../base-xform');

const ConditionalFormattingsExt = require('./cf/conditional-formattings-ext-xform');

class ExtLstXform extends BaseXform {
  constructor() {
    super();
    this.map = {
      'x14:conditionalFormattings': new ConditionalFormattingsExt(),
    };
  }

  get tag() {
    return 'extLst';
  }

  render(xmlStream, model) {
    let hasContent = false;
    xmlStream.addRollback();
    xmlStream.openNode('extLst');

    // conditional formatting
    xmlStream.openNode('ext', {
      uri: '{78C0D931-6437-407d-A8EE-F0AAD7539E65}',
      'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    });
    const cfCursor = xmlStream.cursor;
    this.map['x14:conditionalFormattings'].render(model.conditionalFormattings);
    hasContent = hasContent || (cfCursor !== xmlStream.cursor);
    xmlStream.closeNode();

    xmlStream.closeNode();
    if (hasContent) {
      xmlStream.commit();
    } else {
      xmlStream.rollback();
    }
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case 'extLst':
        this.model = {};
        return true;
      case 'ext':
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }

    return (name !== 'extList');
  }
}

module.exports = ExtLstXform;
