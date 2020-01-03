/* eslint-disable max-classes-per-file */
const BaseXform = require('../../base-xform');

class X14IdXform extends BaseXform {
  get tag() {
    return 'x14:id';
  }

  render(xmlStream, model) {
    xmlStream.leafNode(this.tag, null, model);
  }

  parseOpen() {
    this.model = '';
  }

  parseText(text) {
    this.model += text;
  }

  parseClose(name) {
    return name !== this.tag;
  }
}

class ExtXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'x14:id': new X14IdXform(),
    };
  }

  get tag() {
    return 'ext';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      uri: '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}',
      'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    });

    this.map['x14:id'].render(model.x14Id);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    switch (node.name) {
      case this.tag:
        this.model = {};
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

    return name !== this.tag;
  }
}

class ExtLstRefXform extends BaseXform {
  constructor() {
    super();
    this.map = {
      ext: new ExtXform(),
    };
  }

  get tag() {
    return 'extLst';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);
    this.map.ext.render(xmlStream, model);
    xmlStream.closeNode();
  }

  parseOpen(node) {
    this.model = {
      type: node.attributes.type,
      value: node.attributes.val,
    };
  }

  parseClose(name) {
    return name !== this.tag;
  }
}

module.exports = ExtLstRefXform;
