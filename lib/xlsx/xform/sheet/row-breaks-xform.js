'use strict';

const BrkXform = require('./brk-xform');

const BaseXform = require('../base-xform');

class RowBreaksXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      brk: new BrkXform(),
    };
  }

  get tag() {
    return 'rowBreaks';
  }

  render(xmlStream, model) {
    if (model && model.length) {
      xmlStream.openNode(this.tag);
      xmlStream.addAttribute('count', model.length);
      xmlStream.addAttribute('manualBreakCount', model.length);

      model.forEach(brk => {
        this.map.brk.render(xmlStream, brk);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = [];
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        if (name === 'brk') {
          this.model.push(this.parser.model);
        }
        this.parser = undefined;
      }
      return true;
    }

    return false;
  }
}

module.exports = RowBreaksXform;
