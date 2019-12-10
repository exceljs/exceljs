const BaseXform = require('../../base-xform');

const ConditionalFormattingExtXform = require('./conditional-formatting-ext-xform');

/* eslint-disable */

class ConditionalFormattingsExtXform extends BaseXform {
  constructor() {
    super();

    this.cfXform = new ConditionalFormattingExtXform();
  }

  get tag() {
    return 'x14:conditionalFormattings';
  }

  render(xmlStream, model) {
    // TBD

  }

  parseOpen(node) {
    // TBD
  }

  parseText(text) {
  }

  parseClose(name) {
  }
}

module.exports = ConditionalFormattingsExtXform;
