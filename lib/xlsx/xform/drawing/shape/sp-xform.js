const BaseXform = require('../../base-xform');
const StaticXform = require('../../static-xform');

const NvSpPrXform = require('./nv-sp-pr-xform');
const SpPrXform = require('./sp-pr-xform');

// DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape
class SpXform extends BaseXform {
  constructor() {
    super();

    this.map = {
      'xdr:nvSpPr': new NvSpPrXform(),
      'xdr:spPr': new SpPrXform(),
      'xdr:style': new StaticXform(styleJSON),
    };
  }

  get tag() {
    return 'xdr:sp';
  }

  prepare(model, options) {
    model.index = options.index + 1;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {macro: '', textlink: ''});

    this.map['xdr:nvSpPr'].render(xmlStream, model);
    this.map['xdr:spPr'].render(xmlStream, model);
    this.map['xdr:style'].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
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
    switch (name) {
      case this.tag:
        this.mergeModel({
          ...this.map['xdr:spPr'].model,
        });
        return false;
      default:
        return true;
    }
  }
}

const shadeJSON = {
  tag: 'a:shade',
  $: {val: '15000'},
};

const lnRefJSON = {
  tag: 'a:lnRef',
  $: {idx: '2'},
  c: [
    {
      tag: 'a:schemeClr',
      $: {val: 'accent1'},
      c: [shadeJSON],
    },
  ],
};

const fillRefJSON = {
  tag: 'a:fillRef',
  $: {idx: '1'},
  c: [
    {
      tag: 'a:schemeClr',
      $: {val: 'accent1'},
    },
  ],
};

const effectRefJSON = {
  tag: 'a:effectRef',
  $: {idx: '0'},
  c: [
    {
      tag: 'a:schemeClr',
      $: {val: 'accent1'},
    },
  ],
};

const fontRefJSON = {
  tag: 'a:fontRef',
  $: {idx: 'minor'},
  c: [
    {
      tag: 'a:schemeClr',
      $: {val: 'lt1'},
    },
  ],
};

const styleJSON = {
  tag: 'xdr:style',
  c: [lnRefJSON, fillRefJSON, effectRefJSON, fontRefJSON],
};
// <xdr:txBody>
// <a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/>
// <a:lstStyle/>
// <a:p>
//   <a:pPr algn="l"/>
//   <a:endParaRPr lang="en-US" sz="1100"/>
// </a:p>
// </xdr:txBody>

module.exports = SpXform;
