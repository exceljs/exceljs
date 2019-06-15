const BaseXform = require('../base-xform');


class HeaderFooterXform extends BaseXform {
  get tag() {
    return 'headerFooter';
  }

  render(xmlStream, model) {
    if (model) {
      xmlStream.openNode('headerFooter');
      if (model.differentFirst) {
        xmlStream.addAttribute('differentFirst', '1');
      }
      if (model.differentOddEven) {
        xmlStream.addAttribute('differentOddEven', '1');
      }

      if (model.oddHeader !== null) {
        xmlStream.leafNode('oddHeader', null, model.oddHeader);
      }
      if (model.oddFooter !== null) {
        xmlStream.leafNode('oddFooter', null, model.oddFooter);
      }
      if (model.evenHeader !== null) {
        xmlStream.leafNode('evenHeader', null, model.evenHeader);
      }
      if (model.evenFooter !== null) {
        xmlStream.leafNode('evenFooter', null, model.evenFooter);
      }
      if (model.firstHeader !== null) {
        xmlStream.leafNode('firstHeader', null, model.firstHeader);
      }
      if (model.firstFooter !== null) {
        xmlStream.leafNode('firstFooter', null, model.firstFooter);
      }

      xmlStream.closeNode();
    }
  }
}

module.exports = HeaderFooterXform;