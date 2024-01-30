const colCache = require('../../../utils/col-cache');
const BaseXform = require('../base-xform');

class AutoFilterXform extends BaseXform {
  get tag() {
    return 'autoFilter';
  }

  render(xmlStream, model) {
    if (model) {
      if (typeof model === 'string') {
        // assume range
        xmlStream.leafNode('autoFilter', {ref: model});
      } else {
        const getAddress = function(addr) {
          if (typeof addr === 'string') {
            return addr;
          }
          return colCache.getAddress(addr.row, addr.column).address;
        };

        const firstAddress = getAddress(model.from);
        const secondAddress = getAddress(model.to);
        if (firstAddress && secondAddress) {
          if (model.exclude) {
            xmlStream.openNode('autoFilter');
            xmlStream.addAttribute('ref', `${firstAddress}:${secondAddress}`); 
            for (let i = 0; i < model.exclude.length; i++) {
              xmlStream.leafNode('filterColumn', {colId: model.exclude[i], showButton: '0'}); 
            }
            xmlStream.closeNode(); 
          } else {
            xmlStream.leafNode('autoFilter', {ref: `${firstAddress}:${secondAddress}`});
          }
        }
      }
    }
  }

  parseOpen(node) {
    if (node.name === 'autoFilter') {
      this.model = node.attributes.ref;
    }
  }
}

module.exports = AutoFilterXform;
