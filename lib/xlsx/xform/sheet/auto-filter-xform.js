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
          const ref = `${firstAddress}:${secondAddress}`;
          if (model.exclude && model.exclude.length) {
            // hide the filter button on the excluded columns (0 = leftmost column of the range)
            xmlStream.openNode('autoFilter', {ref});
            model.exclude.forEach(colId => {
              xmlStream.leafNode('filterColumn', {colId, hiddenButton: '1'});
            });
            xmlStream.closeNode();
          } else {
            xmlStream.leafNode('autoFilter', {ref});
          }
        }
      }
    }
  }

  parseOpen(node) {
    switch (node.name) {
      case 'autoFilter':
        this.ref = node.attributes.ref;
        this.exclude = [];
        this.model = this.ref;
        return true;
      case 'filterColumn': {
        const {colId, hiddenButton, showButton} = node.attributes;
        const hidden =
          hiddenButton === '1' ||
          hiddenButton === 'true' ||
          showButton === '0' ||
          showButton === 'false';
        if (hidden && colId !== undefined) {
          this.exclude.push(parseInt(colId, 10));
        }
        return true;
      }
      default:
        return true;
    }
  }

  parseClose(name) {
    if (name === 'autoFilter') {
      if (this.exclude.length && this.ref) {
        const [from, to] = this.ref.split(':');
        this.model = {from, to: to || from, exclude: this.exclude};
      }
      return false;
    }
    return true;
  }
}

module.exports = AutoFilterXform;
