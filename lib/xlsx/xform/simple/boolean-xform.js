const BaseXform = require('../base-xform');

class BooleanXform extends BaseXform {
  constructor(options) {
    super();

    this.tag = options.tag;
    this.attr = options.attr;
  }

  render(xmlStream, model) {
    if (model) {
      xmlStream.openNode(this.tag);
      xmlStream.closeNode();
    }
  }

  parseOpen(node) {
    if (node.name === this.tag) {
      // xsd:boolean - a missing attribute means true, "0" or "false" means false
      const value = this.attr && node.attributes[this.attr];
      this.model = value !== '0' && value !== 'false';
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = BooleanXform;
