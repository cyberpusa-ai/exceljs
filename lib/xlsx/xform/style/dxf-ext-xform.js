const DxfXform = require('./dxf-xform');

// <x14:dxf> holds a differential format inline in an x14:cfRule (extLst).
// It has the same children as <dxf> in styles.xml.
class DxfExtXform extends DxfXform {
  get tag() {
    return 'x14:dxf';
  }

  render(xmlStream, model) {
    // a dxf read from a file holds numFmt as a NumFmtXform model ({id, formatCode})
    if (model.numFmt && typeof model.numFmt === 'object') {
      model = {...model, numFmtId: model.numFmt.id, numFmt: model.numFmt.formatCode};
    }
    super.render(xmlStream, model);
  }
}

module.exports = DxfExtXform;
