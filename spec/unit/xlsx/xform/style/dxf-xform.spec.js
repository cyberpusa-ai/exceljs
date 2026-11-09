const testXformHelper = require('../test-xform-helper');

const DxfXform = verquire('xlsx/xform/style/dxf-xform');

const expectations = [
  {
    title: 'Font',
    create: () => new DxfXform(),
    preparedModel: {font: {bold: true}},
    xml: '<dxf><font><b/></font></dxf>',
    parsedModel: {
      alignment: null,
      border: null,
      fill: null,
      font: {bold: true},
      numFmt: null,
      protection: null,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'NumFmt',
    create: () => new DxfXform(),
    preparedModel: {numFmt: '0.00%', numFmtId: 164},
    xml: '<dxf><numFmt numFmtId="164" formatCode="0.00%"/></dxf>',
    // parsed numFmt keeps the NumFmtXform model shape
    parsedModel: {
      alignment: null,
      border: null,
      fill: null,
      font: null,
      numFmt: {id: 164, formatCode: '0.00%'},
      protection: null,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Font and NumFmt',
    create: () => new DxfXform(),
    preparedModel: {font: {italic: true}, numFmt: '#,##0.000', numFmtId: 165},
    xml: '<dxf><font><i/></font><numFmt numFmtId="165" formatCode="#,##0.000"/></dxf>',
    parsedModel: {
      alignment: null,
      border: null,
      fill: null,
      font: {italic: true},
      numFmt: {id: 165, formatCode: '#,##0.000'},
      protection: null,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('DxfXform', () => {
  testXformHelper(expectations);
});
