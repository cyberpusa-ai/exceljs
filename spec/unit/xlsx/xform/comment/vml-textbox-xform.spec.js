const testXformHelper = require('../test-xform-helper');

const VmlTextboxXform = verquire('xlsx/xform/comment/vml-textbox-xform');

const expectations = [
  {
    title: 'default',
    create() {
      return new VmlTextboxXform();
    },
    preparedModel: {note: {margins: {inset: [0.13, 0.13, 0.25, 0.25]}}},
    xml:
      '<v:textbox style="mso-direction-alt:auto" inset="1.3mm,1.3mm,2.5mm,2.5mm"><div style="text-align:left"/></v:textbox>',
    parsedModel: {inset: [0.13, 0.13, 0.25, 0.25]},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'autoShape',
    create() {
      return new VmlTextboxXform();
    },
    preparedModel: {note: {margins: {inset: [0.13, 0.13, 0.25, 0.25]}, autoShape: true}},
    xml:
      '<v:textbox style="mso-direction-alt:auto;mso-fit-shape-to-text:t" inset="1.3mm,1.3mm,2.5mm,2.5mm"><div style="text-align:left"/></v:textbox>',
    parsedModel: {inset: [0.13, 0.13, 0.25, 0.25], autoShape: true},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('VmlTextboxXform', () => {
  testXformHelper(expectations);
});
