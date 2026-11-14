const testXformHelper = require('../test-xform-helper');

const AutoFilterXform = verquire('xlsx/xform/sheet/auto-filter-xform');

const expectations = [
  {
    title: 'Range',
    create() {
      return new AutoFilterXform();
    },
    preparedModel: 'A1:C1',
    xml: '<autoFilter ref="A1:C1"/>',
    parsedModel: 'A1:C1',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Row and Column Address',
    create() {
      return new AutoFilterXform();
    },
    preparedModel: {from: {row: 1, column: 1}, to: {row: 1, column: 3}},
    xml: '<autoFilter ref="A1:C1"/>',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'String address',
    create() {
      return new AutoFilterXform();
    },
    preparedModel: {from: 'A1', to: 'C1'},
    xml: '<autoFilter ref="A1:C1"/>',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'Excluded columns',
    create() {
      return new AutoFilterXform();
    },
    preparedModel: {from: 'A1', to: 'E1', exclude: [0, 2]},
    xml:
      '<autoFilter ref="A1:E1"><filterColumn colId="0" hiddenButton="1"/><filterColumn colId="2" hiddenButton="1"/></autoFilter>',
    parsedModel: {from: 'A1', to: 'E1', exclude: [0, 2]},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Empty exclude',
    create() {
      return new AutoFilterXform();
    },
    preparedModel: {from: {row: 1, column: 1}, to: {row: 1, column: 3}, exclude: []},
    xml: '<autoFilter ref="A1:C1"/>',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'Filter columns with showButton and filters',
    create() {
      return new AutoFilterXform();
    },
    xml:
      '<autoFilter ref="B2:D9"><filterColumn colId="1" showButton="0"/><filterColumn colId="2"><filters><filter val="x"/></filters></filterColumn></autoFilter>',
    parsedModel: {from: 'B2', to: 'D9', exclude: [1]},
    tests: ['parse'],
  },
  {
    title: 'Filter columns without hidden buttons',
    create() {
      return new AutoFilterXform();
    },
    xml:
      '<autoFilter ref="A1:C1"><filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn></autoFilter>',
    parsedModel: 'A1:C1',
    tests: ['parse'],
  },
];

describe('AutoFilterXform', () => {
  testXformHelper(expectations);
});
