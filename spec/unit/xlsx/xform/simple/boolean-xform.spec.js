const testXformHelper = require('../test-xform-helper');

const BooleanXform = verquire('xlsx/xform/simple/boolean-xform');

const expectations = [
  {
    title: 'true',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<boolean/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'val 0',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: false,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<boolean val="0"/>',
    tests: ['parse'],
  },
  {
    title: 'val false',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: false,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<boolean val="false"/>',
    tests: ['parse'],
  },
  {
    title: 'val 1',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<boolean val="1"/>',
    tests: ['parse'],
  },
  {
    title: 'val true',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<boolean val="true"/>',
    tests: ['parse'],
  },
  {
    title: 'false',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: false,
    xml: '',
    tests: ['render', 'renderIn'],
  },
  {
    title: 'undefined',
    create() {
      return new BooleanXform({tag: 'boolean', attr: 'val'});
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('BooleanXform', () => {
  testXformHelper(expectations);
});
