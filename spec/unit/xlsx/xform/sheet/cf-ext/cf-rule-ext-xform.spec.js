const testXformHelper = require('../../test-xform-helper');

const CfRuleExtXform = verquire('xlsx/xform/sheet/cf-ext/cf-rule-ext-xform');

const expectations = [
  {
    title: 'Icon Set',
    create() {
      return new CfRuleExtXform();
    },
    preparedModel: {
      type: 'iconSet',
      priority: 1,
      x14Id: 'x14-id',
      iconSet: '3Triangles',
      cfvo: [
        {type: 'percent', value: 0},
        {type: 'percent', value: 33},
        {type: 'percent', value: 67},
      ],
    },
    xml: `
      <x14:cfRule type="iconSet" priority="1" id="x14-id">
        <x14:iconSet iconSet="3Triangles">
          <x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>
          <x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>
          <x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>
        </x14:iconSet>
      </x14:cfRule>
    `,
    parsedModel: {
      type: 'iconSet',
      priority: 1,
      x14Id: 'x14-id',
      iconSet: '3Triangles',
      reverse: false,
      showValue: true,
      cfvo: [
        {type: 'percent', value: 0},
        {type: 'percent', value: 33},
        {type: 'percent', value: 67},
      ],
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Databar',
    create() {
      return new CfRuleExtXform();
    },
    preparedModel: {
      type: 'dataBar',
      x14Id: 'x14-id',
      cfvo: [
        {type: 'num', value: 5},
        {type: 'num', value: 20},
      ],
    },
    xml: `
      <x14:cfRule type="dataBar" id="x14-id">
        <x14:dataBar minLength="0" maxLength="100">
          <x14:cfvo type="num"><xm:f>5</xm:f></x14:cfvo>
          <x14:cfvo type="num"><xm:f>20</xm:f></x14:cfvo>
        </x14:dataBar>
      </x14:cfRule>
    `,
    parsedModel: {
      type: 'dataBar',
      x14Id: 'x14-id',
      minLength: 0,
      maxLength: 100,
      border: false,
      gradient: true,
      // reverse: false,
      // showValue: true,
      negativeBarColorSameAsPositive: true,
      negativeBarBorderColorSameAsPositive: true,
      axisPosition: 'auto',
      direction: 'leftToRight',
      cfvo: [
        {type: 'num', value: 5},
        {type: 'num', value: 20},
      ],
    },
    tests: ['render', 'parse'],
  },
  {
    title: 'Expression',
    create() {
      return new CfRuleExtXform();
    },
    preparedModel: {
      type: 'expression',
      x14Id: 'x14-id',
      priority: 1,
      formulae: ['Sheet2!$A$1=1'],
      style: {
        fill: {type: 'pattern', bgColor: {argb: 'FFFF0000'}},
        numFmt: {id: 164, formatCode: '0.0'},
      },
    },
    xml: `
      <x14:cfRule type="expression" priority="1" id="x14-id">
        <xm:f>Sheet2!$A$1=1</xm:f>
        <x14:dxf>
          <numFmt numFmtId="164" formatCode="0.0"/>
          <fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill>
        </x14:dxf>
      </x14:cfRule>
    `,
    parsedModel: {
      type: 'expression',
      x14Id: 'x14-id',
      priority: 1,
      formulae: ['Sheet2!$A$1=1'],
      style: {
        alignment: null,
        border: null,
        fill: {type: 'pattern', bgColor: {argb: 'FFFF0000'}},
        font: null,
        numFmt: {id: 164, formatCode: '0.0'},
        protection: null,
      },
    },
    tests: ['render', 'parse'],
  },
];

describe('CfRuleExtXform', () => {
  testXformHelper(expectations);
});

const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;

describe('CfRuleExtXform prepare', () => {
  it('assigns a GUID x14Id to ext rules', () => {
    const model = {type: 'iconSet', iconSet: '3Triangles'};
    new CfRuleExtXform().prepare(model);
    expect(model.x14Id).to.match(GUID);
  });

  it('does not assign an x14Id to non-ext rules', () => {
    const model = {type: 'iconSet', iconSet: '3Arrows'};
    new CfRuleExtXform().prepare(model);
    expect(model.x14Id).to.be.undefined();
  });

  it('only treats expression rules read from x14:cfRule as ext', () => {
    expect(CfRuleExtXform.isExt({type: 'expression', formulae: ['A1=1']})).to.be.false();
    expect(CfRuleExtXform.isExt({type: 'expression', x14Id: '{ID}', formulae: ['A1=1']})).to.be.true();
  });
});
