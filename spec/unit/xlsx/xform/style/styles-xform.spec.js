const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const StylesXform = verquire('xlsx/xform/style/styles-xform');
const XmlStream = verquire('utils/xml-stream');

const expectations = [
  {
    title: 'Styles with fonts',
    create() {
      return new StylesXform();
    },
    preparedModel: require('./data/styles.1.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/styles.1.2.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('StylesXform', () => {
  testXformHelper(expectations);

  describe('As StyleManager', () => {
    it('Renders empty model', () => {
      const stylesXform = new StylesXform(true);
      const expectedXml = fs
        .readFileSync(`${__dirname}/data/styles.2.2.xml`)
        .toString();

      const xmlStream = new XmlStream();
      stylesXform.render(xmlStream);

      expect(xmlStream.xml).xml.to.equal(expectedXml);
    });

    it('Adds a dxf with a numFmt string', () => {
      const stylesXform = new StylesXform(true);
      const style = {numFmt: '0.00%'};
      const dxfId = stylesXform.addDxfStyle(style);

      const dxf = stylesXform.getDxfStyle(dxfId);
      expect(dxf.numFmt).to.equal('0.00%');
      expect(dxf.numFmtId).to.be.a('number');
      expect(style.numFmtId).to.equal(dxf.numFmtId);
    });

    it('Adds a dxf with a parsed numFmt model (#2698)', () => {
      const stylesXform = new StylesXform(true);
      const style = {font: {bold: true}, numFmt: {id: 10, formatCode: '0.00%'}};
      const dxfId = stylesXform.addDxfStyle(style);

      const dxf = stylesXform.getDxfStyle(dxfId);
      expect(dxf.numFmt).to.equal('0.00%');
      expect(dxf.numFmtId).to.be.a('number');
      expect(dxf.font).to.deep.equal({bold: true});
      // the caller's model keeps its parsed shape
      expect(style.numFmt).to.deep.equal({id: 10, formatCode: '0.00%'});

      const xmlStream = new XmlStream();
      stylesXform.render(xmlStream);
      expect(xmlStream.xml).to.not.include('[object Object]');
      expect(xmlStream.xml).to.include(
        `<dxf><font><b/></font><numFmt numFmtId="${dxf.numFmtId}" formatCode="0.00%"/></dxf>`
      );
    });
  });
});
