const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/test-issue-1751.xlsx';

function checkFormattings(ws) {
  const cfs = ws.conditionalFormattings;
  expect(cfs).to.have.lengthOf(2);
  const [cfB, cfC] = cfs;

  expect(cfB.ref).to.equal('B1');
  expect(cfB.rules).to.have.lengthOf(1);
  expect(cfB.rules[0].type).to.equal('expression');
  expect(cfB.rules[0].formulae).to.deep.equal(['A1=1']);
  expect(cfB.rules[0].style.fill.bgColor).to.deep.equal({argb: 'FF00A933'});

  expect(cfC.ref).to.equal('C1');
  expect(cfC.rules).to.have.lengthOf(1);
  expect(cfC.rules[0].type).to.equal('expression');
  expect(cfC.rules[0].formulae).to.deep.equal(['A1<>0']);
  expect(cfC.rules[0].style.fill.bgColor).to.deep.equal({argb: 'FFFF0000'});
}

describe('github issues', () => {
  it('issue 1751 - xm:f in x14:cfRule', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/test-issue-1751.xlsx');

    checkFormattings(wb.worksheets[0]);

    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

    const wbNew = new ExcelJS.Workbook();
    await wbNew.xlsx.readFile(TEST_XLSX_FILE_NAME);

    checkFormattings(wbNew.worksheets[0]);
  }).timeout(6000);

  it('issue 1751 - x14 expression rules are written back to extLst only', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./spec/integration/data/test-issue-1751.xlsx');
    // a regular expression rule stays in the main <conditionalFormatting>
    wb.worksheets[0].addConditionalFormatting({
      ref: 'A1',
      rules: [{type: 'expression', formulae: ['A1>0'], style: {font: {bold: true}}}],
    });
    const buffer = await wb.xlsx.writeBuffer();

    const zip = await JSZip.loadAsync(buffer);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');
    const stylesXml = await zip.file('xl/styles.xml').async('string');

    const main = sheetXml.slice(0, sheetXml.indexOf('<extLst>'));
    const ext = sheetXml.slice(sheetXml.indexOf('<extLst>'));
    expect(main.match(/<cfRule /g)).to.have.lengthOf(1);
    expect(main).to.include('<formula>A1&gt;0</formula>');
    expect(ext.match(/<x14:cfRule type="expression"/g)).to.have.lengthOf(2);
    expect(ext).to.include('<xm:f>A1=1</xm:f>');
    expect(ext).to.include('<xm:f>A1&lt;&gt;0</xm:f>');
    expect(ext).to.include('<x14:dxf><fill><patternFill><bgColor rgb="FF00A933"/>');
    // only the regular rule needs a dxf in styles.xml
    expect(stylesXml).to.include('<dxfs count="1">');
  }).timeout(6000);
});
