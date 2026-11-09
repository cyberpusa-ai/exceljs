const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

async function getStylesXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('xl/styles.xml').async('string');
}

describe('github issues', () => {
  it('issue 2698 - conditional format numFmt survives a read/write round trip', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 0.5;
    ws.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'cellIs',
          operator: 'greaterThan',
          formulae: [0],
          style: {numFmt: '0.00%'},
        },
      ],
    });

    const buffer1 = await wb.xlsx.writeBuffer();
    expect(await getStylesXml(buffer1)).to.include('formatCode="0.00%"');

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer1);
    const rule = wb2.getWorksheet('Sheet1').conditionalFormattings[0].rules[0];
    expect(rule.style.numFmt.formatCode).to.equal('0.00%');

    const buffer2 = await wb2.xlsx.writeBuffer();
    // writing must not change the loaded model
    expect(rule.style.numFmt.formatCode).to.equal('0.00%');
    const stylesXml = await getStylesXml(buffer2);
    expect(stylesXml).to.not.include('[object Object]');
    expect(stylesXml).to.match(/<dxf>[^]*<numFmt numFmtId="\d+" formatCode="0.00%"\/>[^]*<\/dxf>/);

    const wb3 = new ExcelJS.Workbook();
    await wb3.xlsx.load(buffer2);
    const rule3 = wb3.getWorksheet('Sheet1').conditionalFormattings[0].rules[0];
    expect(rule3.style.numFmt.formatCode).to.equal('0.00%');
  });
});
