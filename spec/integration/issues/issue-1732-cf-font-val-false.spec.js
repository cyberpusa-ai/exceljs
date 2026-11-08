const {PassThrough} = require('readable-stream');

const ExcelJS = verquire('exceljs');
const StylesXform = verquire('xlsx/xform/style/styles-xform');

const XML_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><b val="0"/><i val="false"/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="0"/>
  <borders count="0"/>
  <cellStyleXfs count="0"/>
  <cellXfs count="0"/>
  <dxfs count="1">
    <dxf>
      <font><b val="0"/><i val="0"/><strike val="0"/><u val="none"/><color rgb="FF9C0006"/></font>
    </dxf>
  </dxfs>
</styleSheet>`;

function parse(xml) {
  const xform = new StylesXform();
  const stream = new PassThrough();
  stream.write(xml);
  stream.end();
  return xform.parseStream(stream).then(() => xform.model);
}

describe('github issues', () => {
  it('issue 1732 - <b val="0"/> in fonts and dxfs is read as false, not true', async () => {
    const model = await parse(XML_STYLES);
    const [font] = model.fonts;
    expect(font.bold).to.not.be.ok();
    expect(font.italic).to.not.be.ok();

    const [dxf] = model.dxfs;
    expect(dxf.font.bold).to.not.be.ok();
    expect(dxf.font.italic).to.not.be.ok();
    expect(dxf.font.strike).to.not.be.ok();
    expect(dxf.font.color).to.deep.equal({argb: 'FF9C0006'});
  });

  it('issue 1732 - conditional formatting font round-trips without gaining bold/italic', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'expression',
          priority: 1,
          formulae: ['A1>5'],
          style: {font: {bold: true, color: {argb: 'FF9C0006'}}},
        },
      ],
    });
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(await wb.xlsx.writeBuffer());
    const [cf] = wb2.getWorksheet('Sheet1').conditionalFormattings;
    const {font} = cf.rules[0].style;
    expect(font.bold).to.equal(true);
    expect(font.italic).to.not.be.ok();
    expect(font.strike).to.not.be.ok();
  });
});
