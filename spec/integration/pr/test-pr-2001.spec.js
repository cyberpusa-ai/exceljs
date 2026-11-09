const fs = require('fs');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Some third-party exporters write a plain <t> before the <r> runs of a shared string
const SHARED_STRINGS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
  '<si><t></t><r><rPr><b/></rPr><t>Bold</t></r></si>' +
  '<si><t>Plain </t><r><rPr><b/></rPr><t>Bold</t></r></si>' +
  '</sst>';

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

async function createBuffer() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'first';
  ws.getCell('A2').value = 'second';
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  zip.file('xl/sharedStrings.xml', SHARED_STRINGS);
  return zip.generateAsync({type: 'nodebuffer'});
}

describe('github issues', () => {
  describe('pr 2001 - shared string with plain text before rich text runs', () => {
    it('reads with the document reader', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(await createBuffer());
      const ws = wb.getWorksheet('Sheet1');
      expect(ws.getCell('A1').value).to.deep.equal({
        richText: [{font: {bold: true}, text: 'Bold'}],
      });
      expect(ws.getCell('A2').value).to.deep.equal({
        richText: [{text: 'Plain '}, {font: {bold: true}, text: 'Bold'}],
      });
    });

    it('reads with the streaming reader', async () => {
      fs.writeFileSync(TEST_XLSX_FILE_NAME, await createBuffer());
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(TEST_XLSX_FILE_NAME, {
        sharedStrings: 'cache',
      });
      const values = [];
      for await (const ws of reader) {
        for await (const row of ws) {
          values.push(row.getCell(1).value);
        }
      }
      expect(values).to.deep.equal([
        {richText: [{font: {bold: true}, text: 'Bold'}]},
        {richText: [{font: null, text: 'Plain '}, {font: {bold: true}, text: 'Bold'}]},
      ]);
    });
  });
});
