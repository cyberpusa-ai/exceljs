const fs = require('fs');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

// CT_Worksheet child order from the OOXML schema (subset that the stream writer emits)
const ELEMENT_ORDER = [
  'sheetData',
  'sheetProtection',
  'autoFilter',
  'mergeCells',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'legacyDrawing',
  'picture',
];

describe('github issues', () => {
  describe('pr 2685 - streaming sheet with autoFilter and sheet protection', () => {
    it('writes worksheet elements in schema order', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
        useStyles: true,
        useSharedStrings: true,
      });
      const imageId = wb.addImage({
        buffer: fs.readFileSync(`${__dirname}/../data/image.png`),
        extension: 'png',
      });
      const ws = wb.addWorksheet('Sheet1');
      ws.autoFilter = 'A1:B1';
      await ws.protect('secret', {autoFilter: true});
      const header = ws.addRow(['Name', 'Link']);
      header.addPageBreak();
      header.commit();
      ws.addRow(['foo', {text: 'link', hyperlink: 'https://example.com'}]).commit();
      ws.getCell('A3').value = 'note';
      ws.getCell('A3').note = 'a comment';
      ws.mergeCells('A4:B4');
      ws.dataValidations.add('A5', {type: 'list', allowBlank: true, formulae: ['"a,b"']});
      ws.addConditionalFormatting({
        ref: 'A1:A5',
        rules: [{type: 'expression', formulae: ['TRUE'], style: {font: {bold: true}}}],
      });
      ws.headerFooter.oddHeader = 'header';
      ws.addBackgroundImage(imageId);
      ws.commit();
      await wb.commit();

      const zip = await JSZip.loadAsync(fs.readFileSync(TEST_XLSX_FILE_NAME));
      const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
      const positions = ELEMENT_ORDER.map(tag => {
        const index = xml.search(new RegExp(`<${tag}[ />]`));
        expect(index, `<${tag}> present`).to.be.at.least(0);
        return index;
      });
      const sorted = [...positions].sort((a, b) => a - b);
      expect(positions).to.deep.equal(sorted);

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      const ws2 = wb2.getWorksheet('Sheet1');
      expect(ws2.autoFilter).to.equal('A1:B1');
      expect(ws2.sheetProtection.sheet).to.equal(true);
      expect(ws2.sheetProtection.autoFilter).to.equal(true);
    });
  });
});
