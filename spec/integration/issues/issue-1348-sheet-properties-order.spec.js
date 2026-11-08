const JSZip = require('jszip');
const {PassThrough} = require('readable-stream');

const ExcelJS = verquire('exceljs');

// CT_SheetPr requires the child order: tabColor, outlinePr, pageSetUpPr
function expectSheetPrOrder(xml) {
  const match = xml.match(/<sheetPr>(.*?)<\/sheetPr>/);
  expect(match).to.be.ok();
  const children = match[1].match(/<(\w+)/g).map(tag => tag.slice(1));
  expect(children).to.deep.equal(['tabColor', 'outlinePr', 'pageSetUpPr']);
}

const sheetOptions = {
  properties: {tabColor: {argb: 'FFFF0000'}, outlineProperties: {summaryBelow: false}},
  pageSetup: {fitToPage: true},
};

describe('github issues', () => {
  describe('issue 1348 - fitToPage with outlineProperties produces a broken xlsx', () => {
    it('Workbook writes sheetPr children in schema order', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet', sheetOptions);
      ws.getCell('A1').value = 1;
      const buffer = await wb.xlsx.writeBuffer();

      const zip = await JSZip.loadAsync(buffer);
      expectSheetPrOrder(await zip.file('xl/worksheets/sheet1.xml').async('string'));

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('sheet');
      expect(ws2.properties.outlineProperties.summaryBelow).to.equal(false);
      expect(ws2.pageSetup.fitToPage).to.equal(true);
    });

    it('WorkbookWriter writes sheetPr children in schema order', async () => {
      const stream = new PassThrough();
      const chunks = [];
      stream.on('data', chunk => chunks.push(chunk));
      const done = new Promise(resolve => stream.on('end', resolve));

      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
      const ws = wb.addWorksheet('sheet', sheetOptions);
      ws.getCell('A1').value = 1;
      ws.commit();
      await wb.commit();
      await done;

      const zip = await JSZip.loadAsync(Buffer.concat(chunks));
      expectSheetPrOrder(await zip.file('xl/worksheets/sheet1.xml').async('string'));
    });
  });
});
