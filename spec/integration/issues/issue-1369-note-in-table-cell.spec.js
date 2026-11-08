const JSZip = require('jszip');
const {PassThrough} = require('readable-stream');

const ExcelJS = verquire('exceljs');

// CT_Worksheet requires: ..., drawing, legacyDrawing, ..., picture, ..., tableParts, extLst
const SCHEMA_ORDER = ['drawing', 'legacyDrawing', 'picture', 'tableParts', 'extLst'];

function expectWorksheetTailOrder(xml, expected) {
  // cfRule elements carry their own nested extLst; only look at worksheet children
  const top = xml.replace(/<conditionalFormatting\b.*?<\/conditionalFormatting>/g, '');
  const tags = (top.match(/<(drawing|legacyDrawing|picture|tableParts|extLst)\b/g) || []).map(tag =>
    tag.slice(1)
  );
  expect(tags).to.deep.equal(expected);
}

const image = {
  base64:
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==',
  extension: 'png',
};

describe('github issues', () => {
  describe('issue 1369 - Unable to specify a note in a cell within a table', () => {
    it('Workbook writes legacyDrawing before picture, tableParts and extLst', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      ws.addTable({
        name: 'MyTable',
        ref: 'A1',
        headerRow: true,
        columns: [{name: 'Name'}, {name: 'Value'}],
        rows: [
          ['a', 1],
          ['b', 2],
        ],
      });
      ws.getCell('A2').note = 'note in a table cell';
      ws.addConditionalFormatting({
        ref: 'B2:B3',
        rules: [{type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}]}],
      });
      const imageId = wb.addImage(image);
      ws.addImage(imageId, 'D1:E2');
      ws.addBackgroundImage(imageId);

      const buffer = await wb.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      expectWorksheetTailOrder(
        await zip.file('xl/worksheets/sheet1.xml').async('string'),
        SCHEMA_ORDER
      );

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('sheet');
      expect(ws2.getCell('A2').note).to.equal('note in a table cell');
      expect(ws2.getTable('MyTable')).to.be.ok();
    });

    it('WorkbookWriter writes legacyDrawing before picture', async () => {
      const stream = new PassThrough();
      const chunks = [];
      stream.on('data', chunk => chunks.push(chunk));
      const done = new Promise(resolve => stream.on('end', resolve));

      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
      const ws = wb.addWorksheet('sheet');
      ws.getCell('A1').value = 1;
      ws.getCell('A1').note = 'a note';
      ws.addBackgroundImage(wb.addImage(image));
      ws.commit();
      await wb.commit();
      await done;

      const zip = await JSZip.loadAsync(Buffer.concat(chunks));
      expectWorksheetTailOrder(await zip.file('xl/worksheets/sheet1.xml').async('string'), [
        'legacyDrawing',
        'picture',
      ]);
    });
  });
});
