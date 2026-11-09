const JSZip = require('jszip');
const {PassThrough} = require('readable-stream');

const ExcelJS = verquire('exceljs');

const FIT_TO_TEXT = 'mso-fit-shape-to-text:t';

describe('github issues', () => {
  describe('issue 981 - Note box size does not fit its content', () => {
    it('Workbook writes autoShape notes as fit-to-text and reads them back', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      ws.getCell('A1').note = {
        texts: [{text: 'This\nIs\nA\nComment\nThat\nNeeds\nA\nLarge\nBox'}],
        autoShape: true,
      };
      ws.getCell('A1').value = 'tall';
      ws.getCell('B1').value = 'short';
      ws.getCell('B1').note = 'plain note';

      const buffer = await wb.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      const vml = await zip.file('xl/drawings/vmlDrawing1.vml').async('string');
      expect(vml.split(FIT_TO_TEXT).length - 1).to.equal(1);

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('sheet');
      expect(ws2.getCell('A1').note.autoShape).to.equal(true);
      expect(ws2.getCell('B1').note).to.equal('plain note');

      const zip2 = await JSZip.loadAsync(await wb2.xlsx.writeBuffer());
      const vml2 = await zip2.file('xl/drawings/vmlDrawing1.vml').async('string');
      expect(vml2.split(FIT_TO_TEXT).length - 1).to.equal(1);
    });

    it('WorkbookWriter writes autoShape notes as fit-to-text', async () => {
      const stream = new PassThrough();
      const chunks = [];
      stream.on('data', chunk => chunks.push(chunk));
      const done = new Promise(resolve => stream.on('end', resolve));

      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
      const ws = wb.addWorksheet('sheet');
      ws.getCell('A1').value = 1;
      ws.getCell('A1').note = {texts: [{text: 'a\nlong\nnote'}], autoShape: true};
      ws.commit();
      await wb.commit();
      await done;

      const zip = await JSZip.loadAsync(Buffer.concat(chunks));
      const vml = await zip.file('xl/drawings/vmlDrawing1.vml').async('string');
      expect(vml).to.contain(FIT_TO_TEXT);
    });
  });
});
