const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const IMAGE_FILENAME = `${__dirname}/../data/image.png`;

// a drawing part that does not use the 'xdr:' prefix (default namespace instead)
const DRAWING_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>' +
  '<to><col>2</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></to>' +
  '<clientData/></twoCellAnchor>' +
  '</wsDr>';

async function buildWorkbook({keepRels}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'hello';
  const imageId = wb.addImage({filename: IMAGE_FILENAME, extension: 'png'});
  ws.addImage(imageId, 'A1:B2');
  const buffer = await wb.xlsx.writeBuffer();

  const zip = await JSZip.loadAsync(buffer);
  expect(zip.file('xl/drawings/drawing1.xml')).to.be.ok();
  zip.file('xl/drawings/drawing1.xml', DRAWING_XML);
  if (!keepRels) {
    zip.remove('xl/drawings/_rels/drawing1.xml.rels');
  }
  return zip.generateAsync({type: 'nodebuffer'});
}

describe('github issues', () => {
  describe('issue 2591 - drawing part with an unrecognised root', () => {
    [true, false].forEach(keepRels => {
      it(`loads without crashing (drawing rels ${keepRels ? 'present' : 'missing'})`, async () => {
        const buffer = await buildWorkbook({keepRels});

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const ws = wb.getWorksheet('Sheet1');
        expect(ws.getCell('A1').value).to.equal('hello');
        expect(ws.getImages()).to.deep.equal([]);

        // and the loaded workbook can be written again
        const out = await wb.xlsx.writeBuffer();
        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(out);
        expect(wb2.getWorksheet('Sheet1').getCell('A1').value).to.equal('hello');
      });
    });
  });
});
