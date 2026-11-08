const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

const richA = {richText: [{font: {bold: true}, text: 'value 1'}]};
const richB = {richText: [{font: {italic: true}, text: 'value 2'}]};

function fill(ws) {
  ws.getCell('A1').value = richA;
  ws.getCell('A2').value = richB;
  ws.getCell('A3').value = richA;
  ws.getCell('A4').value = 'plain';
}

function check(ws) {
  expect(ws.getCell('A1').value).to.deep.equal(richA);
  expect(ws.getCell('A2').value).to.deep.equal(richB);
  expect(ws.getCell('A3').value).to.deep.equal(richA);
  expect(ws.getCell('A4').value).to.equal('plain');
}

describe('github issues', () => {
  describe('issue 2267 - distinct richText values are merged with shared strings', () => {
    it('document writer', async () => {
      const wb = new ExcelJS.Workbook();
      fill(wb.addWorksheet('Sheet1'));
      const buffer = await wb.xlsx.writeBuffer({useSharedStrings: true});

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      check(wb2.getWorksheet('Sheet1'));
    });

    it('streaming writer', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
        useSharedStrings: true,
      });
      const ws = wb.addWorksheet('Sheet1');
      fill(ws);
      ws.commit();
      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      check(wb2.getWorksheet('Sheet1'));
    });
  });
});
