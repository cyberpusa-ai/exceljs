const fs = require('fs');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', () => {
  it('issue 1595 - WorkbookWriter.stream.pipe() returns the destination', async () => {
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter();
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['Hello', 'World']).commit();
    ws.commit();

    const out = fs.createWriteStream(TEST_XLSX_FILE_NAME);
    const piped = wb.stream.pipe(out);
    expect(piped).to.equal(out);

    const finished = new Promise((resolve, reject) => {
      piped.on('finish', resolve);
      piped.on('error', reject);
    });
    await wb.commit();
    await finished;

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
    expect(wb2.getWorksheet('Sheet1').getCell('B1').value).to.equal('World');
  });
});
