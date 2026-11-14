const fs = require('fs');
const path = require('path');

const ExcelJS = verquire('exceljs');

const fileName = path.join(__dirname, '../data/gold.xlsx');

describe('github issues', () => {
  describe('issue 2877 - load() accepts Buffer, Uint8Array and ArrayBuffer', () => {
    it('loads the same workbook from each input type', async () => {
      const fileBuffer = fs.readFileSync(fileName);
      const bytes = new Uint8Array(fileBuffer);
      const inputs = [fileBuffer, bytes, bytes.buffer];
      for (const input of inputs) {
        const wb = new ExcelJS.Workbook();
        // eslint-disable-next-line no-await-in-loop
        await wb.xlsx.load(input);
        expect(wb.worksheets.length).to.be.greaterThan(0);
      }
    });

    it('writeBuffer() returns a node Buffer', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('sheet').getCell('A1').value = 7;
      expect(Buffer.isBuffer(await wb.xlsx.writeBuffer())).to.equal(true);
      expect(Buffer.isBuffer(await wb.csv.writeBuffer())).to.equal(true);
    });
  });
});
