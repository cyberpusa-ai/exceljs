const ExcelJS = verquire('exceljs');

const VALUES = {
  A1: {formula: 'B1', result: 0},
  A2: {formula: 'B2', result: false},
  A3: {formula: 'B3', result: ''},
  A4: {formula: 'B4'},
};

describe('github issues', () => {
  describe('issue 2943 - falsy formula results are dropped', () => {
    it('keeps 0, false and empty string results through an xlsx round trip', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      Object.entries(VALUES).forEach(([address, value]) => {
        ws.getCell(address).value = value;
        expect(ws.getCell(address).value).to.deep.equal(value);
      });

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('Sheet1');
      Object.entries(VALUES).forEach(([address, value]) => {
        expect(ws2.getCell(address).value).to.deep.equal(value);
      });
    });

    it('keeps 0, false and empty string results with the streaming writer', async () => {
      const stream = new ExcelJS.stream.xlsx.WorkbookWriter({});
      const ws = stream.addWorksheet('Sheet1');
      Object.entries(VALUES).forEach(([address, value]) => {
        ws.getCell(address).value = value;
      });
      ws.commit();
      await stream.commit();

      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(stream.stream.read());
      const ws2 = wb.getWorksheet('Sheet1');
      Object.entries(VALUES).forEach(([address, value]) => {
        expect(ws2.getCell(address).value).to.deep.equal(value);
      });
    });
  });
});
