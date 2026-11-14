const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', () => {
  describe('issue 1363 - hide AutoFilter buttons on excluded columns', () => {
    it('round-trips autoFilter.exclude through the document writer', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      ws.mergeCells('A1:B1');
      ws.getCell('A1').value = 'Merged';
      ws.getCell('C1').value = 'C';
      ws.autoFilter = {from: 'A1', to: 'C1', exclude: [1]};

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      expect(wb2.getWorksheet('sheet').autoFilter).to.deep.equal({
        from: 'A1',
        to: 'C1',
        exclude: [1],
      });
    });

    it('writes autoFilter.exclude from the streaming writer', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: TEST_XLSX_FILE_NAME});
      const ws = wb.addWorksheet('sheet');
      ws.autoFilter = {from: 'A1', to: 'E1', exclude: [0, 2]};
      ws.addRow(['a', 'b', 'c', 'd', 'e']).commit();
      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.getWorksheet('sheet').autoFilter).to.deep.equal({
        from: 'A1',
        to: 'E1',
        exclude: [0, 2],
      });
    });

    it('keeps plain ranges as strings', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      ws.autoFilter = 'A1:C1';
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(await wb.xlsx.writeBuffer());
      expect(wb2.getWorksheet('sheet').autoFilter).to.equal('A1:C1');
    });
  });
});
