const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('issue 2678 - addTable with an empty rows array', () => {
    const columns = [{name: 'Date'}, {name: 'Id'}, {name: 'Word'}];

    it('keeps one blank data row so the table ref is valid', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();
      const table = ws.addTable({name: 'Empty', ref: 'B2', columns, rows: []});

      expect(table.table.tableRef).to.equal('B2:D3');
      expect(table.table.autoFilterRef).to.equal('B2:D3');
      expect(table.table.rows).to.deep.equal([]);
      expect(ws.getCell('B2').value).to.equal('Date');
      expect(ws.getCell('B3').value).to.be.null();
    });

    it('puts the totals row below the blank data row', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();
      const table = ws.addTable({
        name: 'EmptyTotals',
        ref: 'A1',
        totalsRow: true,
        columns: [{name: 'Label'}, {name: 'Amount', totalsRowFunction: 'sum'}],
        rows: [],
      });

      expect(table.table.tableRef).to.equal('A1:B3');
      expect(table.table.autoFilterRef).to.equal('A1:B2');
      expect(ws.getCell('A2').value).to.be.null();
      expect(ws.getCell('A3').value).to.equal('Total');
      expect(ws.getCell('B3').value).to.deep.include({
        formula: 'SUBTOTAL(109,EmptyTotals[Amount])',
      });
    });

    it('does not leave a blank row above rows added later', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();
      const table = ws.addTable({
        name: 'Filled',
        ref: 'A1',
        columns: [{name: 'a'}, {name: 'b'}],
        rows: [],
      });
      table.addRow([1, 2]);
      table.addRow([3, 4]);
      table.commit();

      expect(table.table.tableRef).to.equal('A1:B3');
      expect(table.table.rows).to.deep.equal([
        [1, 2],
        [3, 4],
      ]);
      expect(ws.getCell('A2').value).to.equal(1);
      expect(ws.getCell('A3').value).to.equal(3);
      expect(ws.getCell('A4').value).to.be.null();
    });

    it('clears the last data row when all rows are removed', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();
      const table = ws.addTable({
        name: 'Emptied',
        ref: 'A1',
        totalsRow: true,
        columns: [{name: 'a'}, {name: 'b', totalsRowFunction: 'sum'}],
        rows: [
          [1, 2],
          [3, 4],
        ],
      });
      table.removeRows(0, 2);
      table.commit();

      expect(table.table.tableRef).to.equal('A1:B3');
      expect(ws.getCell('A2').value).to.be.null();
      expect(ws.getCell('B2').value).to.be.null();
      expect(ws.getCell('A3').value).to.equal('Total');
      expect(ws.getCell('A4').value).to.be.null();
    });

    it('writes a table ref that covers a data row', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('sheet');
      ws.addTable({name: 'Empty', ref: 'B2', columns, rows: []});

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(await wb.xlsx.writeBuffer());
      const table = wb2.getWorksheet('sheet').getTable('Empty');
      expect(table.table.tableRef).to.equal('B2:D3');
      expect(table.table.autoFilterRef).to.equal('B2:D3');
    });
  });
});
