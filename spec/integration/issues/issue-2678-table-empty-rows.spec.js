const ExcelJS = verquire('exceljs');

function tableOptions(rows, extra) {
  return {
    name: 'EmptyTable',
    ref: 'B2',
    columns: [{name: 'Date'}, {name: 'Id'}, {name: 'Word'}],
    rows,
    ...extra,
  };
}

describe('github issues', () => {
  describe('issue 2678 - table with an empty rows array', () => {
    it('keeps one blank data row so the table ref is valid', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = ws.addTable(tableOptions([]));

      expect(table.table.rows).to.deep.equal([[null, null, null]]);
      expect(table.table.tableRef).to.equal('B2:D3');
      expect(table.table.autoFilterRef).to.equal('B2:D3');
      expect(ws.getCell('B2').value).to.equal('Date');
      expect(ws.getCell('B3').value).to.equal(null);

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('blort');
      expect(ws2.getTable('EmptyTable').table.tableRef).to.equal('B2:D3');
    });

    it('puts the totals row below the blank data row', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = ws.addTable(tableOptions([], {totalsRow: true}));

      expect(table.table.tableRef).to.equal('B2:D4');
      expect(table.table.autoFilterRef).to.equal('B2:D3');
      expect(ws.getCell('B3').value).to.equal(null);
      expect(ws.getCell('B4').value).to.equal('Total');
    });

    it('gives a table without a header row one data row', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = ws.addTable(tableOptions([], {headerRow: false}));

      expect(table.table.tableRef).to.equal('B2:D2');
    });

    it('keeps one blank data row after all rows are removed', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = ws.addTable(
        tableOptions([
          [new Date('2020-01-01'), 1, 'one'],
          [new Date('2020-01-02'), 2, 'two'],
        ])
      );
      table.removeRows(0, 2);
      table.commit();

      expect(table.table.tableRef).to.equal('B2:D3');
      expect(ws.getCell('B3').value).to.equal(null);
      expect(ws.getCell('C3').value).to.equal(null);
      expect(ws.getCell('B4').value).to.equal(null);
    });

    it('still requires rows', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      expect(() => ws.addTable(tableOptions(undefined))).to.throw(
        'Table must have row definitions'
      );
    });
  });
});
