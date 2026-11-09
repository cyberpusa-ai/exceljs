const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 2675 - addTable rejects names Excel would have to repair', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet();
    const table = {
      ref: 'A1',
      columns: [{name: 'Code'}, {name: 'Count'}],
      rows: [[404, 3]],
    };

    expect(() => ws.addTable({...table, name: '404 Errors'})).to.throw(
      'Invalid table name \'404 Errors\''
    );
    expect(ws.getTable('404 Errors')).to.be.undefined();

    ws.addTable({...table, name: 'Errors_404'});
    expect(ws.getTable('Errors_404')).to.be.ok();
  });

  it('issue 2675 - a table with a valid name round-trips through xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('sheet');
    ws.addTable({
      name: 'Tabelle_Übersicht.1',
      ref: 'A1',
      columns: [{name: 'Code'}, {name: 'Count'}],
      rows: [[404, 3]],
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(await wb.xlsx.writeBuffer());
    const table = wb2.getWorksheet('sheet').getTable('Tabelle_Übersicht.1');
    expect(table).to.be.ok();
    expect(table.table.tableRef).to.equal('A1:B2');
  });
});
