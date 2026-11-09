const Excel = verquire('exceljs');

function addTable(ws, tableOptions) {
  return ws.addTable(tableOptions);
}

describe('Error handling of table creation', () => {
  let tableOptions;

  beforeEach(() => {
    // Reuse this object for multiple tests of table creation error handling
    tableOptions = {
      name: 'TestTable',
      ref: 'A1',
      headerRow: true,
      totalsRow: true,
      style: {
        theme: 'TableStyleDark3',
        showRowStripes: true,
      },
      columns: [
        {name: 'Date', totalsRowLabel: 'Totals', filterButton: true},
        {
          name: 'Id',
          totalsRowFunction: 'max',
          filterButton: true,
          totalsRowResult: 4,
        },
        {
          name: 'Word',
          filterButton: false,
          style: {font: {bold: true, name: 'Comic Sans MS'}},
        },
      ],
      rows: [
        [new Date('2020-01-01'), 1, 'one'],
        [new Date('2020-01-02'), 2, 'two'],
        [new Date('2020-01-03'), 3, 'three'],
      ],
    };
  });

  it('should create a table without error when all required options are provided', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    expect(() => addTable(ws, tableOptions)).to.not.throw();
  });

  it('should throw an error when adding a table without rows', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    delete tableOptions.rows;
    expect(() => addTable(ws, tableOptions)).to.throw('Table must have row definitions');
  });

  it('should add one blank data row when adding a table with empty rows', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    tableOptions.rows = []; // Excel needs at least one data row
    const table = addTable(ws, tableOptions);
    expect(table.table.rows).to.deep.equal([[null, null, null]]);
    expect(table.table.tableRef).to.equal('A1:C3');
  });
});
