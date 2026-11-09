const Excel = verquire('exceljs');

function addTable(ws, tableOptions) {
  return ws.addTable(tableOptions);
}

describe('Error handling of table creation', () => {
  let tableOptions;

  beforeEach(() => {
    // Resuse this object for multiple tests of table creation error handling
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

  it('should throw an error when adding a table with an invalid name', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    tableOptions.name = '404 Errors'; // invalid name, shouldn't be allowed
    expect(() => addTable(ws, tableOptions)).to.throw('Invalid table name \'404 Errors\'');
  });

  it('should throw an error when adding a table without a name', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    delete tableOptions.name;
    expect(() => addTable(ws, tableOptions)).to.throw('Table must have a name');
  });

  [
    'TestTable',
    '_Table',
    '\\Table',
    'Table.1',
    'Table_2020',
    'ZZZ1',
    'A1B',
    'R1C1x',
    'Tabelle_Übersicht',
    '表1',
    'T'.repeat(255),
  ].forEach(name => {
    it(`should accept the valid table name ${JSON.stringify(name.slice(0, 20))}`, () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      tableOptions.name = name;
      expect(() => addTable(ws, tableOptions)).to.not.throw();
    });
  });

  [
    '1Table',
    '.Table',
    'My Table',
    'My-Table',
    'Table!',
    'A1',
    'xfd1048576',
    'R1C1',
    'rc',
    'R',
    'c',
    'T'.repeat(256),
  ].forEach(name => {
    it(`should reject the invalid table name ${JSON.stringify(name.slice(0, 20))}`, () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      tableOptions.name = name;
      expect(() => addTable(ws, tableOptions)).to.throw('Invalid table name');
    });
  });

  it('should validate the name again when a changed table is committed', () => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('blort');
    const table = addTable(ws, tableOptions);
    table.name = 'Bad Name';
    table.addRow([new Date('2020-01-04'), 4, 'four']);
    expect(() => table.commit()).to.throw('Invalid table name \'Bad Name\'');
  });
});
