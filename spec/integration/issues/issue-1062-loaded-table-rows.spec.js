const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

function createWorkbook(options = {}) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');
  worksheet.getCell('A1').value = 'outside';
  worksheet.addTable({
    name: 'Scores',
    ref: 'B2',
    headerRow: true,
    totalsRow: !!options.totalsRow,
    columns: [
      {name: 'Name', totalsRowLabel: 'Sum'},
      {name: 'Score', totalsRowFunction: 'sum'},
    ],
    rows: [
      ['alice', 1],
      ['bob', 2],
    ],
  });
  return workbook;
}

async function reload(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(buffer);
  return loaded;
}

// Excel omits headerRowCount (it defaults to 1) and writes totalsRowShown="0"
async function excelLike(workbook) {
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const name = Object.keys(zip.files).find(n => /^xl\/tables\/table\d+\.xml$/.test(n));
  let xml = await zip.file(name).async('string');
  xml = xml
    .replace(/ headerRowCount="1"/, '')
    .replace(/ totalsRowShown="1"/, ' totalsRowShown="0"');
  zip.file(name, xml);
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(await zip.generateAsync({type: 'nodebuffer'}));
  return loaded;
}

describe('github issues', () => {
  describe('issue 1062 - append rows to a loaded table', () => {
    it('reads table.rows and ref of a loaded table (issue 2297)', async () => {
      const workbook = await reload(createWorkbook());
      const table = workbook.getWorksheet('Sheet1').getTable('Scores');
      expect(table.ref).to.equal('B2');
      expect(table.headerRow).to.equal(true);
      expect(table.table.rows).to.deep.equal([
        ['alice', 1],
        ['bob', 2],
      ]);
      expect(table.table.columns.map(c => c.name)).to.deep.equal(['Name', 'Score']);
    });

    it('round trips a loaded table unchanged', async () => {
      const workbook = await reload(await reload(createWorkbook({totalsRow: true})));
      const worksheet = workbook.getWorksheet('Sheet1');
      const table = worksheet.getTable('Scores');
      expect(table.table.tableRef).to.equal('B2:C5');
      expect(table.table.rows).to.deep.equal([
        ['alice', 1],
        ['bob', 2],
      ]);
      expect(worksheet.getCell('B5').value).to.equal('Sum');
      expect(worksheet.getCell('C5').formula).to.equal('SUBTOTAL(109,Scores[Score])');
      expect(worksheet.getCell('A1').value).to.equal('outside');
    });

    it('adds a row to a loaded table, commits and writes it', async () => {
      const workbook = await reload(createWorkbook());
      const worksheet = workbook.getWorksheet('Sheet1');
      const table = worksheet.getTable('Scores');
      table.addRow(['carol', 3]);
      table.commit();
      expect(worksheet.getCell('B5').value).to.equal('carol');
      expect(worksheet.getCell('C5').value).to.equal(3);

      const again = await reload(workbook);
      const reloaded = again.getWorksheet('Sheet1').getTable('Scores');
      expect(reloaded.table.tableRef).to.equal('B2:C5');
      expect(reloaded.table.rows).to.deep.equal([
        ['alice', 1],
        ['bob', 2],
        ['carol', 3],
      ]);
      expect(again.getWorksheet('Sheet1').getCell('A1').value).to.equal('outside');
    });

    it('removes rows from a loaded table', async () => {
      const workbook = await reload(createWorkbook());
      const worksheet = workbook.getWorksheet('Sheet1');
      const table = worksheet.getTable('Scores');
      table.removeRows(0, 1);
      table.commit();
      expect(worksheet.getCell('B3').value).to.equal('bob');
      expect(worksheet.getCell('B4').value).to.equal(null);
      const again = await reload(workbook);
      expect(again.getWorksheet('Sheet1').getTable('Scores').table.tableRef).to.equal('B2:C3');
    });

    it('handles an Excel-like table with a totals row', async () => {
      const workbook = await excelLike(createWorkbook({totalsRow: true}));
      const worksheet = workbook.getWorksheet('Sheet1');
      const table = worksheet.getTable('Scores');
      expect(table.headerRow).to.equal(true);
      expect(table.totalsRow).to.equal(true);
      expect(table.table.rows).to.deep.equal([
        ['alice', 1],
        ['bob', 2],
      ]);

      table.addRow(['carol', 3]);
      table.commit();
      expect(worksheet.getCell('B2').value).to.equal('Name');
      expect(worksheet.getCell('B5').value).to.equal('carol');
      expect(worksheet.getCell('B6').value).to.equal('Sum');
      expect(worksheet.getCell('C6').formula).to.equal('SUBTOTAL(109,Scores[Score])');

      const again = await reload(workbook);
      const reloaded = again.getWorksheet('Sheet1').getTable('Scores');
      expect(reloaded.table.tableRef).to.equal('B2:C6');
      expect(reloaded.headerRow).to.equal(true);
      expect(reloaded.table.rows.length).to.equal(3);
    });

    it('keeps a custom totals formula of a loaded table', async () => {
      const source = new ExcelJS.Workbook();
      source.addWorksheet('Sheet1').addTable({
        name: 'Custom',
        ref: 'A1',
        totalsRow: true,
        columns: [
          {name: 'Name', totalsRowLabel: 'Total'},
          {name: 'Score', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(B2:B3)*2'},
        ],
        rows: [
          ['a', 1],
          ['b', 2],
        ],
      });
      const workbook = await reload(source);
      const worksheet = workbook.getWorksheet('Sheet1');
      const table = worksheet.getTable('Custom');
      table.addRow(['c', 3]);
      table.commit();
      expect(worksheet.getCell('B5').formula).to.equal('SUM(B2:B3)*2');
    });
  });
});
