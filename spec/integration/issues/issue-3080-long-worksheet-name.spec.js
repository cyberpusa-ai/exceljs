const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('issue 3080 - worksheet names longer than 31 chars', () => {
    const name = 'ThisIsAWorksheetNameThatIsLongerThan31';

    it('allows reassigning the same long name', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(name);

      ws.name = name;

      expect(ws.name).to.equal(name.substring(0, 31));
    });

    it('still rejects a long name that collides with another sheet', () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet(name);
      const ws = wb.addWorksheet('Other');

      expect(() => {
        ws.name = name.toUpperCase();
      }).to.throw('Worksheet name already exists');
    });

    it('loads a workbook with a long worksheet name', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Short').getCell('A1').value = 'preserved';
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      const xml = await zip.file('xl/workbook.xml').async('string');
      zip.file('xl/workbook.xml', xml.replace('name="Short"', `name="${name}"`));
      const loaded = new ExcelJS.Workbook();

      await loaded.xlsx.load(await zip.generateAsync({type: 'nodebuffer'}));

      expect(loaded.worksheets[0].name).to.equal(name.substring(0, 31));
      expect(loaded.worksheets[0].getCell('A1').value).to.equal('preserved');
    });
  });
});
