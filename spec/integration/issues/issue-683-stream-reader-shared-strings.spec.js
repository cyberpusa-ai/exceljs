const fs = require('fs');

const ExcelJS = verquire('exceljs');

// dateIssue.xlsx stores xl/worksheets/sheet1.xml before xl/sharedStrings.xml, so the sheet is
// buffered to a temp file. unzipper used to signal 'end' before sharedStrings.xml was read and the
// reader then emitted {sharedString: 0} instead of the text. The race only shows up sometimes, so
// read the file several times.
describe('github issues', () => {
  it('issue 683 - stream reader resolves shared strings stored after the worksheet', async () => {
    for (let i = 0; i < 30; i++) {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        fs.createReadStream('./spec/integration/data/dateIssue.xlsx'),
        {sharedStrings: 'cache', styles: 'cache'}
      );
      const values = [];
      // eslint-disable-next-line no-await-in-loop
      for await (const worksheet of workbookReader) {
        for await (const row of worksheet) {
          values.push(row.values[1]);
        }
      }
      expect(values).to.deep.equal(['Date', new Date('2020-11-20T00:00:00.000Z')]);
    }
  });
});
