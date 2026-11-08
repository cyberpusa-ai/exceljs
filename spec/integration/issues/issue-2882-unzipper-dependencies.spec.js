const fs = require('fs');
const path = require('path');
const {PassThrough} = require('readable-stream');

const ExcelJS = verquire('exceljs');

// unzipper 0.10 pulled in binary -> chainsaw (licensed "MIT/X11") and the deprecated
// fstream -> rimraf@2. unzipper 0.12 drops both. Check that the installed unzipper is a version
// without them, and that the streaming reader still reads deflated and stored entries with it.
async function readAll(stream, options) {
  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(stream, options);
  const sheets = [];
  for await (const worksheet of workbookReader) {
    const rows = [];
    for await (const row of worksheet) {
      rows.push(row.values.slice(1));
    }
    sheets.push(rows);
  }
  return sheets;
}

async function writeWorkbook(zip) {
  const stream = new PassThrough();
  const chunks = [];
  stream.on('data', chunk => chunks.push(chunk));
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream,
    useSharedStrings: true,
    useStyles: true,
    zip,
  });
  for (let s = 1; s <= 3; s++) {
    const worksheet = workbook.addWorksheet(`Sheet ${s}`);
    for (let r = 1; r <= 200; r++) {
      worksheet.addRow([`s${s}r${r}`, r * s, new Date(Date.UTC(2020, 0, r))]).commit();
    }
    worksheet.commit();
  }
  await workbook.commit();
  return Buffer.concat(chunks);
}

function expected(s) {
  const rows = [];
  for (let r = 1; r <= 200; r++) {
    rows.push([`s${s}r${r}`, r * s, new Date(Date.UTC(2020, 0, r))]);
  }
  return rows;
}

describe('github issues', () => {
  describe('issue 2882 - unzipper without binary/fstream', () => {
    it('depends on an unzipper release without binary and fstream', () => {
      const pkgPath = require.resolve('unzipper/package.json', {
        paths: [path.dirname(require.resolve('../../../package.json'))],
      });
      const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
      const deps = Object.keys(pkg.dependencies || {});
      expect(deps).to.not.include('binary');
      expect(deps).to.not.include('fstream');
    });

    for (const [name, zip] of [
      ['deflated', undefined],
      ['stored', {store: true}],
    ]) {
      it(`stream reader reads ${name} entries`, async () => {
        const buffer = await writeWorkbook(zip);
        const input = new PassThrough();
        input.end(buffer);
        const sheets = await readAll(input, {sharedStrings: 'cache', styles: 'cache'});
        expect(sheets).to.deep.equal([expected(1), expected(2), expected(3)]);
      });
    }

    it('stream reader still reads a file with the sheet before sharedStrings', async () => {
      const sheets = await readAll(fs.createReadStream('./spec/integration/data/dateIssue.xlsx'), {
        sharedStrings: 'cache',
        styles: 'cache',
      });
      expect(sheets).to.deep.equal([[['Date'], [new Date('2020-11-20T00:00:00.000Z')]]]);
    });
  });
});
