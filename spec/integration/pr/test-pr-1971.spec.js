const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test-pr-1971.xlsx';

async function getDefinedNames(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('xl/workbook.xml').async('string');
  return xml.match(/<definedName [^>]*>[^<]*<\/definedName>/g) || [];
}

async function writeBuffer(printArea) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 1;
  ws.pageSetup.printArea = printArea;
  return wb.xlsx.writeBuffer();
}

async function readPrintArea(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb.getWorksheet('Sheet1').pageSetup.printArea;
}

const CASES = [
  {printArea: 'A1:B2', refs: ['$A$1:$B$2']},
  {printArea: 'A1:B2&&D1:E2', refs: ['$A$1:$B$2', '$D$1:$E$2']},
  {printArea: 'A:B', refs: ['$A:$B']},
  {printArea: '1:3', refs: ['$1:$3']},
  {printArea: 'C3', refs: ['$C$3']},
  {printArea: 'A1:B2&&D:E&&2:4', refs: ['$A$1:$B$2', '$D:$E', '$2:$4']},
];

describe('github issues', () => {
  describe('pr 1971 - multiple print areas are written as one Print_Area', () => {
    CASES.forEach(({printArea, refs}) => {
      it(`writes and reads back printArea ${printArea}`, async () => {
        const buffer = await writeBuffer(printArea);
        const definedNames = await getDefinedNames(buffer);
        expect(definedNames).to.have.length(1);
        expect(definedNames[0]).to.equal(
          `<definedName name="_xlnm.Print_Area" localSheetId="0">${refs
            .map(ref => `&apos;Sheet1&apos;!${ref}`)
            .join(',')}</definedName>`
        );
        expect(await readPrintArea(buffer)).to.equal(printArea);
      });
    });

    it('writes one Print_Area per sheet with the streaming writer', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
        filename: TEST_XLSX_FILE_NAME,
      });
      const ws = wb.addWorksheet('Sheet1');
      ws.pageSetup.printArea = 'A1:B2&&D:E';
      ws.getCell('A1').value = 1;
      ws.commit();
      await wb.commit();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      expect(wb2.getWorksheet('Sheet1').pageSetup.printArea).to.equal(
        'A1:B2&&D:E'
      );
    });

    it('reads every area of a multi-area Print_Area saved by Excel', async () => {
      const zip = await JSZip.loadAsync(await writeBuffer('A1:B2'));
      const xml = await zip.file('xl/workbook.xml').async('string');
      zip.file(
        'xl/workbook.xml',
        xml.replace(
          /<definedNames>.*<\/definedNames>/,
          '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">' +
            'Sheet1!$B$2:$A$1,Sheet1!$D:$E,Sheet1!$4:$6</definedName></definedNames>'
        )
      );
      const buffer = await zip.generateAsync({type: 'nodebuffer'});
      expect(await readPrintArea(buffer)).to.equal('A1:B2&&D:E&&4:6');
    });
  });
});
