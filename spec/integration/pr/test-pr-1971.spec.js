const fs = require('fs');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test-pr-1971.xlsx';

function getDefinedNames(xml) {
  return xml.match(/<definedName [^>]*>[^<]*<\/definedName>/g) || [];
}

async function readWorkbookXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('xl/workbook.xml').async('string');
}

describe('github issues', () => {
  describe('pr 1971 - multiple print areas', () => {
    it('writes one absolute Print_Area defined name per sheet', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet One');
      ws.getCell('A1').value = 1;
      ws.pageSetup.printArea = 'A1:B2&&D1:E2';
      const other = wb.addWorksheet('other');
      other.pageSetup.printArea = 'C3:D4';

      const xml = await readWorkbookXml(await wb.xlsx.writeBuffer());
      const printAreas = getDefinedNames(xml).filter(n => n.includes('_xlnm.Print_Area'));
      expect(printAreas).to.deep.equal([
        '<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Sheet One&apos;!$A$1:$B$2,&apos;Sheet One&apos;!$D$1:$E$2</definedName>',
        '<definedName name="_xlnm.Print_Area" localSheetId="1">&apos;other&apos;!$C$3:$D$4</definedName>',
      ]);
    });

    it('round trips multiple print areas', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Sheet One');
      ws.pageSetup.printArea = 'A1:B2&&D1:E2&&G5:H9';

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(await wb.xlsx.writeBuffer());
      expect(wb2.getWorksheet('Sheet One').pageSetup.printArea).to.equal('A1:B2&&D1:E2&&G5:H9');
    });

    it('reads every range of a multi-area Print_Area written by Excel', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('Sheet1');
      const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
      const xml = await zip.file('xl/workbook.xml').async('string');
      const definedNames =
        '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">' +
        'Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2</definedName></definedNames>';
      const patched = xml.includes('<definedNames')
        ? xml.replace(/<definedNames[\s\S]*<\/definedNames>|<definedNames\/>/, definedNames)
        : xml.replace('<calcPr', `${definedNames}<calcPr`);
      zip.file('xl/workbook.xml', patched);

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(await zip.generateAsync({type: 'nodebuffer'}));
      expect(wb2.getWorksheet('Sheet1').pageSetup.printArea).to.equal('A1:B2&&D1:E2');
    });

    it('writes one Print_Area from the streaming writer', async () => {
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: TEST_XLSX_FILE_NAME});
      const ws = wb.addWorksheet('Sheet One', {pageSetup: {printArea: 'A1:B2&&D1:E2'}});
      ws.getCell('A1').value = 1;
      ws.commit();
      await wb.commit();

      const xml = await readWorkbookXml(fs.readFileSync(TEST_XLSX_FILE_NAME));
      const printAreas = getDefinedNames(xml).filter(n => n.includes('_xlnm.Print_Area'));
      expect(printAreas).to.deep.equal([
        '<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Sheet One&apos;!$A$1:$B$2,&apos;Sheet One&apos;!$D$1:$E$2</definedName>',
      ]);
    });
  });
});
