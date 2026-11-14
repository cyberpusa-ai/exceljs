const ExcelJS = verquire('exceljs');

const ERROR = 'Cannot merge already merged cells';

function createSheet() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('merges');
  ws.mergeCells('C3:E5');
  return ws;
}

describe('github issues', () => {
  describe('issue 1842 - mergeCells overlap check', () => {
    [
      ['top-left corner', 'A1:C3'],
      ['top-right corner', 'E1:G3'],
      ['bottom-left corner', 'A5:C7'],
      ['bottom-right corner', 'E5:G7'],
      ['top edge', 'D1:D3'],
      ['left edge', 'A4:C4'],
      ['contained', 'D4:D4'],
      ['contained, no master cell', 'D4:E5'],
      ['same range', 'C3:E5'],
      ['containing', 'A1:G7'],
      ['crossing, no master cell', 'A4:G4'],
    ].forEach(([name, range]) => {
      it(`rejects a ${name} overlap (${range})`, () => {
        const ws = createSheet();
        expect(() => ws.mergeCells(range)).to.throw(ERROR);
        expect(() => ws.mergeCellsWithoutStyle(range)).to.throw(ERROR);
      });
    });

    it('rejects overlaps with a single-cell merge', () => {
      const ws = createSheet();
      ws.mergeCells('H8:H8');
      expect(() => ws.mergeCells('G7:H8')).to.throw(ERROR);
    });

    it('leaves the sheet unchanged when a merge is rejected', () => {
      const ws = createSheet();
      ws.getCell('A1').value = 'a1';
      ws.getCell('B2').value = 'b2';
      expect(() => ws.mergeCells('A1:D4')).to.throw(ERROR);

      expect(ws.getCell('A1').value).to.equal('a1');
      expect(ws.getCell('B2').value).to.equal('b2');
      expect(ws.getCell('B2').isMerged).to.equal(false);
      expect(ws.getCell('A4').isMerged).to.equal(false);
      expect(ws.model.merges).to.deep.equal(['C3:E5']);
    });

    it('allows adjacent merges and re-merging after unmerge', () => {
      const ws = createSheet();
      ws.mergeCells('F3:F5');
      ws.mergeCells('C6:E6');
      ws.mergeCells('A1:B2');
      ws.unMergeCells('C3');
      ws.mergeCells('C3:D4');
      expect(ws.model.merges.sort()).to.deep.equal(['A1:B2', 'C3:D4', 'C6:E6', 'F3:F5']);
    });

    it('detects overlaps on a sheet loaded from a file', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('merges').mergeCells('C3:E5');
      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      expect(() => wb2.getWorksheet('merges').mergeCells('D4:F6')).to.throw(ERROR);
    });

    it('handles thousands of merges', function() {
      this.timeout(20000);
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('many');
      for (let r = 1; r <= 5000; r++) {
        ws.mergeCells(r, 1, r, 3);
      }
      expect(Object.keys(ws._merges).length).to.equal(5000);
      expect(() => ws.mergeCells('B2500:D2500')).to.throw(ERROR);
    });
  });
});
