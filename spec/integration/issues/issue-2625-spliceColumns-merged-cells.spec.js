const ExcelJS = verquire('exceljs');

// row 1: a | b c (merged B1:C1) | d | e f (merged E1:F1) | g
function createSheet() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('splice');
  ['a', 'b', 'c', 'd', 'e', 'f', 'g'].forEach((value, i) => {
    ws.getCell(1, i + 1).value = value;
  });
  ws.mergeCells('B1:C1');
  ws.mergeCells('E1:F1');
  return {wb, ws};
}

function merges(ws) {
  return Object.values(ws._merges)
    .map(merge => merge.range)
    .sort();
}

function values(ws) {
  return ws.getRow(1).values.slice(1);
}

describe('github issues', () => {
  describe('issue 2625 - spliceColumns with merged cells', () => {
    it('shifts both merges when deleting a column before them', () => {
      const {ws} = createSheet();
      ws.spliceColumns(1, 1);

      expect(merges(ws)).to.deep.equal(['A1:B1', 'D1:E1']);
      expect(values(ws)).to.deep.equal(['b', 'b', 'd', 'e', 'e', 'g']);
      expect(ws.getCell('B1').isMergedTo(ws.getCell('A1'))).to.be.true();
      expect(ws.getCell('C1').isMerged).to.be.false();
      expect(ws.getCell('E1').isMergedTo(ws.getCell('D1'))).to.be.true();
      expect(ws.getCell('F1').isMerged).to.be.false();
    });

    it('drops a two column merge when deleting a column inside it', () => {
      const {ws} = createSheet();
      ws.spliceColumns(3, 1);

      expect(merges(ws)).to.deep.equal(['D1:E1']);
      expect(values(ws)).to.deep.equal(['a', 'b', 'd', 'e', 'e', 'g']);
      expect(ws.getCell('B1').isMerged).to.be.false();
      expect(ws.getCell('E1').isMergedTo(ws.getCell('D1'))).to.be.true();
    });

    it('shifts only the later merge when deleting a column between them', () => {
      const {ws} = createSheet();
      ws.spliceColumns(4, 1);

      expect(merges(ws)).to.deep.equal(['B1:C1', 'D1:E1']);
      expect(values(ws)).to.deep.equal(['a', 'b', 'b', 'e', 'e', 'g']);
      expect(ws.getCell('D1').value).to.equal('e');
      expect(ws.getCell('F1').isMerged).to.be.false();
    });

    it('leaves merges alone when deleting a column after them', () => {
      const {ws} = createSheet();
      ws.spliceColumns(7, 1);

      expect(merges(ws)).to.deep.equal(['B1:C1', 'E1:F1']);
      expect(values(ws)).to.deep.equal(['a', 'b', 'b', 'd', 'e', 'e']);
    });

    it('removes merges whose columns are all deleted', () => {
      const {ws} = createSheet();
      ws.spliceColumns(2, 3);

      expect(merges(ws)).to.deep.equal(['B1:C1']);
      expect(values(ws)).to.deep.equal(['a', 'e', 'e', 'g']);
    });

    it('shrinks a wider, multi-row merge and keeps a merged column', () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('splice');
      ws.getCell('B1').value = 'm';
      ws.mergeCells('B1:D2');
      ws.spliceColumns(3, 2);

      expect(merges(ws)).to.deep.equal(['B1:B2']);
      expect(ws.getCell('B1').value).to.equal('m');
      expect(ws.getCell('B2').isMergedTo(ws.getCell('B1'))).to.be.true();
      expect(ws.getCell('C1').isMerged).to.be.false();
    });

    it('shifts merges after inserted columns and grows merges around them', () => {
      const {ws} = createSheet();
      ws.spliceColumns(2, 0, ['x']);
      expect(merges(ws)).to.deep.equal(['C1:D1', 'F1:G1']);
      expect(values(ws)).to.deep.equal(['a', 'x', 'b', 'b', 'd', 'e', 'e', 'g']);

      ws.spliceColumns(4, 0, ['y']);
      expect(merges(ws)).to.deep.equal(['C1:E1', 'G1:H1']);
      expect(ws.getCell('D1').isMergedTo(ws.getCell('C1'))).to.be.true();
    });

    it('keeps the merged cells own styles', () => {
      const {ws} = createSheet();
      ws.getCell('F1').border = {right: {style: 'thick'}};
      ws.spliceColumns(1, 1);

      expect(ws.getCell('E1').border).to.deep.equal({right: {style: 'thick'}});
    });

    it('writes the shifted merges', async () => {
      const {wb, ws} = createSheet();
      ws.spliceColumns(1, 1);
      const buffer = await wb.xlsx.writeBuffer();

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('splice');
      expect(merges(ws2)).to.deep.equal(['A1:B1', 'D1:E1']);
      expect(values(ws2)).to.deep.equal(['b', 'b', 'd', 'e', 'e', 'g']);
    });
  });
});
