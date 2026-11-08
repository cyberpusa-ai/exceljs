const {createSheetMock} = require('../../utils/index');

const Anchor = verquire('doc/anchor');

describe('Anchor', () => {
  describe('colWidth', () => {
    it('should colWidth equal the 64px Excel default when worksheet is undefined', () => {
      const anchor = new Anchor();
      expect(anchor.colWidth).to.equal(64 * 9525);
    });
    it('should colWidth equal the 64px Excel default when column has no width', () => {
      const anchor = new Anchor(createSheetMock());
      expect(anchor.colWidth).to.equal(64 * 9525);
    });
    it('should colWidth use the worksheet defaultColWidth', () => {
      const worksheet = createSheetMock();
      worksheet.properties.defaultColWidth = 12;
      const anchor = new Anchor(worksheet);
      expect(anchor.colWidth).to.equal(12 * 7 * 9525);
    });
    it('should colWidth equal column width', () => {
      const worksheet = createSheetMock();
      worksheet.properties.defaultColWidth = 12;
      const anchor = new Anchor(worksheet);
      worksheet.addColumn(anchor.nativeCol + 1, {
        width: 10,
      });
      expect(anchor.colWidth).to.equal(10 * 7 * 9525);
    });
    it('should colWidth use the width of the anchored column', () => {
      const worksheet = createSheetMock();
      worksheet.getColumn(1).width = 10;
      worksheet.getColumn(3).width = 30;
      const anchor = new Anchor(worksheet, {nativeCol: 2, nativeRow: 0});
      expect(anchor.colWidth).to.equal(30 * 7 * 9525);
    });
    it('should colWidth fall back to the default for a plain model worksheet', () => {
      const anchor = new Anchor({}, {nativeCol: 2, nativeRow: 0});
      expect(anchor.colWidth).to.equal(64 * 9525);
    });
  });
  describe('rowHeight', () => {
    it('should rowHeight equal 15pt when worksheet is undefined', () => {
      const anchor = new Anchor();
      expect(anchor.rowHeight).to.equal(15 * 12700);
    });
    it('should rowHeight equal 15pt when row has not set height', () => {
      const anchor = new Anchor(createSheetMock());
      expect(anchor.rowHeight).to.equal(15 * 12700);
    });
    it('should rowHeight use the worksheet defaultRowHeight', () => {
      const worksheet = createSheetMock();
      worksheet.properties.defaultRowHeight = 25;
      const anchor = new Anchor(worksheet);
      expect(anchor.rowHeight).to.equal(25 * 12700);
    });
    it('should rowHeight equal row height', () => {
      const worksheet = createSheetMock();
      worksheet.properties.defaultRowHeight = 25;
      worksheet.getRow(1).height = 10;

      const anchor = new Anchor(worksheet);
      expect(anchor.rowHeight).to.equal(10 * 12700);
    });
    it('should rowHeight fall back to the default for a plain model worksheet', () => {
      const anchor = new Anchor({}, {nativeCol: 0, nativeRow: 2});
      expect(anchor.rowHeight).to.equal(15 * 12700);
    });
  });
  describe('fractional col and row', () => {
    it('should place offsets relative to the real cell size', () => {
      const worksheet = createSheetMock();
      worksheet.getColumn(2).width = 20;
      worksheet.getRow(3).height = 40;
      const anchor = new Anchor(worksheet, {col: 1.5, row: 2.25});
      expect(anchor.nativeCol).to.equal(1);
      expect(anchor.nativeColOff).to.equal((20 * 7 * 9525) / 2);
      expect(anchor.nativeRow).to.equal(2);
      expect(anchor.nativeRowOff).to.equal((40 * 12700) / 4);
      expect(anchor.col).to.equal(1.5);
      expect(anchor.row).to.equal(2.25);
    });
  });
  describe('resize worksheet`s cells', () => {
    const context = this;
    before(() => {
      context.worksheet = createSheetMock();
      context.worksheet.getColumn(1).width = 20;
      context.worksheet.getRow(1).height = 20;

      context.anchor = new Anchor(context.worksheet, {col: 0.6, row: 0.6});
    });

    it('should update colWidth', () => {
      const pre = context.anchor.colWidth;
      context.worksheet.getColumn(1).width *= 2;
      expect(context.anchor.colWidth).to.not.equal(pre);
      expect(context.anchor.colWidth).to.equal(pre * 2);
    });
    it('should update rowHeight', () => {
      const pre = context.anchor.rowHeight;
      context.worksheet.getRow(1).height *= 2;
      expect(context.anchor.rowHeight).to.not.equal(pre);
      expect(context.anchor.rowHeight).to.equal(pre * 2);
    });
    it('should recalculate col', () => {
      const pre = context.anchor.col;
      context.worksheet.getColumn(1).width *= 2;
      expect(context.anchor.col).to.not.equal(pre);
    });
    it('should recalculate row', () => {
      const pre = context.anchor.row;
      context.worksheet.getRow(1).height *= 2;
      expect(context.anchor.row).to.not.equal(pre);
    });
    it('should integer part of row and rowOff should be always equals', () => {
      expect(Math.floor(context.anchor.row)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getRow(1).height *= 2;
      expect(Math.floor(context.anchor.row)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getRow(1).height /= 4;
      expect(Math.floor(context.anchor.row)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getRow(1).height = 0.1;
      expect(Math.floor(context.anchor.row)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getRow(1).height = 9999;
      expect(Math.floor(context.anchor.row)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
    });
    it('should integer part of col and colOff should be always equals', () => {
      expect(Math.floor(context.anchor.col)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getColumn(1).width *= 2;
      expect(Math.floor(context.anchor.col)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getColumn(1).width /= 4;
      expect(Math.floor(context.anchor.col)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getColumn(1).width = 0.1;
      expect(Math.floor(context.anchor.col)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
      context.worksheet.getColumn(1).width = 9999;
      expect(Math.floor(context.anchor.col)).to.equal(
        Math.floor(context.anchor.nativeCol)
      );
    });
    it('should update nativeColOff after col has been changed', () => {
      const pre = context.anchor.nativeColOff;
      context.anchor.col -= 0.321;
      expect(context.anchor.nativeColOff).to.not.equal(pre);
    });
    it('should update nativeRowOff after row has been changed', () => {
      const pre = context.anchor.nativeRowOff;
      context.anchor.row -= 0.321;
      expect(context.anchor.nativeColOff).to.not.equal(pre);
    });
  });
});
