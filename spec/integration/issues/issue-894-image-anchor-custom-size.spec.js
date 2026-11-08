const ExcelJS = verquire('exceljs');

const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';
const IMAGE_FILENAME = `${__dirname}/../data/image.png`;

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;

describe('github issues', () => {
  it('issue 894 - fractional image anchors use custom column widths and row heights', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1', {properties: {defaultColWidth: 12, defaultRowHeight: 20}});
    ws.getColumn(2).width = 20;
    ws.getRow(2).height = 40;

    const imageId = wb.addImage({filename: IMAGE_FILENAME, extension: 'png'});
    // tl in the custom-size cell B2, br in default-size cell D4
    ws.addImage(imageId, {tl: {col: 1.5, row: 1.5}, br: {col: 3.5, row: 3.25}});

    await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
    const ws2 = wb2.getWorksheet('Sheet1');
    const {tl, br} = ws2.getImages()[0].range;

    expect(tl.nativeCol).to.equal(1);
    expect(tl.nativeColOff).to.equal((20 * 7 * EMU_PER_PIXEL) / 2);
    expect(tl.nativeRow).to.equal(1);
    expect(tl.nativeRowOff).to.equal((40 * EMU_PER_POINT) / 2);

    expect(br.nativeCol).to.equal(3);
    expect(br.nativeColOff).to.equal((12 * 7 * EMU_PER_PIXEL) / 2);
    expect(br.nativeRow).to.equal(3);
    expect(br.nativeRowOff).to.equal((20 * EMU_PER_POINT) / 4);

    expect(tl.col).to.equal(1.5);
    expect(tl.row).to.equal(1.5);
    expect(br.col).to.equal(3.5);
    expect(br.row).to.equal(3.25);
  });
});
