const {PassThrough} = require('readable-stream');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('pull request 1665 - multi-byte chars split across stream chunks', () => {
    it('WorkbookReader keeps CJK and emoji text intact', async () => {
      // long enough that the inflated xml arrives in many chunks,
      // so some chunk boundaries fall inside a multi-byte character
      const text = 'あいうえお漢字😀'.repeat(20000);
      const output = new PassThrough();
      const buffers = [];
      output.on('data', chunk => buffers.push(chunk));
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({
        stream: output,
        useSharedStrings: true,
      });
      const ws = writer.addWorksheet('Sheet1');
      ws.addRow([text]).commit();
      ws.addRow([{richText: [{text}]}]).commit();
      await writer.commit();

      const input = new PassThrough();
      input.end(Buffer.concat(buffers));
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(input);
      const values = [];
      for await (const worksheetReader of reader) {
        for await (const row of worksheetReader) {
          values.push(row.getCell(1).value);
        }
      }
      expect(values[0]).to.equal(text);
      expect(values[1].richText[0].text).to.equal(text);
    });
  });
});
