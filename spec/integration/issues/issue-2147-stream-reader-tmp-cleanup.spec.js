const fs = require('fs');
const tmp = require('tmp');

const ExcelJS = verquire('exceljs');

// worksheet comes before sharedStrings in this file, so the reader buffers it to a tmp file
const TEST_XLSX_FILE_NAME = './spec/integration/data/test-row-styles.xlsx';

describe('github issues', () => {
  describe('issue 2147 - stream reader cleans up tmp files when not fully iterated', () => {
    let origTmpFile;
    let tmpFiles;

    beforeEach(() => {
      tmpFiles = [];
      origTmpFile = tmp.file;
      // record each tmp file and whether (and when) its cleanup callback ran
      tmp.file = cb =>
        origTmpFile((err, path, fd, cleanupCallback) => {
          const tmpFile = {path, called: false};
          tmpFile.removed = new Promise(resolve => {
            tmpFile.cleanup = () => {
              tmpFile.called = true;
              // tmp removes the file asynchronously; resolve once it is gone
              cleanupCallback(resolve);
            };
          });
          if (!err) tmpFiles.push(tmpFile);
          cb(err, path, fd, tmpFile.cleanup);
        });
    });

    afterEach(() => {
      tmp.file = origTmpFile;
    });

    const expectCleanedUp = async () => {
      expect(tmpFiles.length).to.be.greaterThan(0);
      tmpFiles.forEach(({called}) => expect(called).to.equal(true));
      await Promise.all(tmpFiles.map(({removed}) => removed));
      tmpFiles.forEach(({path}) =>
        expect(fs.existsSync(path)).to.equal(false)
      );
    };

    it('when breaking out of the iteration', async () => {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        fs.createReadStream(TEST_XLSX_FILE_NAME)
      );
      // eslint-disable-next-line no-unused-vars
      for await (const worksheet of workbookReader) {
        break;
      }
      await expectCleanedUp();
    });

    it('when an error is thrown during the iteration', async () => {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        fs.createReadStream(TEST_XLSX_FILE_NAME)
      );
      let caught;
      try {
        // eslint-disable-next-line no-unused-vars
        for await (const worksheet of workbookReader) {
          throw new Error('stop');
        }
      } catch (error) {
        caught = error;
      }
      expect(caught.message).to.equal('stop');
      await expectCleanedUp();
    });

    it('when fully iterated', async () => {
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        fs.createReadStream(TEST_XLSX_FILE_NAME)
      );
      let rows = 0;
      for await (const worksheet of workbookReader) {
        for await (const row of worksheet) {
          if (row) rows++;
        }
      }
      expect(rows).to.be.greaterThan(0);
      await expectCleanedUp();
    });
  });
});
