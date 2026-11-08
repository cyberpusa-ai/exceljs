const crypto = require('crypto');

const {uuidv4} = verquire('utils/uuid');

const UUID_V4 = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

describe('uuid', () => {
  describe('uuidv4', () => {
    it('generates a version 4 UUID', () => {
      expect(uuidv4()).to.match(UUID_V4);
    });

    it('generates unique values', () => {
      const values = new Set();
      for (let i = 0; i < 100; i++) {
        values.add(uuidv4());
      }
      expect(values.size).to.equal(100);
    });

    it('falls back to randomBytes when randomUUID is not available', () => {
      const {randomUUID} = crypto;
      crypto.randomUUID = undefined;
      try {
        for (let i = 0; i < 100; i++) {
          expect(uuidv4()).to.match(UUID_V4);
        }
      } finally {
        crypto.randomUUID = randomUUID;
      }
    });
  });
});
