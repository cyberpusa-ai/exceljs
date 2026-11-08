const parseSax = verquire('utils/parse-sax');

const parseText = async chunks => {
  const texts = [];
  for await (const events of parseSax(chunks)) {
    events.forEach(({eventType, value}) => {
      if (eventType === 'text') {
        texts.push(value);
      }
    });
  }
  return texts.join('');
};

// feed <t>text</t> to parseSax split at every possible pair of byte offsets
const utf8SplitTest = async text => {
  const bytes = Buffer.from(`<t>${text}</t>`, 'utf8');
  const results = [];
  for (let i = 1; i < bytes.length; i++) {
    for (let j = i; j < bytes.length; j++) {
      const chunks = [bytes.subarray(0, i)];
      if (j !== i) {
        chunks.push(bytes.subarray(i, j));
      }
      chunks.push(bytes.subarray(j));
      results.push(parseText(chunks));
    }
  }
  for (let len = 1; len < bytes.length; len++) {
    const chunks = [];
    for (let p = 0; p < bytes.length; p += len) {
      chunks.push(bytes.subarray(p, p + len));
    }
    results.push(parseText(chunks));
  }
  (await Promise.all(results)).forEach(result => expect(result).to.equal(text));
};

describe('parse-sax', () => {
  it('decodes single byte UTF-8 chars split across chunks', async () => {
    await utf8SplitTest('abcdefgh');
  });
  it('decodes 2 byte UTF-8 chars split across chunks', async () => {
    await utf8SplitTest('αβγδε');
  });
  it('decodes 3 byte UTF-8 chars split across chunks', async () => {
    await utf8SplitTest('あいうえお');
  });
  it('decodes 4 byte UTF-8 chars split across chunks', async () => {
    await utf8SplitTest('😀😁😂');
  });
  it('decodes Uint8Array chunks', async () => {
    const bytes = new Uint8Array(Buffer.from('<t>あい</t>', 'utf8'));
    expect(await parseText([bytes.subarray(0, 5), bytes.subarray(5)])).to.equal('あい');
  });
  it('accepts string chunks', async () => {
    expect(await parseText(['<t>あ', 'い</t>'])).to.equal('あい');
  });
});
