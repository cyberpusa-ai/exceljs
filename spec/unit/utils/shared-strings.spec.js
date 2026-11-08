const SharedStrings = verquire('utils/shared-strings');

describe('SharedStrings', () => {
  it('Stores and shares string values', () => {
    const ss = new SharedStrings();

    const iHello = ss.add('Hello');
    const iHelloV2 = ss.add('Hello');
    const iGoodbye = ss.add('Goodbye');

    expect(iHello).to.equal(iHelloV2);
    expect(iGoodbye).to.not.equal(iHelloV2);

    expect(ss.count).to.equal(2);
    expect(ss.totalRefs).to.equal(3);
  });

  it('Does not escape values', () => {
    // that's the job of the xml utils
    const ss = new SharedStrings();

    const iXml = ss.add('<tag>value</tag>');
    const iAmpersand = ss.add('&');

    expect(ss.getString(iXml)).to.equal('<tag>value</tag>');
    expect(ss.getString(iAmpersand)).to.equal('&');
  });

  it('Distinguishes richText values', () => {
    const ss = new SharedStrings();

    const iRich1 = ss.add({richText: [{text: 'value 1'}]});
    const iRich2 = ss.add({richText: [{text: 'value 2'}]});
    const iRich1V2 = ss.add({richText: [{text: 'value 1'}]});
    const iBold = ss.add({richText: [{font: {bold: true}, text: 'value 1'}]});

    expect(iRich1).to.not.equal(iRich2);
    expect(iRich1).to.equal(iRich1V2);
    expect(iBold).to.not.equal(iRich1);
    expect(ss.getString(iRich2)).to.deep.equal({richText: [{text: 'value 2'}]});

    expect(ss.count).to.equal(3);
    expect(ss.totalRefs).to.equal(4);
  });

  it('Does not mix richText with plain strings', () => {
    const ss = new SharedStrings();
    const rich = {richText: [{text: 'value'}]};

    const iJson = ss.add(JSON.stringify(rich));
    const iObject = ss.add('[object Object]');
    const iRich = ss.add(rich);

    expect(iRich).to.not.equal(iJson);
    expect(iRich).to.not.equal(iObject);
    expect(ss.count).to.equal(3);
  });
});
