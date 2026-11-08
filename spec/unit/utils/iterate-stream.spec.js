const {Duplex, PassThrough} = require('stream');

const iterateStream = verquire('utils/iterate-stream');

async function collect(stream) {
  const items = [];
  for await (const item of iterateStream(stream)) {
    items.push(item);
  }
  return items;
}

describe('iterateStream', () => {
  it('yields every chunk of a stream', async () => {
    const stream = new PassThrough({objectMode: true});
    stream.write('a');
    stream.write('b');
    stream.end('c');
    expect(await collect(stream)).to.deep.equal(['a', 'b', 'c']);
  });

  it('ignores a synthetic end event emitted before the stream has really ended', async () => {
    // mimics unzipper's Parse, which emits 'end' and 'close' on 'finish' (#683)
    const stream = new Duplex({
      objectMode: true,
      read() {},
      write(chunk, encoding, callback) {
        callback();
      },
    });
    stream.on('finish', () => {
      stream.emit('end');
      stream.emit('close');
    });

    stream.push('a');
    setTimeout(() => stream.end(), 5);
    setTimeout(() => {
      stream.push('b');
      stream.push(null);
    }, 20);

    expect(await collect(stream)).to.deep.equal(['a', 'b']);
  });
});
