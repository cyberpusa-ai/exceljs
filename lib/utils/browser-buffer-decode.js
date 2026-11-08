// eslint-disable-next-line node/no-unsupported-features/node-builtins
const textDecoder = typeof TextDecoder === 'undefined' ? null : new TextDecoder('utf-8');

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  if (textDecoder) {
    return textDecoder.decode(chunk);
  }
  return chunk.toString();
}

// Returns a decoder for a stream of byte chunks. Unlike bufferToString, it keeps
// the trailing bytes of a multi-byte UTF-8 character that is split across two
// chunks and decodes them together with the next chunk.
function createStreamDecoder() {
  if (textDecoder) {
    // eslint-disable-next-line node/no-unsupported-features/node-builtins
    const decoder = new TextDecoder('utf-8');
    return {
      write: chunk => (typeof chunk === 'string' ? chunk : decoder.decode(chunk, {stream: true})),
      end: () => decoder.decode(),
    };
  }
  // eslint-disable-next-line global-require
  const {StringDecoder} = require('string_decoder');
  const decoder = new StringDecoder('utf8');
  return {
    write: chunk => (typeof chunk === 'string' ? chunk : decoder.write(chunk)),
    end: () => decoder.end(),
  };
}

exports.bufferToString = bufferToString;
exports.createStreamDecoder = createStreamDecoder;
