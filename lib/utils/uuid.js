'use strict';

const crypto = require('crypto');

/**
 * Generate an RFC 4122 version 4 UUID.
 * Uses crypto.randomUUID when available (Node >= 14.17, modern browsers)
 * and falls back to crypto.randomBytes otherwise.
 * @returns {string} e.g. '1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed'
 */
function uuidv4() {
  if (typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  const bytes = crypto.randomBytes(16);
  bytes[6] = 0x40 + (bytes[6] % 0x10); // version 4
  bytes[8] = 0x80 + (bytes[8] % 0x40); // variant 10xx
  const hex = bytes.toString('hex');
  return [
    hex.slice(0, 8),
    hex.slice(8, 12),
    hex.slice(12, 16),
    hex.slice(16, 20),
    hex.slice(20, 32),
  ].join('-');
}

module.exports = {uuidv4};
