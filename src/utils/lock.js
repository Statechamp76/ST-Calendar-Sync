const crypto = require('node:crypto');

function makeLockHolderId() {
  return crypto.randomBytes(8).toString('hex');
}

module.exports = {
  makeLockHolderId,
};

