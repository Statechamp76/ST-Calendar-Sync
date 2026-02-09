const { OAuth2Client } = require('google-auth-library');
const { loadConfig } = require('../config');

const config = loadConfig();
const client = new OAuth2Client();

async function requireOidcAuth(req, res, next) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';

  if (!token) {
    res.status(401).json({
      error: 'Unauthorized',
      message: 'Missing Bearer token',
    });
    return;
  }

  try {
    await client.verifyIdToken({
      idToken: token,
      audience: config.runSyncAudience,
    });
    next();
  } catch (error) {
    console.error('OIDC token verification failed', {
      message: error.message,
    });
    res.status(401).json({
      error: 'Unauthorized',
      message: 'Invalid identity token',
    });
  }
}

module.exports = {
  requireOidcAuth,
};
