async function getSecrets(secretNames) {
  const secrets = {};

  for (const name of secretNames) {
    const value = process.env[name];
    if (!value) {
      throw new Error(`Missing required environment variable: ${name}`);
    }
    secrets[name] = value.trim();
  }

  return secrets;
}

module.exports = { getSecrets };
