const { SecretManagerServiceClient } = require('@google-cloud/secret-manager');

const client = new SecretManagerServiceClient();

async function getProjectId() {
  // Cloud Run usually sets one of these; client.getProjectId() is the most reliable fallback.
  return (
    process.env.GOOGLE_CLOUD_PROJECT ||
    process.env.GCLOUD_PROJECT ||
    process.env.GCP_PROJECT ||
    await client.getProjectId()
  );
}

async function getSecrets(secretNames) {
  const secrets = {};
  const projectId = await getProjectId();

  for (const name of secretNames) {
    const secretPath = `projects/${projectId}/secrets/${name}/versions/latest`;
    try {
      const [version] = await client.accessSecretVersion({ name: secretPath });
      secrets[name] = version.payload.data.toString('utf8').trim();
    } catch (error) {
      console.error(`Failed to access secret: ${name}. Path: ${secretPath}. Error: ${error.message}`);
      throw new Error(`Could not retrieve secret: ${name}`);
    }
  }
  return secrets;
}

module.exports = { getSecrets };