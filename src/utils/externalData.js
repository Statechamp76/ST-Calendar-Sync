function extractIntegrationExternalId(detail, applicationGuid) {
  const guid = String(applicationGuid || '').trim().toLowerCase();
  if (!detail || !guid) return null;

  const externalData = detail.externalData;
  const candidates = [];

  if (externalData && typeof externalData === 'object' && !Array.isArray(externalData)) {
    candidates.push(externalData);
  }
  if (Array.isArray(externalData)) {
    candidates.push(...externalData.filter(Boolean));
  }

  for (const item of candidates) {
    const ag = String(item.applicationGuid || item.applicationGUID || '').trim().toLowerCase();
    if (ag && ag !== guid) continue;
    const ext = item.externalId || item.externalID || item.externalKey || item.externalValue;
    if (ext) return String(ext);
  }

  return null;
}

module.exports = {
  extractIntegrationExternalId,
};

