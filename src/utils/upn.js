function normalizeUpn(value) {
  return String(value || '').trim().toLowerCase();
}

module.exports = {
  normalizeUpn,
};
