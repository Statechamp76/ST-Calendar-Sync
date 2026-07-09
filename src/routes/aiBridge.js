const crypto = require('crypto');
const express = require('express');
const aiBridgeStore = require('../services/aiBridgeStore');

const ALLOWED_TYPES = new Set(['decision', 'prompt', 'runbook', 'note', 'incident', 'todo']);
const BLOCKED_KEYS = new Set(['attendees', 'body', 'location']);
const DEFAULT_LIMIT = 20;
const MAX_LIMIT = 100;
const MAX_CONTENT_LENGTH = 50_000;

function parseBoolean(value, defaultValue = false) {
  if (value === undefined || value === null || value === '') return defaultValue;
  return String(value).trim().toLowerCase() === 'true';
}

function safeCompareApiKey(expected, provided) {
  const a = Buffer.from(String(expected || ''));
  const b = Buffer.from(String(provided || ''));
  if (!a.length || a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

function hasBlockedKeyDeep(value) {
  if (!value || typeof value !== 'object') return false;
  if (Array.isArray(value)) {
    return value.some((item) => hasBlockedKeyDeep(item));
  }

  for (const key of Object.keys(value)) {
    if (BLOCKED_KEYS.has(String(key).toLowerCase())) {
      return true;
    }
    if (hasBlockedKeyDeep(value[key])) {
      return true;
    }
  }
  return false;
}

function hasSensitiveContentPattern(content) {
  const text = String(content || '');
  if (!text) return false;

  const patterns = [
    /"attendees"\s*:/i,
    /"location"\s*:/i,
    /"body"\s*:/i,
    /"emailAddress"\s*:/i,
  ];

  return patterns.some((pattern) => pattern.test(text));
}

function validateCreatePayload(payload) {
  if (!payload || typeof payload !== 'object') {
    return 'Request body must be a JSON object';
  }

  if (hasBlockedKeyDeep(payload)) {
    return 'Payload contains blocked keys';
  }

  const project = String(payload.project || '').trim();
  const type = String(payload.type || '').trim().toLowerCase();
  const title = String(payload.title || '').trim();
  const content = String(payload.content || '');

  if (!project || !type || !title || !content) {
    return 'project, type, title, and content are required';
  }
  if (!ALLOWED_TYPES.has(type)) {
    return 'Invalid type';
  }
  if (content.length > MAX_CONTENT_LENGTH) {
    return `content exceeds max length ${MAX_CONTENT_LENGTH}`;
  }
  if (hasSensitiveContentPattern(content)) {
    return 'content appears to contain raw payload fields';
  }

  if (payload.tags !== undefined && !Array.isArray(payload.tags)) {
    return 'tags must be an array';
  }

  return null;
}

function normalizeCreatePayload(payload) {
  const project = String(payload.project || '').trim();
  const type = String(payload.type || '').trim().toLowerCase();
  const title = String(payload.title || '').trim();
  const content = String(payload.content || '');
  const tags = Array.isArray(payload.tags)
    ? payload.tags.map((tag) => String(tag).trim()).filter(Boolean)
    : [];
  const source = String(payload.source || 'chatgpt').trim() || 'chatgpt';
  const hash = crypto
    .createHash('sha256')
    .update(`${project}|${type}|${title}|${content}`)
    .digest('hex');

  return {
    project,
    type,
    title,
    tags,
    content,
    source,
    hash,
    redactionFlags: [],
  };
}

function parseLimit(value) {
  if (value === undefined || value === null || value === '') return DEFAULT_LIMIT;
  const parsed = Number.parseInt(String(value), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) return DEFAULT_LIMIT;
  return Math.min(parsed, MAX_LIMIT);
}

function createAiBridgeRouter(deps = {}) {
  const router = express.Router();
  const store = deps.store || aiBridgeStore;

  router.use((req, res, next) => {
    const enabled = parseBoolean(process.env.AI_BRIDGE_ENABLED, false);
    if (!enabled) {
      res.status(404).json({ error: 'ai_bridge_disabled' });
      return;
    }

    const expectedApiKey = String(process.env.AI_BRIDGE_API_KEY || '');
    if (!expectedApiKey) {
      console.error('ai.bridge.misconfigured', { reason: 'missing_api_key_env' });
      res.status(503).json({ error: 'ai_bridge_unconfigured' });
      return;
    }

    const providedApiKey = req.get('X-API-Key');
    if (!safeCompareApiKey(expectedApiKey, providedApiKey)) {
      res.status(401).json({ error: 'unauthorized' });
      return;
    }

    next();
  });

  router.post('/entries', async (req, res) => {
    const validationError = validateCreatePayload(req.body);
    if (validationError) {
      res.status(400).json({ error: validationError });
      return;
    }

    const payload = normalizeCreatePayload(req.body);

    try {
      const created = await store.addEntry(payload);
      console.log('ai.bridge.entry.created', {
        entryId: created.id,
        project: payload.project,
        type: payload.type,
        title: payload.title,
      });
      res.status(201).json(created);
    } catch (error) {
      console.error('ai.bridge.entry.create_failed', { message: error.message });
      res.status(500).json({ error: 'Failed to create entry' });
    }
  });

  router.get('/entries', async (req, res) => {
    const project = String(req.query.project || '').trim();
    if (!project) {
      res.status(400).json({ error: 'project query param is required' });
      return;
    }

    const limit = parseLimit(req.query.limit);
    const type = req.query.type ? String(req.query.type).trim().toLowerCase() : null;
    const tag = req.query.tag ? String(req.query.tag).trim() : null;

    if (type && !ALLOWED_TYPES.has(type)) {
      res.status(400).json({ error: 'Invalid type' });
      return;
    }

    try {
      const entries = await store.listEntries({ project, limit, type, tag });
      const latest = entries[0] || null;
      console.log('ai.bridge.entries.read', {
        entryId: latest ? latest.id : null,
        project,
        type: type || 'read',
        title: latest ? latest.title : null,
      });
      res.status(200).json({ project, count: entries.length, entries });
    } catch (error) {
      console.error('ai.bridge.entries.read_failed', { message: error.message, project, type });
      res.status(500).json({ error: 'Failed to read entries' });
    }
  });

  router.get('/state', async (req, res) => {
    const project = String(req.query.project || '').trim();
    if (!project) {
      res.status(400).json({ error: 'project query param is required' });
      return;
    }

    try {
      const state = await store.getState(project);
      const latest = state.latestTitles[0] || null;
      console.log('ai.bridge.state.read', {
        entryId: latest ? latest.id : null,
        project,
        type: 'state',
        title: latest ? latest.title : null,
      });
      res.status(200).json(state);
    } catch (error) {
      console.error('ai.bridge.state.read_failed', { message: error.message, project });
      res.status(500).json({ error: 'Failed to read state' });
    }
  });

  return router;
}

module.exports = {
  ALLOWED_TYPES,
  MAX_CONTENT_LENGTH,
  createAiBridgeRouter,
  validateCreatePayload,
  hasBlockedKeyDeep,
};
