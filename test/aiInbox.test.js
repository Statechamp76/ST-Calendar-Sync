const test = require('node:test');
const assert = require('node:assert/strict');
const {
  parseInboxRows,
  readInbox,
  clearInboxRows,
  processInbox,
  SPREADSHEET_ID,
} = require('../src/services/aiInbox');

function unlockedFirestoreStub() {
  const state = {};
  return {
    collection: () => ({
      doc: () => ({ __state: state }),
    }),
    runTransaction: async (fn) => {
      const tx = {
        get: async (ref) => ({
          exists: Boolean(ref.__state.doc),
          data: () => ref.__state.doc || {},
        }),
        set: (ref, payload) => {
          ref.__state.doc = { ...(ref.__state.doc || {}), ...payload };
        },
      };
      return fn(tx);
    },
  };
}

test('readInbox parses rows with non-empty id and matching project', async () => {
  const rawRows = [
    ['123', '2026-02-12T00:00:00Z', 'ST-Calendar-Sync', 'note', 'A', 'ops,sync', 'content A', 'chatgpt'],
    ['', '2026-02-12T00:01:00Z', 'ST-Calendar-Sync', 'note', 'B', '', 'content B', 'chatgpt'],
    ['456', '2026-02-12T00:02:00Z', 'OtherProject', 'note', 'C', '', 'content C', 'chatgpt'],
    ['789', '2026-02-12T00:03:00Z', 'ST-Calendar-Sync', 'todo', 'D', '["x","y"]', 'content D', 'codex'],
  ];

  const mockSheets = {
    readSheetRows: async (range, spreadsheetId) => {
      assert.equal(range, 'AI_INBOX!A2:H');
      assert.equal(spreadsheetId, SPREADSHEET_ID);
      return rawRows;
    },
  };

  const entries = await readInbox('ST-Calendar-Sync', 50, { sheets: mockSheets });
  assert.equal(entries.length, 2);
  assert.equal(entries[0].id, '123');
  assert.equal(entries[0].rowIndex, 2);
  assert.deepEqual(entries[0].tags, ['ops', 'sync']);
  assert.equal(entries[1].id, '789');
  assert.equal(entries[1].rowIndex, 5);
  assert.deepEqual(entries[1].tags, ['x', 'y']);
});

test('parseInboxRows enforces limit', () => {
  const rawRows = [
    ['1', '', 'ST-Calendar-Sync', '', '', '', '', ''],
    ['2', '', 'ST-Calendar-Sync', '', '', '', '', ''],
    ['3', '', 'ST-Calendar-Sync', '', '', '', '', ''],
  ];
  const parsed = parseInboxRows(rawRows, { project: 'ST-Calendar-Sync', limit: 2 });
  assert.equal(parsed.length, 2);
  assert.equal(parsed[0].id, '1');
  assert.equal(parsed[1].id, '2');
});

test('clearInboxRows blanks expected row ranges in one batch call', async () => {
  let seenData = null;
  let seenSpreadsheetId = null;

  const mockSheets = {
    batchUpdateSheetRanges: async (data, spreadsheetId) => {
      seenData = data;
      seenSpreadsheetId = spreadsheetId;
    },
  };

  await clearInboxRows(
    [
      { rowIndex: 2, id: 'a' },
      { rowIndex: 5, id: 'b' },
    ],
    { sheets: mockSheets },
  );

  assert.equal(seenSpreadsheetId, SPREADSHEET_ID);
  assert.equal(Array.isArray(seenData), true);
  assert.equal(seenData.length, 2);
  assert.equal(seenData[0].range, 'AI_INBOX!A2:H2');
  assert.equal(seenData[1].range, 'AI_INBOX!A5:H5');
  assert.deepEqual(seenData[0].values, [['', '', '', '', '', '', '', '']]);
  assert.deepEqual(seenData[1].values, [['', '', '', '', '', '', '', '']]);
});

test('processInbox does not clear inbox when outbox append fails', async () => {
  let clearCalled = 0;
  const mockSheets = {
    readSheetRows: async () => ([
      ['in-1', '2026-02-12T00:00:00Z', 'ST-Calendar-Sync', 'note', 'Title', '', 'content', 'chatgpt'],
    ]),
    appendSheetRows: async () => {
      throw new Error('append_failed');
    },
    batchUpdateSheetRanges: async () => {
      clearCalled += 1;
    },
  };

  await assert.rejects(
    async () => processInbox(
      { project: 'ST-Calendar-Sync', limit: 50 },
      { sheets: mockSheets, firestore: unlockedFirestoreStub() },
    ),
    /append_failed/,
  );
  assert.equal(clearCalled, 0);
});

test('processInbox writes outbox error row for malformed inbox item and still clears it', async () => {
  let appended = null;
  let cleared = null;
  const mockSheets = {
    readSheetRows: async () => ([
      ['in-2', '2026-02-12T00:00:00Z', 'ST-Calendar-Sync', 'note', 'Broken Row', '', '', 'chatgpt'],
    ]),
    appendSheetRows: async (range, rows) => {
      assert.equal(range, 'AI_OUTBOX!A:G');
      appended = rows;
    },
    batchUpdateSheetRanges: async (ranges) => {
      cleared = ranges;
    },
  };

  const summary = await processInbox(
    { project: 'ST-Calendar-Sync', limit: 50 },
    { sheets: mockSheets, firestore: unlockedFirestoreStub() },
  );

  assert.equal(summary.read, 1);
  assert.equal(summary.outboxWritten, 1);
  assert.equal(summary.cleared, 1);
  assert.equal(Array.isArray(appended), true);
  assert.equal(appended.length, 1);
  assert.equal(appended[0][6], 'in-2'); // related_inbox_id
  assert.equal(String(appended[0][2]), 'ST-Calendar-Sync');
  assert.equal(String(appended[0][3]), 'incident');
  assert.match(String(appended[0][5]), /status=error/);
  assert.equal(Array.isArray(cleared), true);
  assert.equal(cleared.length, 1);
  assert.equal(cleared[0].range, 'AI_INBOX!A2:H2');
});
