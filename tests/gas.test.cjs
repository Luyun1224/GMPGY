const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function backend(dashboardPassword = '') {
  const roster = [['2026-10', '測試學員', 'PGY1']];
  const responses = [['受評學員', '狀態'], ['測試學員', '已評核']];
  const settingsRead = [];
  let writes = 0;
  const rosterSheet = {
    getLastRow: () => roster.length + 1,
    getRange: () => ({
      getValues: () => roster,
      clearContent: () => { writes++; },
      setNumberFormat() { return this; },
      setValues: () => { writes++; },
    }),
  };
  const sheets = {
    '設定': { getRange(row, col) {
      assert.equal(col, 2);
      settingsRead.push(row);
      return { getDisplayValue: () => ({1: '20', 2: '', 3: dashboardPassword, 4: 'test-admin'})[row] };
    } },
    'PGY名單': rosterSheet,
    '回覆': { getLastRow: () => responses.length, getDataRange: () => ({ getDisplayValues: () => responses }) },
  };
  const context = vm.createContext({
    SpreadsheetApp: { getActive: () => ({ getSheetByName: name => sheets[name] }) },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: text => ({ setMimeType: () => JSON.parse(text) }),
    },
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/PGY評核_Code.gs'), 'utf8'), context);
  return {
    get: parameter => context.doGet({ parameter }),
    post: body => context.doPost({ postData: { contents: JSON.stringify(body) } }),
    settingsRead,
    writes: () => writes,
  };
}

for (const password of ['', 'old-dashboard-password']) {
  test(`dashboard reads data without key when B3 is ${password ? 'populated' : 'empty'}`, () => {
    const api = backend(password);
    const result = api.get({ action: 'data' });
    assert.equal(result.ok, true);
    assert.deepEqual(result.headers, ['受評學員', '狀態']);
    assert.deepEqual(result.rows, [['測試學員', '已評核']]);
    assert.equal(result.roster['2026-10'][0].name, '測試學員');
    assert.equal(api.settingsRead.includes(3), false);
    assert.equal(api.writes(), 0);
  });
}

test('questionnaire configuration stays public', () => {
  const result = backend().get({ action: 'config' });
  assert.equal(result.ok, true);
  assert.equal(result.openDay, 20);
  assert.equal(result.roster['2026-10'][0].level, 'PGY1');
});

for (const action of ['adminLogin', 'adminSave']) {
  for (const key of ['', 'incorrect']) {
    test(`${action} rejects ${key ? 'incorrect' : 'missing'} administrator password`, () => {
      const api = backend();
      assert.deepEqual(api.post({ action, key, roster: [] }), { ok: false, error: 'auth' });
      assert.equal(api.writes(), 0);
    });
  }
  test(`${action} accepts the B4 administrator password`, () => {
    const api = backend();
    const result = api.post({ action, key: 'test-admin', roster: [{ month: '2026-10', name: '測試學員', level: 'PGY1' }] });
    assert.equal(result.ok, true);
    assert.equal(result.roster[0].name, '測試學員');
    assert.equal(api.writes() > 0, action === 'adminSave');
    assert.deepEqual(api.settingsRead, [4]);
  });
}

test('unknown actions still fail', () => {
  assert.deepEqual(backend().get({ action: 'other' }), { ok: false, error: 'unknown action' });
});
