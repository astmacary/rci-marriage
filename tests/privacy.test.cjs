const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const D = require('../src/data.js');

test('private response selection uses the Microsoft object ID, not a supplied author field', () => {
  const items = [
    { id: '1', createdBy: { user: { id: 'person-a' } }, fields: { AppData: '{"author":"person-b"}' } },
    { id: '2', createdBy: { user: { id: 'person-b' } } },
    { id: '3', fields: { AppData: '{"author":"person-a"}' } }
  ];
  assert.deepEqual(D.own(items, 'PERSON-A').map(item => item.id), ['1']);
});
test('untrusted list text and links cannot become executable markup', () => {
  assert.equal(D.escape('<img src=x onerror="alert(1)">'), '&lt;img src=x onerror=&quot;alert(1)&quot;&gt;');
  assert.equal(D.safeLink('javascript:alert(1)'), '');
  assert.equal(D.safeLink('https://example.org/resource'), 'https://example.org/resource');
  assert.deepEqual(D.parse({ fields: { AppData: '<not json>' } }), {});
});
test('oversize responses cannot be silently truncated by SharePoint', () => {
  assert.throws(() => D.validate({ version: 1, kind: 'prompt-answer', answer: 'a'.repeat(25000) }));
});
test('demo never initializes Microsoft sign-in or fetches live data', async () => {
  const context = vm.createContext({ window: { RCI_APP: { DEMO: true } } });
  await vm.runInContext(fs.readFileSync('live.js', 'utf8'), context);
});
test('signed-out live mode shows sign-in, with no demo couple or answers', async () => {
  const root = { innerHTML: '', className: '' };
  const context = vm.createContext({
    window: { RCI_APP: { DEMO: false }, RCI_CONFIG: { clientId: '', siteHost: 'example.sharepoint.com', sitePath: '/sites/Test' }, RCIData: D },
    document: { getElementById: () => root, addEventListener: () => {} }, AbortController, URL, Date
  });
  await vm.runInContext(fs.readFileSync('live.js', 'utf8'), context);
  assert.match(root.innerHTML, /Live sign-in is being configured/);
  assert.doesNotMatch(root.innerHTML, /Marcus|Tasha|Both answered|30 couples/);
});
test('repository source has no personal couple identities, secrets or keyed flow endpoints', () => {
  const files = ['config.js', 'live.js', 'src/data.js', 'index.html'];
  for (const file of files) {
    const source = fs.readFileSync(file, 'utf8');
    assert.doesNotMatch(source, /[A-Z0-9._%+-]+@therampchurch\.com/i, file);
    assert.doesNotMatch(source, /client_secret\s*[:=]\s*["'][^"']+/i, file);
    assert.doesNotMatch(source, /https?:[^\s"']*[?&](sig|code|key)=/i, file);
  }
});

async function liveHarness({ deniedAnswers = false } = {}) {
  const root = { innerHTML: '', className: '' }, listeners = {}, calls = [];
  const names = { couples: 'Couples', answers: 'Prompt Answers', reveals: 'Couple Journal', checkins: 'Sunday Check-ins', participation: 'Participation', content: 'Monthly Content', rsvps: 'RSVPs' };
  const account = { name: 'Sample Person', username: 'sample@example.org', localAccountId: 'person-a', tenantId: 'tenant', idTokenClaims: { oid: 'person-a' } };
  class Auth {
    async initialize() {} async handleRedirectPromise() { return { account }; }
    setActiveAccount() {} async acquireTokenSilent() { return { accessToken: 'test-token' }; }
  }
  const context = vm.createContext({
    window: { RCI_APP: { DEMO: false }, RCI_CONFIG: { clientId: 'public-id', tenantId: 'tenant', siteHost: 'example.sharepoint.com', sitePath: '/sites/Test', lists: names }, RCIData: D },
    document: { getElementById: () => root, addEventListener: (name, cb) => { listeners[name] = cb; } },
    AbortController, URL, Date, msal: { PublicClientApplication: Auth },
    FormData: class { constructor(form) { return form.values; } },
    fetch: async (url, options) => {
      const path = String(url); calls.push({ path, options });
      const result = body => new Response(JSON.stringify(body), { status: 200 });
      if (options.method) return new Response('', { status: 403 });
      if (path.endsWith('/sites/example.sharepoint.com:/sites/Test')) return result({ id: 'site' });
      if (path.includes('/lists?$select=')) return result({ value: Object.entries(names).map(([id, displayName]) => ({ id, displayName })) });
      if (path.includes('/lists/couples/items')) return result({ value: [{ id: '1', fields: { AppData: JSON.stringify({ kind: 'couple', memberObjectIds: ['person-a'], journalItemId: '9' }) } }] });
      if (path.includes('/lists/reveals/items/9?')) return result({ id: '9', eTag: 'revision-1', fields: { AppData: '{"entries":[]}' } });
      if (deniedAnswers && path.includes('/lists/answers/')) return new Response('', { status: 403 });
      return result({ value: [] });
    }
  });
  await vm.runInContext(fs.readFileSync('live.js', 'utf8'), context);
  const clickView = async name => listeners.click({ target: { closest: selector => selector === '[data-view]' ? { dataset: { view: name } } : null } });
  return { root, calls, listeners, clickView };
}
test('leader view requests only its four lists, never private writing', async () => {
  const app = await liveHarness(); app.calls.length = 0;
  await app.clickView('leaders');
  assert.equal(app.calls.length, 4);
  assert.ok(app.calls.every(({path}) => /\/lists\/(couples|participation|rsvps|content)\//.test(path)));
  assert.doesNotMatch(app.root.innerHTML, /30 couples|Marcus|Tasha/);
});
test('journal reads one permissioned folder, never the entire journal list', async () => {
  const app = await liveHarness(); app.calls.length = 0;
  await app.clickView('journal');
  const journals = app.calls.filter(({path}) => path.includes('/lists/reveals/'));
  assert.equal(journals.length, 1);
  assert.match(journals[0].path, /\/items\/9\?/);
});
test('denied reads remain an access error instead of becoming sample or empty data', async () => {
  const app = await liveHarness({ deniedAnswers: true });
  assert.match(app.root.innerHTML, /not granted your account access/);
  assert.doesNotMatch(app.root.innerHTML, /No saved responses yet/);
});
test('failed SharePoint writes preserve the form and never report a successful save', async () => {
  const app = await liveHarness();
  const button = { disabled: false }, status = { textContent: '' };
  const form = { dataset: { form: 'answer' }, values: new Map([['promptId','1'], ['answer','Fictional test response']]), querySelector: selector => selector === '[type="submit"]' ? button : status };
  await app.listeners.submit({ preventDefault() {}, target: { closest: () => form } });
  assert.match(status.textContent, /not granted your account access/);
  assert.equal(button.disabled, false);
  assert.doesNotMatch(app.root.innerHTML, /Saved to SharePoint\./);
});

test('answers and check-ins save to named SharePoint columns, not one text blob', () => {
  const a = D.columns({ version: 1, kind: 'prompt-answer', date: '2027-01-04', answer: 'Fictional' }, 'c1');
  assert.deepEqual(a, { PromptDate: '2027-01-04T12:00:00Z', Answer: 'Fictional', CoupleKey: 'c1' });
  const c = D.columns({ version: 1, kind: 'check-in', date: '2027-01-03', ratings: [4, 3, 9, 0], need: 'n', saw: 's', mentor: true }, 'c1');
  assert.equal(c.Connected, 4); assert.equal(c.Spiritual, null); assert.equal(c.Fun, null); assert.equal(c.MentorRequest, true);
  assert.equal('AppData' in a || 'AppData' in c, false);
});
test('rows read from columns first and still read older AppData rows', () => {
  const fromCols = D.read('content', { fields: { Title: 'What one word…?', Kind: 'Prompt', PromptType: 'Fun', PromptDate: '2027-01-01T12:00:00Z' } });
  assert.equal(fromCols.kind, 'prompt'); assert.equal(fromCols.date, '2027-01-01'); assert.equal(fromCols.text, 'What one word…?');
  const legacy = D.read('content', { fields: { Title: 'x', AppData: '{"kind":"prompt","date":"2027-01-02","text":"Old"}' } });
  assert.equal(legacy.text, 'Old');
  const couple = D.read('couples', { id: '7', fields: { Spouse1Email: 'a@example.org', Spouse2Email: 'b@example.org', CoupleKey: 'k7' } });
  assert.equal(couple.kind, 'couple'); assert.deepEqual(couple.memberEmails, ['a@example.org', 'b@example.org']); assert.equal(couple.coupleKey, 'k7');
});
test('couple journal link comes from the JournalItemId column', () => {
  assert.equal(D.read('couples', { fields: { Spouse1Email: 'a@example.org', JournalItemId: '12' } }).journalItemId, '12');
});
