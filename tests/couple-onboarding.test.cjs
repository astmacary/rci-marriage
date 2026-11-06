const test = require("node:test");
const assert = require("node:assert/strict");
const { createHash } = require("node:crypto");
const { requestData, run } = require("../scripts/couple-onboarding.cjs");

const spouses = [{ name: "Sample One", email: "first@example.test" }, { name: "Sample Two", email: "second@example.test" }];
const title = "onboard:" + createHash("sha256").update(spouses.map(p => p.email).join("\n")).digest("hex");
const item = (overrides = {}) => ({ Id: 7, AuthorId: 5, Title: title, AppData: JSON.stringify({
  kind: "couple-onboarding", version: 1, coupleName: "Sample Couple", spouses, status: "queued", ...overrides
}) });

test("onboarding validates two distinct emails and binds the request key to the pair", () => {
  assert.equal(requestData(item()).spouses.length, 2);
  assert.throws(() => requestData(item({ spouses: [spouses[0], spouses[0]] })), /Invalid couple details/);
  assert.throws(() => requestData(item({ coupleName: "../Other Folder" })), /Invalid couple details/);
  assert.throws(() => requestData({ ...item(), Title: "different" }), /Request identity changed/);
});

test("the worker ignores a request whose SharePoint author is not a leader", async t => {
  const oldEnv = { ...process.env };
  Object.assign(process.env, {
    RCI_ONBOARDING_ENABLED: "true", RCI_ONBOARDING_CLIENT_ID: "11111111-1111-4111-8111-111111111111",
    RCI_ONBOARDING_SITE: "https://example.sharepoint.com/sites/RCIMarriage", RCI_ONBOARDING_LEADER_IDS: "5",
    IDENTITY_ENDPOINT: "https://identity.example/token", IDENTITY_HEADER: "sample"
  });
  t.after(() => { for (const key of Object.keys(process.env)) if (!(key in oldEnv)) delete process.env[key]; Object.assign(process.env, oldEnv); });
  const calls = [];
  t.mock.method(globalThis, "fetch", async (input, options) => {
    const url = new URL(input);
    calls.push({ url, method: options.method || "GET" });
    if (url.hostname === "identity.example") return Response.json({ access_token: "sample" });
    assert.match(url.pathname, /Couple%20Onboarding/);
    return Response.json({ value: [{ ...item(), AuthorId: 99 }] });
  });
  await run(() => {});
  assert.equal(calls.length, 2);
  assert.ok(calls.every(call => call.method === "GET"));
  assert.ok(calls.every(call => call.url.hostname !== "graph.microsoft.com"));
});
