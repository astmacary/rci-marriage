const { test } = require("node:test");
const assert = require("node:assert/strict");
const { generateKeyPairSync, webcrypto } = require("node:crypto");
const { project, decryptNote } = require("../scripts/pilot-sync.cjs");

const couple = { id: "42", data: {
  kind: "couple", journalItemId: "7", members: [{ objectId: "member-a", name: "Sample A" }, { objectId: "member-b", name: "Sample B" }]
}};
const row = (kind, date, values = {}, id = "member-a", title = "") => ({
  id: "1", createdBy: { user: { id } }, createdDateTime: date, lastModifiedDateTime: date,
  fields: { Title: title, AppData: JSON.stringify({ kind, season: "2027", ...values }) }
});

test("weekly and orchard projections use server authorship and Eastern dates and omit writing", () => {
  const answer = row("prompt-answer", "2027-02-01T03:00:00Z", { answer: "Example private text", date: "2027-06-01" });
  const checkin = row("check-in", "2027-02-01T06:00:00Z", { need: "Example check-in text" }, "member-b");
  const outsider = row("prompt-answer", "2027-02-01T06:00:00Z", {}, "not-on-roster");
  const result = project([couple], [answer, outsider], [checkin], [], [], Date.parse("2027-02-03"), "");
  assert.deepEqual([...result.participation.values()].map(r => [r.week, r.memberId, r.answered, r.checkedIn]), [
    ["2027-01-25", "member-a", true, false], ["2027-02-01", "member-b", false, true]
  ]);
  assert.deepEqual([...result.journals.get("42").months], ["2027-01", "2027-02"]);
  assert.doesNotMatch(JSON.stringify([...result.participation.values()]), /Example private|Example check-in/);
});

test("month-end release uses Eastern time and keeps future months sealed", () => {
  const { monthlyRelease } = require("../src/data.js");
  assert.equal(new Date(monthlyRelease("2027-01")).toISOString(), "2027-02-01T04:59:00.000Z");
  assert.equal(new Date(monthlyRelease("2027-02")).toISOString(), "2027-03-01T04:59:00.000Z");
  assert.equal(new Date(monthlyRelease("2027-03")).toISOString(), "2027-04-01T03:59:00.000Z");
  const note = row("sealed-note", "2027-01-29T15:00:00Z", { month: "2027-01" }, "member-a", "sealed:2027-01");
  const feb = row("sealed-note", "2027-02-02T15:00:00Z", { month: "2027-02" }, "member-a", "sealed:2027-02");
  const late = row("sealed-note", "2027-02-01T04:59:00Z", { month: "2027-01" }, "member-b", "sealed:2027-01");
  const notes = now => project([couple], [], [], [note, late, feb], [], now).journals.get("42").notes.size;
  assert.equal(notes(monthlyRelease("2027-01") - 1), 0);
  assert.equal(notes(monthlyRelease("2027-01")), 1);
  assert.equal(notes(monthlyRelease("2027-02")), 2);
});

test("the browser encryption envelope decrypts only with the release key", async () => {
  const pair = generateKeyPairSync("rsa", { modulusLength: 2048 });
  const publicKey = await webcrypto.subtle.importKey("jwk", pair.publicKey.export({ format: "jwk" }), { name: "RSA-OAEP", hash: "SHA-256" }, false, ["wrapKey"]);
  const key = await webcrypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt"]);
  const iv = webcrypto.getRandomValues(new Uint8Array(12));
  const ciphertext = await webcrypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode("Example note"));
  const wrappedKey = await webcrypto.subtle.wrapKey("raw", key, publicKey, "RSA-OAEP");
  const encode = value => Buffer.from(value).toString("base64");
  const data = { iv: encode(iv), ciphertext: encode(ciphertext), wrappedKey: encode(wrappedKey) };
  assert.equal(decryptNote(data, pair.privateKey), "Example note");
  assert.throws(() => decryptNote({ ...data, iv: Buffer.alloc(12).toString("base64") }, pair.privateKey));
});

test("RSVP updates take the newest response for a published event and never count as participation", () => {
  const event = { kind: "event", season: "2027", event: "retreat" };
  const old = row("rsvp", "2026-12-02T10:00:00Z", { event: "retreat", response: "yes" });
  const updated = row("rsvp", "2026-12-03T10:00:00Z", { event: "retreat", response: "no" });
  const result = project([couple], [], [], [updated, old], [event], Date.parse("2026-12-04"), "");
  assert.equal(result.participation.size, 0);
  assert.deepEqual([...result.rsvps.values()].map(record => record.response), ["no"]);
});
