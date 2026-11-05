// Run on a private scheduler with a managed identity. Never run in the browser.
const { privateDecrypt, createDecipheriv, constants } = require("node:crypto");
const { parse, read: readRow, ministryDay, weekStart, monthlyRelease, instant, events } = require("../src/data.js");

const origin = "https://graph.microsoft.com";
const season = "2027";
const lower = value => String(value || "").toLowerCase();

function project(roster, answers, checkins, intake, published, now) {
  const couples = roster.filter(row => row.data.kind === "couple" && row.data.active !== false);
  const members = new Map();
  const folders = new Set();
  for (const couple of couples) {
    if (couple.data.journalItemId) {
      const folder = String(couple.data.journalItemId);
      if (!/^\d+$/.test(folder) || folders.has(folder)) throw new Error("Invalid or duplicate journal folder");
      folders.add(folder);
    }
    for (const member of couple.data.members || []) {
      const id = lower(member.objectId);
      if (!id || members.has(id)) throw new Error("Invalid or duplicate roster member");
      members.set(id, { couple, name: member.name });
    }
  }
  const participation = new Map();
  const rsvps = new Map();
  const journals = new Map();
  const eventKeys = new Set(published.filter(row => row.kind === "event" && row.season === season).map(row => row.event));

  function record(row, source) {
    const memberId = lower(row.createdBy?.user?.id);
    const member = members.get(memberId);
    const created = instant(row.createdDateTime);
    const updated = instant(row.lastModifiedDateTime || row.createdDateTime);
    if (!member || !Number.isFinite(created) || !Number.isFinite(updated) || updated > now) return;
    const date = ministryDay(new Date(created));
    const { couple, name } = member, data = readRow(source, row);
    const month = date.slice(0, 7);
    const journal = journals.get(couple.id) || { months: new Set(), notes: new Map(), folder: couple.data.journalItemId };
    journals.set(couple.id, journal);
    if (date.startsWith(season + "-") && ((source === "answers" && data.kind === "prompt-answer") ||
        (source === "checkins" && data.kind === "check-in"))) {
      const week = weekStart(date), key = `week:${week}:${couple.id}:${memberId}`;
      const flags = participation.get(key) || { version: 1, kind: "participation", week, coupleId: couple.id, memberId, answered: false, checkedIn: false };
      if (source === "answers") flags.answered = true;
      else flags.checkedIn = true;
      participation.set(key, flags);
      journal.months.add(month);
    }
    if (source === "pilotIntake" && data.kind === "rsvp" && data.season === season &&
      events.includes(data.event) && eventKeys.has(data.event) && ["yes", "no"].includes(data.response)) {
      const key = `rsvp:${season}:${data.event}:${couple.id}:${memberId}`;
      if (!rsvps.has(key) || updated > rsvps.get(key).updated) rsvps.set(key, {
        version: 1, kind: "rsvp", event: data.event, season, coupleId: couple.id, memberId, response: data.response, updated
      });
    }
    const releaseTime = monthlyRelease(month);
    if (source === "pilotIntake" && now >= releaseTime && data.kind === "sealed-note" && data.season === season &&
        data.month === month && row.fields?.Title === `sealed:${month}` && date >= "2027-01-29" && date.startsWith(season + "-") && updated < releaseTime) {
      const key = `${month}:${memberId}`;
      const previous = journal.notes.get(key);
      if (!previous || created < previous.created) journal.notes.set(key, { month, memberId, author: name, data, created });
    }
  }
  answers.forEach(row => record(row, "answers"));
  checkins.forEach(row => record(row, "checkins"));
  intake.forEach(row => record(row, "pilotIntake"));
  return { participation, rsvps, journals };
}

function decryptNote(data, privateKey) {
  const key = privateDecrypt({ key: privateKey, padding: constants.RSA_PKCS1_OAEP_PADDING, oaepHash: "sha256" }, Buffer.from(data.wrappedKey, "base64"));
  const payload = Buffer.from(data.ciphertext, "base64"), iv = Buffer.from(data.iv, "base64");
  if (key.length !== 32 || iv.length !== 12 || payload.length < 17 || payload.length > 8016) throw new Error("Invalid sealed note");
  const cipher = createDecipheriv("aes-256-gcm", key, iv);
  cipher.setAuthTag(payload.subarray(-16));
  return Buffer.concat([cipher.update(payload.subarray(0, -16)), cipher.final()]).toString("utf8");
}

async function managedToken() {
  const { IDENTITY_ENDPOINT, IDENTITY_HEADER } = process.env;
  if (!IDENTITY_ENDPOINT || !IDENTITY_HEADER) throw new Error("Managed identity is required");
  const url = new URL(IDENTITY_ENDPOINT);
  url.searchParams.set("resource", origin);
  url.searchParams.set("api-version", "2019-08-01");
  const response = await fetch(url, { redirect: "error", headers: { "X-IDENTITY-HEADER": IDENTITY_HEADER } });
  if (!response.ok) throw new Error(`Managed identity failed (${response.status})`);
  return (await response.json()).access_token;
}

async function run(log = console.log) {
  const { RCI_SITE_ID, RCI_SEALED_PRIVATE_KEY } = process.env;
  if (!RCI_SITE_ID || !/^[a-zA-Z0-9.,-]+$/.test(RCI_SITE_ID)) throw new Error("Set RCI_SITE_ID");
  const token = await managedToken();
  const site = `/v1.0/sites/${RCI_SITE_ID}`;
  async function request(path, method = "GET", body, etag) {
    const url = new URL(path, origin);
    if (url.origin !== origin || !url.pathname.startsWith(site + "/")) throw new Error("Unexpected Graph destination");
    const response = await fetch(url, {
      method, redirect: "error", headers: {
        Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(etag ? { "If-Match": etag } : {})
      }, body: body && JSON.stringify(body)
    });
    if (!response.ok) throw new Error(`Graph ${method} failed (${response.status})`);
    return response.status === 204 ? null : response.json();
  }
  async function pages(path) {
    const items = [];
    while (path) {
      const page = await request(path);
      items.push(...page.value);
      path = page["@odata.nextLink"];
    }
    return items;
  }
  const lists = await pages(site + "/lists?$select=id,displayName");
  const path = name => {
    const list = lists.find(item => item.displayName === name);
    if (!list) throw new Error("Missing list " + name);
    return `${site}/lists/${list.id}/items`;
  };
  const read = name => pages(path(name) + "?$expand=fields&$top=200");
  const [roster, answers, checkins, intake, content, participationRows, rsvpRows] = await Promise.all(
    ["Couples", "Prompt Answers", "Sunday Check-ins", "Pilot Intake", "Monthly Content", "Participation", "RSVPs"].map(read)
  );
  const result = project(roster.map(item => ({ id: item.id, data: readRow("couples", item) })), answers, checkins, intake, content.map(parse), Date.now());
  async function publish(name, existing, records) {
    const byTitle = new Map(existing.map(item => [item.fields?.Title, item]));
    for (const [Title, record] of records) {
      const fields = { Title, AppData: JSON.stringify(record) };
      const prior = byTitle.get(Title);
      if (prior && prior.fields.AppData !== fields.AppData) {
        if (!prior.eTag) throw new Error("Missing projection eTag");
        await request(`${path(name)}/${prior.id}/fields`, "PATCH", fields, prior.eTag);
      }
      if (!prior) await request(path(name), "POST", { fields });
    }
  }
  await publish("Participation", participationRows, result.participation);
  await publish("RSVPs", rsvpRows, result.rsvps);
  let noteErrors = 0;
  for (const journal of result.journals.values()) {
    if (!journal.folder) continue;
    const target = `${path("Couple Journal")}/${journal.folder}`;
    const item = await request(target + "?$expand=fields");
    const old = parse(item);
    const pilotMonths = [...new Set([...(old.pilotMonths || []), ...journal.months])].sort();
    const releasedNotes = [...(old.releasedNotes || [])];
    for (const note of journal.notes.values()) {
      if (releasedNotes.some(saved => saved.month === note.month && saved.memberId === note.memberId)) continue;
      try {
        releasedNotes.push({ season, month: note.month, memberId: note.memberId, author: note.author, text: decryptNote(note.data, RCI_SEALED_PRIVATE_KEY) });
      } catch { noteErrors++; }
    }
    if (JSON.stringify(old.pilotMonths || []) === JSON.stringify(pilotMonths) &&
        JSON.stringify(old.releasedNotes || []) === JSON.stringify(releasedNotes)) continue;
    const AppData = JSON.stringify({ ...old, version: 1, pilotMonths, releasedNotes });
    if (!item.eTag || AppData.length > 60000) throw new Error("Journal missing eTag or at capacity");
    await request(target + "/fields", "PATCH", { AppData }, item.eTag);
  }
  log(`Processed ${result.participation.size} participation and ${result.rsvps.size} RSVP projections.`);
  if (noteErrors) throw new Error(`${noteErrors} sealed notes could not be released`);
}

if (require.main === module) run().catch(error => {
  // Never print Graph response bodies, note text, secrets or raw fetch errors.
  console.error("Pilot sync failed:", error.message);
  process.exitCode = 1;
});
module.exports = { run, project, decryptNote };
