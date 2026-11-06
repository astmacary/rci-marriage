const { createHash } = require("node:crypto");

const graphOrigin = "https://graph.microsoft.com";
const guid = /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
const list = name => "web/lists/getbytitle('" + encodeURIComponent(name) + "')";
const number = value => {
  if (!Number.isSafeInteger(Number(value)) || Number(value) < 1) throw new Error("Invalid SharePoint ID");
  return Number(value);
};
const parse = value => {
  try { return JSON.parse(value || "{}") || {}; } catch { return {}; }
};
const lower = value => String(value || "").trim().toLowerCase();

class Review extends Error {}

function requestData(item) {
  const data = parse(item.AppData);
  const name = String(data.coupleName || "").trim();
  const spouses = Array.isArray(data.spouses) ? data.spouses.map(spouse => ({
    name: String(spouse?.name || "").trim(), email: lower(spouse?.email)
  })) : [];
  if (data.version !== 1 || data.kind !== "couple-onboarding" || !name || name.length > 120 ||
      /[\\/:*?"<>|#%\u0000-\u001f]/.test(name) || /[. ]$/.test(name) || spouses.length !== 2 ||
      spouses.some(p => !p.name || p.name.length > 120 || p.email.length > 200 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) ||
      spouses[0].email === spouses[1].email) throw new Review("Invalid couple details");
  const hash = createHash("sha256").update(spouses.map(p => p.email).sort().join("\n")).digest("hex");
  if (item.Title !== "onboard:" + hash) throw new Review("Request identity changed");
  return { name, spouses, key: "onboard-" + hash, suffix: hash.slice(0, 16) };
}

async function token(resource, clientId) {
  const { IDENTITY_ENDPOINT, IDENTITY_HEADER } = process.env;
  if (!IDENTITY_ENDPOINT || !IDENTITY_HEADER || !guid.test(clientId)) throw new Error("Onboarding identity missing");
  const url = new URL(IDENTITY_ENDPOINT);
  url.searchParams.set("resource", resource);
  url.searchParams.set("client_id", clientId);
  url.searchParams.set("api-version", "2019-08-01");
  const response = await fetch(url, { redirect: "error", signal: AbortSignal.timeout(10000), headers: { "X-IDENTITY-HEADER": IDENTITY_HEADER } });
  if (!response.ok) throw new Error("Onboarding identity token unavailable");
  return (await response.json()).access_token;
}

function sharePoint(site, accessToken) {
  const root = site + "/_api/";
  async function call(path, body, headers = {}) {
    const url = new URL(path, root);
    if (!url.href.startsWith(root)) throw new Error("Unexpected SharePoint destination");
    const response = await fetch(url, {
      method: body === undefined ? "GET" : "POST", redirect: "error", signal: AbortSignal.timeout(15000),
      headers: { Authorization: "Bearer " + accessToken, Accept: "application/json;odata=minimalmetadata",
        "Content-Type": "application/json;odata=minimalmetadata", ...headers },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (!response.ok) { const error = new Error("SharePoint onboarding request failed (" + response.status + ")"); error.status = response.status; throw error; }
    const text = await response.text();
    return text ? JSON.parse(text) : {};
  }
  async function all(path) {
    const items = [];
    while (path) {
      const page = await call(path);
      if (!Array.isArray(page.value)) throw new Error("Unexpected SharePoint result");
      items.push(...page.value);
      path = page["odata.nextLink"] || page["@odata.nextLink"];
    }
    return items;
  }
  async function merge(path, fields) {
    const current = await call(path);
    const etag = current["odata.etag"] || current["@odata.etag"];
    if (!etag) throw new Error("SharePoint did not return an item version");
    await call(path, fields, { "IF-MATCH": etag, "X-HTTP-Method": "MERGE" });
  }
  return { call, all, merge };
}

async function directoryAccount(person, accessToken) {
  const q = person.email.replace(/'/g, "''");
  const url = new URL("/v1.0/users", graphOrigin);
  url.searchParams.set("$filter", "userPrincipalName eq '" + q + "' or mail eq '" + q + "'");
  url.searchParams.set("$select", "id,userPrincipalName,mail");
  const response = await fetch(url, { redirect: "error", signal: AbortSignal.timeout(10000), headers: { Authorization: "Bearer " + accessToken } });
  if (!response.ok) throw new Error("Directory lookup failed (" + response.status + ")");
  const result = await response.json();
  const matches = result.value?.filter(user => [lower(user.mail), lower(user.userPrincipalName)].includes(person.email));
  if (matches?.length !== 1 || result["@odata.nextLink"] || !guid.test(matches[0].id)) throw new Review("One email did not resolve to one existing Microsoft account");
  return { ...person, objectId: lower(matches[0].id), upn: lower(matches[0].userPrincipalName) };
}

async function run(log = console.log) {
  const { RCI_ONBOARDING_CLIENT_ID, RCI_ONBOARDING_SITE, RCI_ONBOARDING_LEADER_IDS, RCI_ONBOARDING_ENABLED } = process.env;
  if (!["true", "probe"].includes(RCI_ONBOARDING_ENABLED)) return;
  const site = new URL(RCI_ONBOARDING_SITE);
  if (site.protocol !== "https:" || !/^[a-z0-9-]+\.sharepoint\.com$/i.test(site.hostname) ||
      site.port || site.username || site.password || site.search || site.hash || site.pathname !== "/sites/RCIMarriage")
    throw new Error("Invalid onboarding site");
  const leaders = new Set(RCI_ONBOARDING_LEADER_IDS.split(",").map(number));
  if (leaders.size < 1) throw new Error("No approved onboarding leaders");
  const sp = sharePoint(site.href, await token(site.origin, RCI_ONBOARDING_CLIENT_ID));
  const queue = list("Couple Onboarding"), couples = list("Couples"), journal = list("Couple Journal");
  const requests = await sp.all(queue + "/items?$select=Id,Title,AppData,AuthorId&$top=200");
  const pending = requests.filter(row => leaders.has(Number(row.AuthorId)) && ["queued", "processing"].includes(parse(row.AppData).status)).slice(0, 5);
  if (!pending.length && RCI_ONBOARDING_ENABLED !== "probe") return;
  const graphToken = await token(graphOrigin, RCI_ONBOARDING_CLIENT_ID);
  const owner = number((await sp.call("web/associatedownergroup?$select=Id")).Id);
  const roles = Object.fromEntries((await sp.all("web/roledefinitions?$select=Id,Name")).map(row => [row.Name, number(row.Id)]));
  const getRole = name => number(roles[name]);
  const assignments = path => sp.all(path + "/roleassignments?$expand=RoleDefinitionBindings");
  const hasRole = (row, role) => row.RoleDefinitionBindings?.some(item => Number(item.Id) === role);
  const grant = async (path, principal, role) => {
    const existing = await assignments(path);
    if (existing.some(row => Number(row.PrincipalId) === principal && hasRole(row, role))) return;
    await sp.call(path + "/roleassignments/addroleassignment(principalid=" + principal + ",roledefid=" + role + ")", {});
  };
  const privateItem = async (path, grants) => {
    const row = await sp.call(path + "?$select=HasUniqueRoleAssignments");
    if (!row.HasUniqueRoleAssignments) await sp.call(path + "/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)", {});
    const allowed = new Set([owner, ...grants.map(([id]) => id)]);
    const unexpected = actual => actual.some(row => !allowed.has(Number(row.PrincipalId)) &&
      row.RoleDefinitionBindings?.some(role => role.Name !== "Limited Access"));
    if (unexpected(await assignments(path))) throw new Review("An unexpected principal has access to the item");
    await grant(path, owner, getRole("Full Control"));
    for (const [id, role] of grants) await grant(path, id, role);
    const actual = await assignments(path);
    for (const [id, role] of [[owner, getRole("Full Control")], ...grants]) {
      if (!actual.some(row => Number(row.PrincipalId) === id && hasRole(row, role))) throw new Error("Item access could not be verified");
    }
    if (unexpected(actual)) throw new Review("An unexpected principal has access to the item");
  };
  const itemPath = (parent, id) => parent + "/items(" + number(id) + ")";
  const sourceLists = ["Prompt Answers", "Sunday Check-ins", "Pilot Intake"].map(list);
  for (const source of sourceLists) {
    const settings = await sp.call(source + "?$select=HasUniqueRoleAssignments,ReadSecurity,WriteSecurity");
    if (!settings.HasUniqueRoleAssignments || settings.ReadSecurity !== 2 || settings.WriteSecurity !== 2)
      throw new Error("An intake list is missing author-only permissions");
  }
  const journalMeta = await sp.call(journal + "?$select=HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&$expand=RootFolder");
  if (!journalMeta.HasUniqueRoleAssignments || (await assignments(journal)).some(row =>
    Number(row.PrincipalId) !== owner && row.RoleDefinitionBindings?.some(role => role.Name !== "Limited Access")))
    throw new Error("Journal parent is not restricted to owners");
  const rootPath = journalMeta.RootFolder?.ServerRelativeUrl;
  if (!rootPath?.startsWith(site.pathname + "/Lists/")) throw new Error("Unexpected journal root");
  if (RCI_ONBOARDING_ENABLED === "probe") {
    for (const id of leaders) {
      const person = await sp.call("web/getuserbyid(" + id + ")?$select=Email,UserPrincipalName");
      await directoryAccount({ email: lower(person.UserPrincipalName || person.Email) }, graphToken);
    }
    log("Verified read-only SharePoint and directory access. Onboarding is in probe mode.");
    return;
  }
  let completed = 0, blocked = 0;
  for (const row of pending) {
    const path = itemPath(queue, row.Id);
    const status = async (value, reason) => sp.merge(path, { AppData: JSON.stringify({ ...parse(row.AppData), status: value, reason }) });
    try {
      const data = requestData(row);
      const spouses = await Promise.all(data.spouses.map(person => directoryAccount(person, graphToken)));
      if (spouses[0].objectId === spouses[1].objectId || spouses.some(person => !person.upn)) throw new Review("Two different existing Microsoft accounts are required");
      for (const person of spouses) {
        const user = await sp.call("web/ensureuser", { logonName: "i:0#.f|membership|" + person.upn });
        if (![lower(user.Email), lower(user.UserPrincipalName), lower(user.LoginName).split("|").at(-1)].includes(person.upn) &&
            lower(user.Email) !== person.email) throw new Review("SharePoint could not confirm a spouse account");
        person.principalId = number(user.Id);
      }
      if (spouses[0].principalId === spouses[1].principalId) throw new Review("SharePoint resolved both emails to the same account");
      const records = await sp.all(couples + "/items?$select=Id,Title,CoupleKey,AppData,Spouse1Email,Spouse2Email,JournalItemId&$top=200");
      const ours = records.find(record => record.CoupleKey === data.key);
      if (ours && parse(ours.AppData).onboardingKey !== data.key) throw new Review("Roster key is already in use");
      if (records.some(record => record.CoupleKey !== data.key && spouses.some(person =>
        [record.Spouse1Email, record.Spouse2Email, ...(parse(record.AppData).memberEmails || [])].map(lower).includes(person.email) ||
        (parse(record.AppData).members || []).some(member => lower(member.objectId) === person.objectId))))
        throw new Review("A spouse is already on the roster");

      let roster = ours;
      const reserved = parse(roster?.AppData);
      const leaf = data.name + " — " + data.suffix;
      if (ours && (reserved.onboardingFolder !== leaf || JSON.stringify(reserved.onboardingIds) !== JSON.stringify(spouses.map(person => person.objectId))))
        throw new Review("A previously started request has different couple details");
      if (!roster) roster = await sp.call(couples + "/items", {
        Title: "Onboarding " + data.suffix, CoupleKey: data.key,
        AppData: JSON.stringify({ version: 1, kind: "couple", active: false, onboardingKey: data.key,
          onboardingFolder: leaf, onboardingIds: spouses.map(person => person.objectId) })
      });
      const rosterPath = itemPath(couples, roster.Id);
      await privateItem(rosterPath, [...new Set([...leaders, ...spouses.map(person => person.principalId)])].map(id => [id, getRole("Read")]));
      const folderPath = "web/GetFolderByServerRelativePath(decodedurl='" + encodeURIComponent((rootPath + "/" + leaf).replace(/'/g, "''")) + "')/ListItemAllFields";
      let folder;
      try { folder = await sp.call(folderPath); }
      catch (error) {
        if (error.status !== 404) throw error;
        const result = await sp.call(journal + "/AddValidateUpdateItemUsingPath", {
          listItemCreateInfo: { FolderPath: { DecodedUrl: site.origin + rootPath }, UnderlyingObjectType: 1 },
          formValues: [{ FieldName: "FileLeafRef", FieldValue: leaf }, { FieldName: "Title", FieldValue: data.name },
            { FieldName: "AppData", FieldValue: JSON.stringify({ version: 1, onboardingKey: data.key }) }],
          bNewDocumentUpdate: false
        });
        if (result.value?.some(field => field.HasException)) throw new Review("SharePoint rejected the folder name or metadata");
        folder = await sp.call(folderPath);
      }
      if (parse(folder.AppData).onboardingKey !== data.key) throw new Review("The journal folder name is already in use");
      const folderId = number(folder.Id);
      const folderItem = itemPath(journal, folderId);
      await privateItem(folderItem, spouses.map(person => [person.principalId, getRole("Contribute")]));

      for (const person of spouses) {
        for (const source of sourceLists) await grant(source, person.principalId, getRole("Contribute"));
        await grant(list("Monthly Content"), person.principalId, getRole("Read"));
        await grant("web/GetFolderByServerRelativePath(decodedurl='" + encodeURIComponent(site.pathname + "/Shared Documents/Leader Videos") + "')/ListItemAllFields", person.principalId, getRole("Read"));
      }
      if (!reserved.active) await sp.merge(rosterPath, {
        Title: data.name, Spouse1Name: spouses[0].name, Spouse2Name: spouses[1].name,
        Spouse1Email: spouses[0].email, Spouse2Email: spouses[1].email, JournalItemId: String(folderId),
        AppData: JSON.stringify({ ...parse(roster.AppData), version: 1, kind: "couple", active: true, onboardingKey: data.key,
          onboardingFolder: leaf, onboardingIds: spouses.map(person => person.objectId),
          memberEmails: spouses.map(person => person.email), members: spouses.map(({ name, objectId }) => ({ name, objectId })),
          journalItemId: String(folderId) })
      });
      await status("ready");
      completed++;
    } catch (error) {
      if (!(error instanceof Review)) throw error;
      await status("blocked", error.message);
      blocked++;
    }
  }
  log("Couple onboarding completed " + completed + " requests; " + blocked + " need review.");
}

module.exports = { run, requestData };
