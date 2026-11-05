/* Live data stays in memory and SharePoint. The demo never starts this code. */
(async function () {
  if (window.RCI_APP.DEMO) return;
  const C = window.RCI_CONFIG, D = window.RCIData, esc = D.escape;
  const scopes = ["Sites.ReadWrite.All", "User.Read"];
  const state = { shared: [], account: null, site: null, lists: {}, rows: {}, errors: {}, loading: false, message: "", view: "today", week: D.weekStart(D.ministryDay()) };
  let auth, controller = new AbortController();
  const root = document.getElementById("app");
  const siteUrl = `https://${C.siteHost}${C.sitePath}`;
  const day = D.ministryDay;
  const oid = () => state.account?.idTokenClaims?.oid || state.account?.localAccountId;
  const own = key => D.own(state.rows[key] || [], oid());
  const link = key => `${siteUrl}/Lists/${encodeURIComponent(C.lists[key])}/AllItems.aspx`;
  const button = (action, text) => `<button class="btn" type="button" data-live="${action}">${text}</button>`;
  const empty = text => `<p class="sub">${esc(text)}</p>`;
  const card = (title, body) => `<section class="card"><h2 style="margin:0;font-size:21px">${esc(title)}</h2>${body}</section>`;
  function notice(error) {
    if (error.status === 403) return "SharePoint has not granted your account access to this area. Ask a ministry administrator to finish your couple setup.";
    if (error.status === 412) return "This entry changed in another session. Refresh before editing it again.";
    if (error.status === 429) return "Microsoft is busy. Please wait a moment, then refresh.";
    if (error.status === 401 || error.name === "InteractionRequiredAuthError") return "Please sign in again to reconnect to Microsoft.";
    return "The connection could not finish. Your response has not been confirmed saved. Check your saved entries before trying again.";
  }
  async function graph(path, options = {}) {
    const url = new URL(path, "https://graph.microsoft.com/v1.0/");
    if (url.origin !== "https://graph.microsoft.com" || !url.pathname.startsWith("/v1.0/sites/")) throw new Error("Unexpected data destination");
    const token = await auth.acquireTokenSilent({ scopes, account: state.account });
    const response = await fetch(url, { ...options, cache: "no-store", signal: controller.signal,
      headers: { Authorization: `Bearer ${token.accessToken}`, "Content-Type": "application/json", ...options.headers } });
    if (!response.ok) { const error = new Error("Microsoft request failed"); error.status = response.status; throw error; }
    return response.status === 204 ? null : response.json();
  }
  async function pages(path) {
    const rows = []; let next = path;
    while (next) { const result = await graph(next); rows.push(...result.value); next = result["@odata.nextLink"]; }
    return rows;
  }
  const base = key => `sites/${state.site.id}/lists/${state.lists[key]}`;
  async function rows(key) {
    if (!state.lists[key]) throw new Error("List is not configured");
    return pages(`${base(key)}/items?$expand=fields&$top=200`);
  }
  async function load(key) {
    delete state.errors[key];
    try { state.rows[key] = await rows(key); }
    catch (error) { state.rows[key] = []; state.errors[key] = notice(error); }
  }
  async function loadView() {
    state.loading = true; render();
    try {
      if (!state.site) {
        state.site = await graph(`sites/${C.siteHost}:${C.sitePath}`);
        const lists = await pages(`sites/${state.site.id}/lists?$select=id,displayName`);
        for (const [key, name] of Object.entries(C.lists)) state.lists[key] = lists.find(list => list.displayName === name)?.id;
      }
      const needed = state.view === "leaders" ? ["couples", "participation", "rsvps", "content"]
        : state.view === "checkin" ? ["checkins", "couples"]
        : ["journal", "orchard"].includes(state.view) ? ["couples"]
        : ["gather", "sealed"].includes(state.view) ? ["content", "pilotIntake"] : ["content", "answers", "couples"];
      await Promise.all(needed.map(load));
      if (["today", "checkin", "journal", "orchard"].includes(state.view)) await loadJournal();
    } catch (error) { state.message = notice(error); }
    state.loading = false; render();
  }
  function profile() {
    const email = (state.account?.username || "").toLowerCase();
    return (state.rows.couples || []).map(item => D.read("couples", item)).find(data => data.kind === "couple" && data.active !== false &&
      ((Array.isArray(data.members) && data.members.some(member => member.objectId && String(member.objectId).toLowerCase() === String(oid()).toLowerCase())) ||
        (Array.isArray(data.memberObjectIds) && data.memberObjectIds.includes(oid())) ||
        (Array.isArray(data.memberEmails) && data.memberEmails.some(value => String(value).toLowerCase() === email))));
  }
  async function loadJournal() {
    delete state.errors.reveals;
    state.journal = null; state.shared = [];
    const couple = profile();
    if (!couple?.journalItemId) return;
    try {
      const [folder, items] = await Promise.all([
        graph(`${base("reveals")}/items/${encodeURIComponent(couple.journalItemId)}?$expand=fields`), rows("reveals")]);
      state.journal = folder; state.shared = D.shared(items, couple);
    } catch (error) { state.errors.reveals = notice(error); }
  }
  // Both answers side by side, the signed-in spouse first.
  function pair(entry) {
    const mine = D.side(profile(), state.account?.username);
    const [me, them] = mine === "spouse2" ? [entry.spouse2, entry.spouse1] : [entry.spouse1, entry.spouse2];
    const col = (label, text) => `<div style="flex:1 1 220px;background:var(--paper,#faf7f2);border-radius:12px;padding:12px"><p class="label" style="margin:0 0 6px">${label}</p><p style="white-space:pre-wrap;margin:0">${esc(text) || '<span class="sub">Nothing shared.</span>'}</p></div>`;
    return `<article style="border-top:1px solid var(--line);padding-top:12px"><p class="sub" style="margin:0 0 4px">${esc(entry.date)}</p>${entry.title && entry.kind === "reveal" ? `<p style="margin:0 0 8px"><b>${esc(entry.title)}</b></p>` : ""}<div style="display:flex;gap:10px;flex-wrap:wrap">${col(mine ? "You" : "Spouse 1", me)}${col(mine ? "Your spouse" : "Spouse 2", them)}</div></article>`;
  }
  function sharedList(kind, limit, none) {
    if (state.errors.reveals) return empty(state.errors.reveals);
    if (!state.journal) return empty("Your couple profile has not been linked to its private journal folder yet. Ask a ministry administrator to finish setup.");
    const list = state.shared.filter(entry => entry.kind === kind).slice(0, limit);
    return list.map(pair).join("") || empty(none);
  }
  async function save(key, logicalKey, data) {
    D.validate(data);
    const prior = own(key).find(item => item.fields.Title === logicalKey);
    const fields = { Title: logicalKey, ...(key === "pilotIntake" ? { AppData: JSON.stringify(data) } : D.columns(data, profile()?.coupleKey)) };
    if (prior) await graph(`${base(key)}/items/${prior.id}/fields`, { method: "PATCH", headers: { "If-Match": prior.eTag || "" }, body: JSON.stringify(fields) });
    else await graph(`${base(key)}/items`, { method: "POST", body: JSON.stringify({ fields }) });
    await load(key);
  }
  function responseList(key) {
    if (state.errors[key]) return empty(state.errors[key]);
    return own(key).slice().sort((a,b) => b.createdDateTime.localeCompare(a.createdDateTime)).map(item => {
      const data = D.read(key, item);
      return `<article style="border-top:1px solid var(--line);padding-top:12px"><b>${esc(data.date || item.fields.Title)}</b><p style="white-space:pre-wrap">${esc(data.answer || data.need || "Check-in saved")}</p></article>`;
    }).join("") || empty("No saved responses yet.");
  }
  function today() {
    const content = (state.rows.content || []).map(item => ({ item, data: D.read("content", item) }));
    const prompt = content.find(({data}) => data.kind === "prompt" && data.date === day());
    const body = state.errors.content ? empty(state.errors.content) : !prompt ? empty("No daily prompt has been published for today.")
      : prompt.data.type === "Check-in" ? `<p class="big">It’s Sunday: no prompt today.</p><p class="sub">${esc(prompt.data.text)}. Do your check-in instead.</p><button class="btn" type="button" data-view="checkin">Open the Sunday check-in</button>`
      : `${prompt.data.type ? `<p class="label" style="color:var(--fig);margin:0">${esc(prompt.data.type)}${prompt.data.week ? " · " + esc(prompt.data.week) : ""}</p>` : ""}<p class="big">${esc(prompt.data.text)}</p><form data-form="answer"><input type="hidden" name="promptId" value="${esc(prompt.item.id)}"><label for="live-answer">Your answer</label><textarea id="live-answer" name="answer" maxlength="12000" required rows="5"></textarea><p class="sub">Only you can see your answer until your spouse answers too. Then you’ll both see the two answers side by side.</p><button class="btn" type="submit">Save my answer</button><p role="status" class="form-status"></p></form>`;
    const revealed = state.shared.find(entry => entry.kind === "reveal" && entry.date === day());
    const answered = own("answers").some(item => D.read("answers", item).date === day());
    const reveal = revealed ? pair(revealed) : answered ? empty("Your answer is saved. When your spouse answers today’s prompt, both answers will appear here. Refresh to check.") : empty("Answer today’s prompt and it will be revealed here once you’ve both answered.");
    return card("Today’s prompt", body) + card("Revealed together", reveal) + card("Earlier reveals", sharedList("reveal", 7, "Reveals will collect here as you both answer prompts.")) + card("My saved answers", responseList("answers")) + card("Grow", state.errors.content ? empty(state.errors.content) :
      content.filter(({data}) => ["video", "devotional", "challenge"].includes(data.kind)).map(({item,data}) =>
        `<article><h3>${esc(item.fields.Title)}</h3><p style="white-space:pre-wrap">${esc(data.text)}</p>${D.safeLink(data.url) ? `<a href="${esc(D.safeLink(data.url))}" target="_blank" rel="noopener noreferrer">Open resource</a>` : ""}</article>`).join("") || empty("Ministry resources will appear here when published."));
  }
  function checkin() {
    return card("Sunday check-in", `<form data-form="checkin" style="display:grid;gap:12px"><p class="sub">Your check-in is shared with your spouse when you save it. If you ask for your mentors, they get a short note to reach out. They don’t see your answers.</p>
      ${["Connected", "Heard", "Spiritually together", "Having fun"].map((name,i) => `<label>${name}<select name="rating${i}" required style="display:block;font:inherit;padding:10px;width:100%"><option value="">Choose 1–5</option>${[1,2,3,4,5].map(n=>`<option>${n}</option>`).join("")}</select></label>`).join("")}
      <label for="live-need">One thing I need this week</label><textarea id="live-need" name="need" maxlength="6000" rows="3"></textarea>
      <label for="live-saw">Where I saw growth in you</label><textarea id="live-saw" name="saw" maxlength="6000" rows="3"></textarea>
      <label style="display:flex;gap:10px;align-items:center;font-weight:600"><input type="checkbox" name="mentor" value="yes" style="width:22px;height:22px"> We’d like our mentor couple to reach out</label><p class="sub" style="margin-top:-6px">Optional. Only your mentor couple gets a heads-up, with no details.</p>
      <button class="btn" type="submit">Save my check-in</button><p role="status" class="form-status"></p></form>`) + card("Our check-ins", sharedList("check-in", 8, "Check-ins you and your spouse save will appear here.")) + card("My saved check-ins", responseList("checkins"));
  }
  function journal() {
    if (state.errors.couples || state.errors.reveals) return card("Our journal", empty(state.errors.reveals || state.errors.couples));
    if (!state.journal) return card("Our journal", empty("Your couple profile has not been linked to its private journal folder yet. Ask a ministry administrator to finish setup."));
    const savedEntries = D.parse(state.journal).entries;
    const entries = Array.isArray(savedEntries) ? savedEntries : [];
    return card("Our shared journal", `<p class="sub">Notes added here are visible immediately to both spouses. Site owners and tenant administrators retain administrative access.</p>
      ${entries.map(entry=>`<article style="border-top:1px solid var(--line)"><p class="sub">${esc(entry.date)}</p><p style="white-space:pre-wrap">${esc(entry.text)}</p></article>`).join("") || empty("No shared notes yet.")}
      <form data-form="journal"><label for="journal-note">Add a note together</label><textarea id="journal-note" name="text" maxlength="6000" rows="4" required></textarea><button class="btn" type="submit">Save to our journal</button><p role="status" class="form-status"></p></form>`) + card("All our reveals", sharedList("reveal", 400, "No reveals yet."));
  }
  function pilotJournal() {
    const issue = state.errors.couples || state.errors.reveals;
    if (!state.journal) return card("Our Orchard", empty(issue || "Ask a ministry administrator to link your couple profile."));
    const data = D.parse(state.journal);
    const months = Array.isArray(data.pilotMonths) ? data.pilotMonths : [];
    const orchard = Array.from({ length: 12 }, (_, index) => {
      const month = new Intl.DateTimeFormat("en", { month: "long", timeZone: "UTC" }).format(new Date(Date.UTC(2027, index)));
      const key = "2027-" + String(index + 1).padStart(2, "0");
      return '<li>' + month + ' · ' + (months.includes(key) ? "Gathered ✓" : "Awaiting participation") + '</li>';
    }).join("");
    const goals = D.goals.map(goal => {
      const saved = data.unionGoals?.[goal.key] || "";
      if (day() < goal.date) return card(goal.label, empty("Opens " + goal.date));
      return card(goal.label, `<form data-form="goal"><input type="hidden" name="goal" value="${goal.key}">
        <label for="goal-${goal.key}">Our ${goal.label.toLowerCase()} goal</label>
        <textarea id="goal-${goal.key}" name="text" required maxlength="2000" rows="4">${esc(saved)}</textarea>
        <button class="btn" type="submit">Save goal</button><p class="form-status" role="status"></p></form>`);
    }).join("");
    const released = Array.isArray(data.releasedNotes) ? data.releasedNotes : [];
    const notes = released.map(note => `<article><h3>${esc(note.author)}</h3><p style="white-space:pre-wrap">${esc(note.text)}</p></article>`).join("");
    return card("Our Orchard", `<p class="sub">A month fills in after either spouse answers a prompt or checks in.</p><ul>${orchard}</ul>`)
      + card("Union Goals", empty("Your three goals for 2027 are shared with your spouse.")) + goals
      + card("Released notes", notes || empty("Your notes will appear here after the ministry releases them."));
  }
  function sealed() {
    const config = (state.rows.content || []).map(D.parse).find(row => row.kind === "sealed-release" && row.season === "2027");
    if (state.errors.content || state.errors.pilotIntake) return card("Sealed note", empty(state.errors.content || state.errors.pilotIntake));
    const saved = own("pilotIntake").some(item => item.fields?.Title === "sealed:2027");
    if (saved) return card("Sealed note", empty("Your note is sealed. Open Our Orchard to see it after release."));
    if (day() < "2027-01-29") return card("Sealed note", empty("Opens January 29, 2027."));
    if (!config?.publicKey || !Number.isFinite(D.instant(config.releaseAt))) return card("Sealed note", empty("The ministry is preparing the release."));
    if (Date.now() >= D.instant(config.releaseAt)) return card("Sealed note", empty("The submission window has closed."));
    const release = new Intl.DateTimeFormat("en-US", { dateStyle: "long", timeStyle: "short", timeZone: "America/New_York" }).format(new Date(config.releaseAt));
    return card("Sealed note", `<p class="sub">Opens ${esc(release)} Eastern. Your note is encrypted when you seal it and can’t be reopened in the app before release.</p>
      <form data-form="sealed"><label for="sealed-note">A note to your spouse</label><textarea id="sealed-note" name="text" rows="5" required maxlength="2000"></textarea>
      <button class="btn" type="submit">Seal my note</button><p class="form-status" role="status"></p></form>`);
  }
  function gather() {
    if (state.errors.content || state.errors.pilotIntake) return card("Gather", empty(state.errors.content || state.errors.pilotIntake));
    const content = (state.rows.content || []).map(D.parse);
    return D.events.map(event => {
      const title = event === "retreat" ? "Marriage Retreat" : "Harvest Celebration";
      const details = content.find(row => row.kind === "event" && row.event === event && row.season === "2027");
      if (!details) return card(title, empty("Details and RSVPs will open when the ministry publishes this event."));
      const saved = own("pilotIntake").find(item => item.fields?.Title === "rsvp:2027:" + event);
      const selection = saved ? D.parse(saved).response : "";
      return card(title, `<p class="sub" style="white-space:pre-wrap">${esc(details.text)}</p>
        <form data-form="rsvp"><input type="hidden" name="event" value="${event}">
        <label for="rsvp-${event}">Your RSVP</label><select id="rsvp-${event}" name="response" required>
          <option value="">Choose a response</option>${[["yes","Attending"],["no","Unable to attend"]].map(([value,label]) =>
            `<option value="${value}" ${selection === value ? "selected" : ""}>${label}</option>`).join("")}
        </select><button class="btn" type="submit">Save RSVP</button><p class="form-status" role="status"></p></form>`);
    }).join("");
  }
  async function updateJournal(changes) {
    const item = state.journal;
    if (!item?.eTag) throw new Error("Refresh the journal before saving");
    const body = JSON.stringify({ ...D.parse(item), ...changes, version: 1 });
    if (body.length > 60000) throw new Error("Journal capacity reached");
    await graph(`${base("reveals")}/items/${item.id}/fields`, {
      method: "PATCH", headers: { "If-Match": item.eTag }, body: JSON.stringify({ AppData: body })
    });
    await loadJournal();
  }
  async function sealNote(text) {
    const config = (state.rows.content || []).map(D.parse).find(row => row.kind === "sealed-release" && row.season === "2027");
    if (!text || text.length > 2000 || !config?.publicKey || day() < "2027-01-29" ||
      !(Date.now() < D.instant(config.releaseAt)) ||
      own("pilotIntake").some(item => item.fields?.Title === "sealed:2027")) throw new Error("Sealing is unavailable");
    const publicKey = await crypto.subtle.importKey("jwk", config.publicKey, { name: "RSA-OAEP", hash: "SHA-256" }, false, ["wrapKey"]);
    const key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt"]);
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const bytes = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(text));
    const wrapped = await crypto.subtle.wrapKey("raw", key, publicKey, { name: "RSA-OAEP" });
    const encode = value => btoa(String.fromCharCode(...new Uint8Array(value)));
    await save("pilotIntake", "sealed:2027", { version: 1, kind: "sealed-note", season: "2027",
      iv: encode(iv), ciphertext: encode(bytes), wrappedKey: encode(wrapped) });
  }
  function leaders() {
    if (state.errors.couples || state.errors.participation) {
      return card("Weekly participation", empty(state.errors.couples || state.errors.participation));
    }
    const activity = (state.rows.participation || []).map(D.parse).filter(row => row.week === state.week);
    const roster = (state.rows.couples || []).map(item => ({ ...D.read("couples", item), id: item.id, title: item.fields?.Title }))
      .filter(couple => couple.kind === "couple" && couple.active !== false);
    const branches = [...new Set(roster.map(couple => couple.branch || "Unassigned"))].sort();
    const summary = branches.map(branch => card(branch, '<div style="overflow-x:auto"><table><thead><tr><th>Participant</th><th>Answered</th><th>Checked in</th></tr></thead><tbody>' +
      roster.filter(couple => (couple.branch || "Unassigned") === branch).flatMap(couple => {
        if (!Array.isArray(couple.members) || !couple.members.length) {
          return [`<tr><td>${esc(couple.title)}</td><td colspan="2">Roster setup needed</td></tr>`];
        }
        const members = couple.members;
        return members.map(member => {
          const status = activity.filter(row => row.coupleId === couple.id && (!member.objectId || row.memberId === member.objectId.toLowerCase()));
          return `<tr><td>${esc(member.name)}</td><td>${status.some(row => row.answered) ? "Yes" : "No record"}</td>
            <td>${status.some(row => row.checkedIn) ? "Yes" : "No record"}</td></tr>`;
        });
      }).join("") + "</tbody></table></div>")).join("");
    return card("Weekly participation", `<label for="leader-week">Week containing</label>
      <input id="leader-week" type="date" value="${state.week}"><p class="sub">Week of ${state.week} · Monday through Sunday, Eastern time</p>`)
      + (summary || empty("No active couples are listed."))
      + ["rsvps", "content"].map(key => card(C.lists[key], state.errors[key] ? empty(state.errors[key]) :
        `<a href="${esc(link(key))}" target="_blank" rel="noopener noreferrer">Open ${esc(C.lists[key])} in SharePoint</a>`)).join("");
  }
  function render() {
    const account = state.account;
    root.className = "app";
    root.innerHTML = `<header class="head"><div class="eyebrow">Fruit of Our Union · Live</div><h1>RCI Marriage</h1>${account ? `<p class="sub">${esc(account.name || "Signed in")}</p>` : ""}</header><main>
      ${!account ? card("Welcome", `<p class="sub">Sign in with your church Microsoft account to open your ministry data. The sign-in screen uses the shared RCI Productions registration.</p>${C.clientId ? button("signin", "Sign in with Microsoft") : empty("Live sign-in is being configured.")}${state.message ? `<p role="status">${esc(state.message)}</p>` : ""}<a href="?demo=1">Explore the sample demo</a>`) :
        `<div style="display:flex;gap:8px;flex-wrap:wrap">${[["today","Today"],["checkin","Check-in"],["journal","Our journal"],["orchard","Our Orchard"],["sealed","Sealed note"],["gather","Gather"],["leaders","Leaders"]].map(([key,label]) => `<button class="btn ghost" type="button" data-view="${key}" ${state.loading ? "disabled" : ""} ${key===state.view?'aria-current="page"':""}>${label}</button>`).join("")}</div><div style="display:flex;gap:8px">${button("refresh","Refresh")}${button("signout","Sign out")}</div>
        ${state.message ? `<p role="status">${esc(state.message)}</p>` : ""}${state.loading ? empty("Loading from SharePoint…") : !state.site ? empty("The connection is not ready. Refresh to retry.") : ({today,checkin,journal,orchard:pilotJournal,sealed,gather,leaders}[state.view])()}`}
      </main>`;
  }
  window.renderLive = render;
  document.addEventListener("click", async event => {
    const view = event.target.closest("[data-view]");
    if (view && !state.loading) { state.view = view.dataset.view; state.message = ""; await loadView(); return; }
    const target = event.target.closest("[data-live]"); if (!target) return;
    try {
      if (target.dataset.live === "signin") await auth.loginRedirect({ scopes, redirectUri: C.redirectUri });
      if (target.dataset.live === "refresh" && !state.loading) await loadView();
      if (target.dataset.live === "signout") {
        controller.abort(); controller = new AbortController(); const account = state.account;
        state.account = null; state.rows = {}; state.journal = null; state.shared = []; state.site = null; render();
        await auth.logoutRedirect({ account, postLogoutRedirectUri: C.redirectUri });
      }
    } catch (error) { state.message = notice(error); render(); }
  });
  document.addEventListener("change", event => {
    if (event.target.id !== "leader-week" || !/^\d{4}-\d{2}-\d{2}$/.test(event.target.value)) return;
    state.week = D.weekStart(event.target.value);
    render();
  });
  document.addEventListener("submit", async event => {
    const form = event.target.closest("[data-form]"); if (!form) return;
    event.preventDefault(); const submit = form.querySelector('[type="submit"]');
    if (submit.disabled) return; submit.disabled = true;
    const status = form.querySelector(".form-status"), values = new FormData(form);
    status.textContent = "Saving to SharePoint…";
    try {
      if (form.dataset.form === "answer") await save("answers", `prompt:${values.get("promptId")}:${day()}`, { version: 1, kind: "prompt-answer", promptId: values.get("promptId"), date: day(), answer: values.get("answer") });
      if (form.dataset.form === "checkin") await save("checkins", `check-in:${day()}`, { version: 1, kind: "check-in", date: day(), ratings: [0,1,2,3].map(i=>Number(values.get(`rating${i}`))), need: values.get("need"), saw: values.get("saw"), mentor: values.get("mentor") === "yes" });
      if (form.dataset.form === "goal") {
        const goal = D.goals.find(item => item.key === values.get("goal"));
        const text = String(values.get("text") || "").trim();
        if (!goal || day() < goal.date || !text || text.length > 2000) throw new Error("Invalid goal");
        await updateJournal({ unionGoals: { ...D.parse(state.journal).unionGoals, [goal.key]: text } });
      }
      if (form.dataset.form === "sealed") await sealNote(String(values.get("text") || "").trim());
      if (form.dataset.form === "rsvp") {
        const event = values.get("event"), response = values.get("response");
        const published = (state.rows.content || []).map(D.parse).some(row => row.kind === "event" && row.season === "2027" && row.event === event);
        if (!published || !D.events.includes(event) || !["yes", "no"].includes(response)) throw new Error("Invalid RSVP");
        await save("pilotIntake", `rsvp:2027:${event}`, { version: 1, kind: "rsvp", season: "2027", event, response });
      }
      if (form.dataset.form === "journal") {
        const item = state.journal, data = D.parse(item);
        if (!item.eTag) throw new Error("Refresh the journal before saving");
        const entry = { version: 1, kind: "journal-entry", id: crypto.randomUUID(), date: new Date().toISOString(), text: values.get("text") };
        D.validate(entry); const entries = [...(Array.isArray(data.entries) ? data.entries : []), entry];
        if (JSON.stringify(entries).length > 55000) throw new Error("Journal capacity reached");
        await updateJournal({ entries });
      }
      state.message = "Saved to SharePoint."; render();
    } catch (error) { status.textContent = notice(error); submit.disabled = false; }
  });
  render();
  if (!C.clientId) return;
  try {
    auth = new msal.PublicClientApplication({ auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`, redirectUri: C.redirectUri, navigateToLoginRequestUrl: false }, cache: { cacheLocation: "sessionStorage" } });
    await auth.initialize(); const result = await auth.handleRedirectPromise();
    state.account = result?.account || auth.getAllAccounts().find(account => account.tenantId === C.tenantId) || null;
    if (state.account) { auth.setActiveAccount(state.account); await loadView(); }
  } catch (error) { state.message = notice(error); render(); }
})();
