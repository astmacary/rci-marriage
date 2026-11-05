/* Live data stays in memory and SharePoint. The demo never starts this code. */
(async function () {
  if (window.RCI_APP.DEMO) return;
  const C = window.RCI_CONFIG, D = window.RCIData, esc = D.escape;
  const scopes = ["Sites.ReadWrite.All", "User.Read"];
  const state = { account: null, site: null, lists: {}, rows: {}, errors: {}, loading: false, message: "", view: "today" };
  let auth, controller = new AbortController();
  const root = document.getElementById("app");
  const siteUrl = `https://${C.siteHost}${C.sitePath}`;
  const day = () => new Date().toLocaleDateString("en-CA");
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
        : state.view === "checkin" ? ["checkins", "couples"] : state.view === "journal" ? ["couples"] : ["content", "answers", "couples"];
      await Promise.all(needed.map(load));
      if (state.view === "journal") await loadJournal();
    } catch (error) { state.message = notice(error); }
    state.loading = false; render();
  }
  function profile() {
    const email = (state.account?.username || "").toLowerCase();
    return (state.rows.couples || []).map(item => D.read("couples", item)).find(data => data.kind === "couple" && data.active !== false &&
      ((Array.isArray(data.memberObjectIds) && data.memberObjectIds.includes(oid())) ||
        (Array.isArray(data.memberEmails) && data.memberEmails.some(value => String(value).toLowerCase() === email))));
  }
  async function loadJournal() {
    delete state.errors.reveals;
    state.journal = null;
    const couple = profile();
    if (!couple?.journalItemId) return;
    try { state.journal = await graph(`${base("reveals")}/items/${encodeURIComponent(couple.journalItemId)}?$expand=fields`); }
    catch (error) { state.errors.reveals = notice(error); }
  }
  async function save(key, logicalKey, data) {
    D.validate(data);
    const prior = own(key).find(item => item.fields.Title === logicalKey);
    const fields = { Title: logicalKey, ...D.columns(data, profile()?.coupleKey) };
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
      : `${prompt.data.type ? `<p class="label" style="color:var(--fig);margin:0">${esc(prompt.data.type)}${prompt.data.week ? " · " + esc(prompt.data.week) : ""}</p>` : ""}<p class="big">${esc(prompt.data.text)}</p><form data-form="answer"><input type="hidden" name="promptId" value="${esc(prompt.item.id)}"><label for="live-answer">Your answer</label><textarea id="live-answer" name="answer" maxlength="12000" required rows="5"></textarea><p class="sub">Saved to your private answers. Sharing with your spouse requires the ministry’s Reveal Together flow.</p><button class="btn" type="submit">Save my answer</button><p role="status" class="form-status"></p></form>`;
    return card("Today’s prompt", body) + card("My saved answers", responseList("answers")) + card("Grow", state.errors.content ? empty(state.errors.content) :
      content.filter(({data}) => ["video", "devotional", "challenge"].includes(data.kind)).map(({item,data}) =>
        `<article><h3>${esc(item.fields.Title)}</h3><p style="white-space:pre-wrap">${esc(data.text)}</p>${D.safeLink(data.url) ? `<a href="${esc(D.safeLink(data.url))}" target="_blank" rel="noopener noreferrer">Open resource</a>` : ""}</article>`).join("") || empty("Ministry resources will appear here when published."));
  }
  function checkin() {
    return card("Sunday check-in", `<form data-form="checkin" style="display:grid;gap:12px"><p class="sub">These responses save privately. Spouse sharing and mentor notifications need the ministry flows.</p>
      ${["Connected", "Heard", "Spiritually together", "Having fun"].map((name,i) => `<label>${name}<select name="rating${i}" required style="display:block;font:inherit;padding:10px;width:100%"><option value="">Choose 1–5</option>${[1,2,3,4,5].map(n=>`<option>${n}</option>`).join("")}</select></label>`).join("")}
      <label for="live-need">One thing I need this week</label><textarea id="live-need" name="need" maxlength="6000" rows="3"></textarea>
      <label for="live-saw">Where I saw growth in you</label><textarea id="live-saw" name="saw" maxlength="6000" rows="3"></textarea>
      <label style="display:flex;gap:10px;align-items:center;font-weight:600"><input type="checkbox" name="mentor" value="yes" style="width:22px;height:22px"> We’d like our mentor couple to reach out</label><p class="sub" style="margin-top:-6px">Optional. Only your mentor couple gets a heads-up, with no details.</p>
      <button class="btn" type="submit">Save my check-in</button><p role="status" class="form-status"></p></form>`) + card("My saved check-ins", responseList("checkins"));
  }
  function journal() {
    if (state.errors.couples || state.errors.reveals) return card("Our journal", empty(state.errors.reveals || state.errors.couples));
    if (!state.journal) return card("Our journal", empty("Your couple profile has not been linked to its private journal folder yet. Ask a ministry administrator to finish setup."));
    const savedEntries = D.parse(state.journal).entries;
    const entries = Array.isArray(savedEntries) ? savedEntries : [];
    return card("Our shared journal", `<p class="sub">Notes added here are visible immediately to both spouses. Site owners and tenant administrators retain administrative access.</p>
      ${entries.map(entry=>`<article style="border-top:1px solid var(--line)"><p class="sub">${esc(entry.date)}</p><p style="white-space:pre-wrap">${esc(entry.text)}</p></article>`).join("") || empty("No shared notes yet.")}
      <form data-form="journal"><label for="journal-note">Add a note together</label><textarea id="journal-note" name="text" maxlength="6000" rows="4" required></textarea><button class="btn" type="submit">Save to our journal</button><p role="status" class="form-status"></p></form>`);
  }
  function leaders() {
    return card("Leader dashboard", `<p class="sub">This view loads the four ministry lists only. It does not read couples’ answers, check-ins or journal entries.</p>`) +
      ["couples", "participation", "rsvps", "content"].map(key => card(C.lists[key], state.errors[key] ? empty(state.errors[key]) :
        `<p class="big">${(state.rows[key] || []).length} records</p><a href="${esc(link(key))}" target="_blank" rel="noopener noreferrer">Open ${esc(C.lists[key])} in SharePoint</a>`)).join("");
  }
  function render() {
    const account = state.account;
    root.className = "app";
    root.innerHTML = `<header class="head"><div class="eyebrow">Fruit of Our Union · Live</div><h1>RCI Marriage</h1>${account ? `<p class="sub">${esc(account.name || "Signed in")}</p>` : ""}</header><main>
      ${!account ? card("Welcome", `<p class="sub">Sign in with your church Microsoft account to open your ministry data. The sign-in screen uses the shared RCI Productions registration.</p>${C.clientId ? button("signin", "Sign in with Microsoft") : empty("Live sign-in is being configured.")}${state.message ? `<p role="status">${esc(state.message)}</p>` : ""}<a href="?demo=1">Explore the sample demo</a>`) :
        `<div style="display:flex;gap:8px;flex-wrap:wrap">${[["today","Today"],["checkin","Check-in"],["journal","Our journal"],["leaders","Leaders"]].map(([key,label]) => `<button class="btn ghost" type="button" data-view="${key}" ${state.loading ? "disabled" : ""} ${key===state.view?'aria-current="page"':""}>${label}</button>`).join("")}</div><div style="display:flex;gap:8px">${button("refresh","Refresh")}${button("signout","Sign out")}</div>
        ${state.message ? `<p role="status">${esc(state.message)}</p>` : ""}${state.loading ? empty("Loading from SharePoint…") : !state.site ? empty("The connection is not ready. Refresh to retry.") : ({today,checkin,journal,leaders}[state.view])()}`}
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
        state.account = null; state.rows = {}; state.journal = null; state.site = null; render();
        await auth.logoutRedirect({ account, postLogoutRedirectUri: C.redirectUri });
      }
    } catch (error) { state.message = notice(error); render(); }
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
      if (form.dataset.form === "journal") {
        const item = state.journal, data = D.parse(item);
        if (!item.eTag) throw new Error("Refresh the journal before saving");
        const entry = { version: 1, kind: "journal-entry", id: crypto.randomUUID(), date: new Date().toISOString(), text: values.get("text") };
        D.validate(entry); const entries = [...(Array.isArray(data.entries) ? data.entries : []), entry];
        if (JSON.stringify(entries).length > 55000) throw new Error("Journal capacity reached");
        await graph(`${base("reveals")}/items/${item.id}/fields`, { method: "PATCH", headers: { "If-Match": item.eTag }, body: JSON.stringify({ AppData: JSON.stringify({ ...data, version: 1, entries }) }) });
        await loadJournal();
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
