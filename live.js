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
  const card = (title, body) => `<section class="card"><h2 class="label" style="margin:0;color:var(--fig)">${esc(title)}</h2>${body}</section>`;
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
        : state.view === "journal" ? ["couples", "answers"]
        : state.view === "orchard" ? ["couples", "content", "pilotIntake"]
        : state.view === "branch" ? ["couples"]
        : state.view === "grow" ? ["content"]
        : ["gather", "sealed"].includes(state.view) ? ["content", "pilotIntake"] : ["content", "answers", "couples", "pilotIntake"];
      await Promise.all(needed.map(load));
      if (["today", "checkin", "journal", "orchard", "branch"].includes(state.view)) await loadJournal();
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
    const col = (label, text) => `<div style="flex:1 1 220px;background:var(--fig-soft);border-radius:12px;padding:12px"><p class="label" style="margin:0 0 6px">${label}</p><p style="white-space:pre-wrap;margin:0">${esc(text) || '<span class="sub">Nothing shared.</span>'}</p></div>`;
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
  const FRUIT = [["January","The Vine","John 15:5"],["February","Love","1 Corinthians 13:4–7"],["March","Joy","Philippians 4:4"],["April","Peace","Colossians 3:15"],
    ["May","Patience","Ephesians 4:2"],["June","Kindness","Ephesians 4:32"],["July","Goodness","Galatians 6:10"],["August","Faithfulness","Proverbs 3:3"],
    ["September","Gentleness","Proverbs 15:1"],["October","Self-control","Proverbs 25:28"],["November","Gratitude","1 Thessalonians 5:18"],["December","Harvest","John 15:8"]];
  const FRUIT_COLOR = ["#2F5D3F","#B23A48","#D99A1E","#3E7CB1","#C4955F","#C77D3A","#5B7F3A","#6B4F9E","#4F8A8B","#8A5A44","#B8860B","#C4955F"];
  const STEPS = ["Plant","Water","Grow","Harvest"];
  const parts = () => { const [y,m,d] = day().split("-").map(Number); return { y, m, d, date: new Date(Date.UTC(y, m - 1, d)) }; };
  const fmt = (opts) => new Intl.DateTimeFormat("en-US", { timeZone: "UTC", ...opts }).format(parts().date);
  const weekIndex = () => Math.min(3, Math.floor((parts().d - 1) / 7));
  const firstName = () => esc(String(state.account?.name || "").split(/[\s,]+/).filter(Boolean)[0] || "");
  const label = (text, extra = "") => `<span class="label" style="color:var(--fig)${extra}">${text}</span>`;
  const panel = (lbl, body, extra = "") => `<section class="card"${extra}>${lbl ? label(lbl) : ""}${body}</section>`;
  const tile = (view, lbl, big, small) => `<button type="button" class="card" data-view="${view}" style="gap:4px;text-align:left;cursor:pointer;font:inherit;color:var(--ink)"><span style="font-size:12px;font-weight:700;color:var(--fig)">${lbl}</span><span style="font-size:16px;font-weight:700">${big}</span><span class="sub" style="font-size:13px">${small}</span></button>`;
  const TAB_OF = { today: "today", checkin: "us", journal: "us", orchard: "us", grow: "grow", branch: "branch", gather: "gather" };
  const TAB_ICON = {
    today: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>',
    us: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="9" cy="13" r="5"/><circle cx="15" cy="13" r="5"/></svg>',
    grow: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21v-9M12 12c0-4 3-7 8-7 0 4-3 7-8 7zM12 14c0-3-2.5-5.5-7-5.5 0 3 2.5 5.5 7 5.5z"/></svg>',
    branch: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="9" cy="8" r="3"/><circle cx="17" cy="9" r="2.5"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6M15 14.5c3 0 6 2 6 5.5"/></svg>',
    gather: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>'
  };
  function tabBar() {
    const cur = TAB_OF[state.view];
    const item = (id, view, text) => `<button type="button" data-view="${view}" ${cur === id ? 'aria-current="page"' : ""} ${state.loading ? "disabled" : ""}>${TAB_ICON[id]}${text}</button>`;
    return `<nav class="tabs" aria-label="Main"><div>${item("today","today","Today")}${item("us","checkin","Us")}${item("grow","grow","Grow")}${item("branch","branch","Branch")}${item("gather","gather","Gather")}</div></nav>`;
  }
  function usSeg(cur) {
    const a = (view, text) => `<button type="button" data-view="${view}" ${cur === view ? 'aria-current="page"' : ""}>${text}</button>`;
    return `<div class="seg seg3">${a("checkin","Check-in")}${a("orchard","Our Orchard")}${a("journal","Journal")}</div>`;
  }
  function header() {
    const v = state.view, f = FRUIT[parts().m - 1];
    const H = (eyebrow, title, top = "") => top ? `<header class="head">${top}<h1 style="margin-top:6px">${title}</h1>${eyebrow ? `<div class="eyebrow" style="font-weight:400">${eyebrow}</div>` : ""}</header>` : `<header class="head">${eyebrow ? `<div class="eyebrow">${eyebrow}</div>` : ""}<h1>${title}</h1></header>`;
    if (!state.account) return H("Fruit of Our Union · 2027", "Welcome");
    if (v === "today") { const n = firstName(); return H(`${fmt({ weekday: "long", month: "long", day: "numeric" })} · Week ${weekIndex() + 1} of 4`, `Good ${new Date().getHours() < 12 ? "morning" : new Date().getHours() < 17 ? "afternoon" : "evening"}${n ? ", " + n : ""}`); }
    if (v === "checkin") return H("Shared with your spouse", "Sunday check-in", usSeg("checkin"));
    if (v === "orchard") return H("Fills in month by month. Revealed in full at the Harvest Celebration.", "Our Orchard", usSeg("orchard"));
    if (v === "journal") return H("Visible to both of you", "Our journal", usSeg("journal"));
    if (v === "grow") return H(`${f[0]} · ${f[2]}`, `${f[1]}, week by week`);
    if (v === "branch") return H("Your Branch", esc(profile()?.branch ? profile().branch + " Branch" : "Branch"));
    if (v === "gather") return H("RCI Marriage Ministry · 2027", "Gather");
    if (v === "leaders") return H("Leader dashboard", "How the ministry is growing");
    return H("", "RCI Marriage");
  }
  function accountStrip() {
    if (!state.account) return "";
    const leader = state.view === "leaders";
    return `<div class="demo live-strip"><span>${esc(state.account.name || "Signed in")}</span><span style="display:flex;gap:6px;flex-wrap:wrap">
      <button type="button" data-view="${leader ? "today" : "leaders"}" ${state.loading ? "disabled" : ""}>${leader ? "Couple view" : "Leader view"}</button>
      <button type="button" data-live="refresh">Refresh</button><button type="button" data-live="signout">Sign out</button></span></div>`;
  }
  function today() {
    const content = (state.rows.content || []).map(item => ({ item, data: D.read("content", item) }));
    const prompt = content.find(({data}) => data.kind === "prompt" && data.date === day());
    const f = FRUIT[parts().m - 1], w = weekIndex();
    const hero = `<section class="month" aria-label="This month"><div class="row"><span class="label">${esc(f[0])} fruit</span><span style="font-size:12px;color:#C9D2DD">Fruit of Our Union</span></div>
      <div class="name">${esc(f[1])}</div><div style="font-size:15px;line-height:1.4;color:#C9D2DD">${esc(f[2])}</div>
      <div class="rhythm">${STEPS.map((step, i) => `<div class="${i < w ? "done" : i === w ? "now" : ""}">${step}${i < w ? " ✓" : ""}</div>`).join("")}</div></section>`;
    const revealed = state.shared.find(entry => entry.kind === "reveal" && entry.date === day());
    const answered = own("answers").some(item => D.read("answers", item).date === day());
    let promptCard;
    if (state.errors.content) promptCard = panel("Reveal Together", empty(state.errors.content));
    else if (!prompt) promptCard = panel("Reveal Together", `<div class="big">No prompt has been published for today.</div><p class="sub">Check back tomorrow, or open your journal.</p>`);
    else if (prompt.data.type === "Check-in") promptCard = panel("Sunday · Check-in", `<div class="big">It’s Sunday: no prompt today.</div><p class="sub">${esc(prompt.data.text)}. Do your check-in instead.</p><button class="btn" type="button" data-view="checkin">Open the Sunday check-in</button>`);
    else promptCard = `<section class="card"><div class="row">${label("Reveal Together" + (prompt.data.type ? " · " + esc(prompt.data.type) : ""))}<span class="pill" style="background:var(--soft);color:var(--ink)">${revealed ? "Both answered" : answered ? "You answered" : "Waiting on you"}</span></div>
      <div class="big">${esc(prompt.data.text)}</div>
      ${revealed ? pair(revealed) : answered ? `<div class="lock">Your answer is locked in. Your spouse’s answer unlocks once they reply. Refresh to check.</div>`
        : `<form data-form="answer" style="display:flex;flex-direction:column;gap:10px"><input type="hidden" name="promptId" value="${esc(prompt.item.id)}"><label for="live-answer" class="sub" style="font-size:13px;font-weight:400">Only you can see this until your spouse answers too.</label><textarea id="live-answer" name="answer" maxlength="12000" required rows="4" placeholder="Type your answer…"></textarea><button class="btn" type="submit">Answer privately</button><p role="status" class="form-status sub" style="margin:0"></p></form>`}
      ${revealed ? `<p class="sub">Take 5 minutes to talk about what you each wrote. No fixing, just listening.</p>` : ""}</section>`;
    const challenge = content.find(({data}) => data.kind === "challenge");
    const challengeCard = challenge ? `<button type="button" class="card" data-view="grow" style="text-align:left;cursor:pointer;font:inherit;color:var(--ink)">${label("This week’s challenge · " + STEPS[w])}<span style="font-size:16px;font-weight:700">${esc(challenge.item.fields.Title)}</span><span class="sub">${esc(challenge.data.text)}</span></button>` : "";
    if (state.errors.answers) promptCard = panel("Reveal Together", empty(state.errors.answers));
    return hero + promptCard + challengeCard + `<div class="grid2">${tile("checkin", "Sunday check-in", "This week", "2 minutes")}${tile("gather", "Gatherings", "Retreat &amp; Harvest", "RSVP")}</div>`
      + `<button type="button" class="card" data-view="journal" style="flex-direction:row;justify-content:space-between;align-items:center;cursor:pointer;font:inherit;color:var(--ink);text-align:left"><span><span class="label" style="color:var(--muted)">Earlier reveals</span><br><span style="font-weight:600">Our journal</span></span><span aria-hidden="true" style="color:var(--fig)">›</span></button>`;
  }
  function grow() {
    if (state.errors.content) return panel("Grow", empty(state.errors.content));
    const content = (state.rows.content || []).map(item => ({ item, data: D.read("content", item) }));
    const order = ["video", "devotional", "challenge"];
    const items = content.filter(({data}) => order.includes(data.kind)).sort((a, b) => order.indexOf(a.data.kind) - order.indexOf(b.data.kind));
    const w = weekIndex();
    const kindLabel = { video: "Plant · Leader video", devotional: "Plant · Devotional", challenge: "Water · Couple challenge" };
    return (items.map(({item, data}) => `<section class="card"${data.kind === "challenge" && w === 1 ? ' style="border:2px solid var(--brand-gold)"' : ""}>${label(kindLabel[data.kind])}
        ${data.kind === "video" ? `<div style="display:flex;gap:12px;align-items:center"><div class="video"><svg aria-hidden="true" width="26" height="26" viewBox="0 0 24 24" fill="#C4955F"><path d="M8 5v14l11-7z"/></svg></div><div style="font-weight:700">${esc(item.fields.Title)}</div></div>` : `<div style="font-size:16px;font-weight:700">${esc(item.fields.Title)}</div>`}
        <p class="sub" style="white-space:pre-wrap">${esc(data.text)}</p>${D.safeLink(data.url) ? `<a href="${esc(D.safeLink(data.url))}" target="_blank" rel="noopener noreferrer" style="font-weight:700">Open ${data.kind === "video" ? "video" : "resource"} ›</a>` : ""}</section>`).join("")
      || panel("Week 1 · Plant", `<div style="font-size:16px;font-weight:700">This month’s resources are on the way</div>${empty("The leader video, devotional and couple challenge appear here when the ministry publishes them.")}`))
      + panel("Week 3 · Grow", `<div style="font-size:16px;font-weight:700">Share and pray with your Branch</div><button type="button" data-view="branch" class="linkbtn">Open your Branch ›</button>`)
      + panel("Week 4 · Harvest", `<div style="font-size:16px;font-weight:700">Monthly gathering</div><p class="sub">Finish the month and this fruit is added to your Orchard.</p><button type="button" data-view="orchard" class="linkbtn">See Our Orchard ›</button>`);
  }
  function branch() {
    const couple = profile();
    const issue = state.errors.couples;
    if (issue) return panel("Your Branch", empty(issue));
    return panel("Your Branch", `<div style="font-size:16px;font-weight:700">${esc(couple?.branch ? couple.branch + " Branch" : "Branch not assigned yet")}</div><p class="sub">Your Branch is 4–5 couples with a mentor couple. You meet in Grow week (week 3) to share and pray together.</p>`)
      + panel("Your mentor couple", `<div style="font-size:16px;font-weight:700">${esc(couple?.mentor || "To be announced")}</div><p class="sub">Want to talk something through? Tick “We’d like our mentor couple to reach out” on your Sunday check-in. Only they get a heads-up, with no details.</p><button class="btn" type="button" data-view="checkin">Open the Sunday check-in</button>`);
  }
  function checkin() {
    return `<form data-form="checkin" style="display:flex;flex-direction:column;gap:12px">
      <section class="card" style="gap:14px"><p class="sub" style="font-size:13px">This week, how much did you feel… (1 = not at all, 5 = very)</p>
      ${["Connected", "Heard", "Spiritually together", "Having fun"].map((name, i) => `<fieldset style="border:0;margin:0;padding:0;display:flex;flex-direction:column;gap:6px"><legend style="font-size:15px;font-weight:600;padding:0;margin-bottom:6px">${name}</legend><div class="rates">${[1,2,3,4,5].map(n => `<label class="rate"><input type="radio" name="rating${i}" value="${n}" required><span>${n}</span></label>`).join("")}</div></fieldset>`).join("")}</section>
      <section class="card"><label for="live-need">One thing I need this week</label><input type="text" id="live-need" name="need" maxlength="6000" placeholder="e.g. a quiet evening, just us">
      <label for="live-saw" style="margin-top:6px">Where I saw God at work in you</label><input type="text" id="live-saw" name="saw" maxlength="6000" placeholder="e.g. with the kids at bedtime"></section>
      <section class="card"><label class="tog-check"><input type="checkbox" name="mentor" value="yes"><span>We’d like our mentor couple to reach out</span></label><p class="sub" style="font-size:13px">Optional. Only your mentor couple gets a heads-up, with no details.</p></section>
      <button class="btn" type="submit">Send check-in</button><p role="status" class="form-status sub" style="margin:0"></p></form>`
      + panel("Our check-ins", sharedList("check-in", 8, "Check-ins you and your spouse save will appear here."));
  }
  function journal() {
    if (state.errors.couples || state.errors.reveals) return card("Our journal", empty(state.errors.reveals || state.errors.couples));
    if (!state.journal) return card("Our journal", empty("Your couple profile has not been linked to its private journal folder yet. Ask a ministry administrator to finish setup."));
    const savedEntries = D.parse(state.journal).entries;
    const entries = Array.isArray(savedEntries) ? savedEntries : [];
    return card("Our shared journal", `<p class="sub">Notes added here are visible immediately to both spouses. Site owners and tenant administrators retain administrative access.</p>
      ${entries.map(entry=>`<article style="border-top:1px solid var(--line)"><p class="sub">${esc(entry.date)}</p><p style="white-space:pre-wrap">${esc(entry.text)}</p></article>`).join("") || empty("No shared notes yet.")}
      <form data-form="journal"><label for="journal-note">Add a note together</label><textarea id="journal-note" name="text" maxlength="6000" rows="4" required></textarea><button class="btn" type="submit">Save to our journal</button><p role="status" class="form-status"></p></form>`) + card("All our reveals", sharedList("reveal", 400, "No reveals yet.")) + card("My saved answers", responseList("answers"));
  }
  function pilotJournal() {
    const issue = state.errors.couples || state.errors.reveals;
    if (!state.journal) return card("Our Orchard", empty(issue || "Ask a ministry administrator to link your couple profile."));
    const data = D.parse(state.journal);
    const months = Array.isArray(data.pilotMonths) ? data.pilotMonths : [];
    const curMonth = parts().m;
    const orchard = `<section class="card" style="padding:14px"><div class="orchard">${FRUIT.map((f, index) => {
      const key = "2027-" + String(index + 1).padStart(2, "0");
      const st = months.includes(key) ? "done" : parts().y === 2027 && index + 1 === curMonth ? "now" : "later";
      const fill = st === "done" ? FRUIT_COLOR[index] : st === "now" ? "rgba(196,149,95,.35)" : "none";
      const stroke = st === "done" ? FRUIT_COLOR[index] : st === "now" ? "#C4955F" : "#9AA3AE";
      return `<div class="orchard-cell${st === "now" ? " now" : ""}" role="img" aria-label="${f[0]}, ${f[1]}, ${st === "done" ? "gathered" : st === "now" ? "in progress" : "not yet"}"><svg aria-hidden="true" width="40" height="40" viewBox="0 0 40 40"><circle cx="20" cy="23" r="12" fill="${fill}" stroke="${stroke}" stroke-width="2" stroke-dasharray="${st === "done" ? "0" : "3 3"}"/><path d="M20 11c0-4.5 3.5-7 8-7 0 4.5-3.5 7-8 7z" fill="${st === "later" ? "#9AA3AE" : "#C4955F"}"/></svg><small>${f[0].slice(0,3)}</small><b>${f[1]}</b></div>`;
    }).join("")}</div><p class="sub" style="font-size:13px">A month fills in after either of you answers a prompt or checks in.</p></section>`;
    const goals = D.goals.map(goal => {
      const saved = data.unionGoals?.[goal.key] || "";
      if (day() < goal.date) return card(goal.label, empty("Opens " + goal.date));
      return card(goal.label, `<form data-form="goal"><input type="hidden" name="goal" value="${goal.key}">
        <label for="goal-${goal.key}">Our ${goal.label.toLowerCase()} goal</label>
        <textarea id="goal-${goal.key}" name="text" required maxlength="2000" rows="4">${esc(saved)}</textarea>
        <button class="btn" type="submit">Save goal</button><p class="form-status" role="status"></p></form>`);
    }).join("");
    const released = Array.isArray(data.releasedNotes) ? data.releasedNotes : [];
    const notes = released.map(note => `<article><h3>${esc(note.author)} · ${esc(note.month)}</h3><p style="white-space:pre-wrap">${esc(note.text)}</p></article>`).join("");
    return orchard
      + card("Our Union Goals for 2027", empty("Your three goals for 2027 are shared with your spouse.")) + goals + sealed()
      + card("Released notes", notes || empty("Your notes will appear here after the ministry releases them."));
  }
  function sealed() {
    const config = (state.rows.content || []).map(D.parse).find(row => row.kind === "sealed-release" && row.season === "2027");
    if (state.errors.content || state.errors.pilotIntake) return card("Sealed note", empty(state.errors.content || state.errors.pilotIntake));
    const month = day().slice(0, 7);
    const releaseAt = D.monthlyRelease(month);
    const saved = own("pilotIntake").some(item => item.fields?.Title === `sealed:${month}`);
    if (saved) return card("Sealed note", empty("Your note is sealed. Open Our Orchard to see it after release."));
    if (day() < "2027-01-29") return card("Sealed note", empty("Opens January 29, 2027."));
    if (!config?.publicKey || !Number.isFinite(releaseAt)) return card("Sealed note", empty("The ministry is preparing the release."));
    if (Date.now() >= releaseAt) return card("Sealed note", empty("The submission window has closed."));
    const release = new Intl.DateTimeFormat("en-US", { dateStyle: "long", timeStyle: "short", timeZone: "America/New_York" }).format(new Date(releaseAt));
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
      return panel(event === "retreat" ? "Marriage Retreat 2027" : "December · Year-end", `<div style="font-size:17px;font-weight:700">${title}</div><p class="sub" style="white-space:pre-wrap">${esc(details.text)}</p>
        <form data-form="rsvp"><input type="hidden" name="event" value="${event}">
        <fieldset style="border:0;margin:0;padding:0"><legend class="sub" style="padding:0;margin-bottom:6px">Your RSVP</legend><div class="grid2">${[["yes","We’re going"],["no","Can’t make it"]].map(([value,text]) =>
            `<label class="tog-check"><input type="radio" name="response" value="${value}" required ${selection === value ? "checked" : ""}><span>${text}</span></label>`).join("")}</div></fieldset><button class="btn" type="submit" style="margin-top:10px">Save RSVP</button><p class="form-status sub" role="status"></p></form>`);
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
    const month = day().slice(0, 7);
    const releaseAt = D.monthlyRelease(month);
    if (!text || text.length > 2000 || !config?.publicKey || day() < "2027-01-29" ||
      !(Date.now() < releaseAt) ||
      own("pilotIntake").some(item => item.fields?.Title === `sealed:${month}`)) throw new Error("Sealing is unavailable");
    const publicKey = await crypto.subtle.importKey("jwk", config.publicKey, { name: "RSA-OAEP", hash: "SHA-256" }, false, ["wrapKey"]);
    const key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt"]);
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const bytes = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(text));
    const wrapped = await crypto.subtle.wrapKey("raw", key, publicKey, { name: "RSA-OAEP" });
    const encode = value => btoa(String.fromCharCode(...new Uint8Array(value)));
    await save("pilotIntake", `sealed:${month}`, { version: 1, kind: "sealed-note", season: "2027", month,
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
    const account = state.account, wide = state.view === "leaders" && !!account;
    root.className = "app" + (wide ? " wide" : "");
    const bar = document.getElementById("brandbar"); if (bar && bar !== root) bar.className = "brandbar" + (wide ? " wide" : "");
    const views = { today, checkin, journal, orchard: pilotJournal, sealed, grow, branch, gather, leaders };
    root.innerHTML = header() + accountStrip() + `<main>
      ${!account ? card("Sign in", `<p class="sub">Sign in with your church Microsoft account to open your ministry data. The sign-in screen uses the shared RCI Productions registration.</p>${C.clientId ? button("signin", "Sign in with Microsoft") : empty("Live sign-in is being configured.")}${state.message ? `<p role="status">${esc(state.message)}</p>` : ""}<a href="?demo=1">Explore the sample demo</a>`) :
        `${state.message ? `<div class="banner" role="status">${esc(state.message)}</div>` : ""}${state.loading ? empty("Loading from SharePoint…") : !state.site ? empty("The connection is not ready. Refresh to retry.") : (views[state.view] || today)()}`}
      </main>${account && !wide ? tabBar() : ""}`;
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
