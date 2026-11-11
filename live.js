/* Live data stays in memory and SharePoint. The demo never starts this code. */
(async function () {
  if (window.RCI_APP.DEMO) return;
  const C = window.RCI_CONFIG, D = window.RCIData, esc = D.escape;
  const scopes = ["Sites.ReadWrite.All", "User.Read"];
  const state = { shared: [], account: null, site: null, lists: {}, rows: {}, errors: {}, loading: false, message: "", view: (() => { try { const v = JSON.parse(localStorage.getItem("rci-marriage-view") || "{}"); return typeof v.view === "string" && v.view ? v.view : "today"; } catch { return "today"; } })(), coupleId: (() => { try { return JSON.parse(localStorage.getItem("rci-marriage-view") || "{}").coupleId || null; } catch { return null; } })(), week: D.weekStart(D.ministryDay()) };
  let auth, controller = new AbortController();
  const root = document.getElementById("app");
  const pushApi = "https://rci-push-astmacary-16343.azurewebsites.net/api/push/";
  const pushApp = "rci-marriage";
  const push = { key: null, enabled: false, busy: false, message: "" };
  const pushSupported = () => window.isSecureContext && "Notification" in window && "PushManager" in window && "serviceWorker" in navigator;
  async function pushRegistration() {
    const scope = new URL("./", location.href).href;
    const reg = await navigator.serviceWorker.getRegistration(scope);
    return reg?.scope === scope ? reg : null;
  }
  async function pushRequest(path, data) {
    const options = { cache: "no-store", redirect: "error", signal: AbortSignal.timeout(15000) };
    if (data) {
      const { accessToken } = await auth.acquireTokenSilent({ scopes, account: state.account });
      Object.assign(options, { method: "POST", headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" }, body: JSON.stringify({ appId: pushApp, ...data }) });
    }
    const response = await fetch(pushApi + path, options);
    if (!response.ok) throw new Error("Notification service unavailable");
    return response.json();
  }
  async function checkPush() {
    if (!pushSupported()) return;
    try {
      const reg = await pushRegistration();
      const sub = await reg?.pushManager.getSubscription();
      push.enabled = Boolean(sub && Notification.permission === "granted");
      const { publicKey } = await pushRequest(`config?appId=${pushApp}`);
      const base64 = publicKey.replace(/-/g, "+").replace(/_/g, "/");
      const key = Uint8Array.from(atob(base64), char => char.charCodeAt(0));
      if (key.length !== 65 || key[0] !== 4) throw new Error("Invalid push key");
      push.key = key;
    } catch { push.message = "Couldn’t connect to notifications. Try again."; }
    render();
  }
  async function changePush(enable) {
    if (push.busy || !pushSupported()) return;
    push.busy = true; push.message = "";
    try {
      if (enable && await Notification.requestPermission() !== "granted") {
        push.message = "Notifications are off. You can allow them in your device settings.";
        return;
      }
      const reg = await pushRegistration();
      let sub = await reg?.pushManager.getSubscription();
      if (enable) {
        if (!reg?.active || !push.key) throw new Error("Notifications are not ready");
        sub ||= await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: push.key });
        try { await pushRequest("subscriptions", { action: "upsert", subscription: sub.toJSON() }); }
        catch (error) { await sub.unsubscribe(); throw error; }
      } else {
        if (sub) {
          try { await pushRequest("subscriptions", { action: "remove", endpoint: sub.endpoint }); } catch { /* Unsubscribing invalidates the device endpoint. */ }
          if (!await sub.unsubscribe() && await reg.pushManager.getSubscription()) throw new Error("Unsubscribe failed");
        }
      }
      push.enabled = enable;
    } catch {
      push.message = enable ? "Couldn’t enable notifications. Please try again." : "Couldn’t turn off notifications on this device. Please try again.";
      if (!enable) throw new Error(push.message);
    } finally { push.busy = false; render(); }
  }
  const pushDoneKey = "rci-marriage-push-on";
  const pushDone = () => { try { return localStorage.getItem(pushDoneKey) === "1"; } catch { return false; } };
  function notifications() {
    // Same rule as RCI Security: only an "Enable Notifications" prompt, and only when this device can turn them on.
    if (push.enabled || pushDone()) { if (push.enabled) { try { localStorage.setItem(pushDoneKey, "1"); } catch {} } return ""; }
    if (!pushSupported() || !push.key || Notification.permission === "denied") return "";
    return `<section class="card" style="gap:8px"><h2 class="label" style="margin:0;color:var(--fig)">Lock-Screen Notifications</h2><p class="sub">Receive reminders and ministry updates, even when the app is closed.</p>
      ${push.message ? `<p role="alert" class="sub">${esc(push.message)}</p>` : ""}<button class="btn badge-red" type="button" data-live="push-on" ${push.busy ? "disabled" : ""}>${push.busy ? "Please wait…" : "Enable Notifications"}</button></section>`;
  }

  const siteUrl = `https://${C.siteHost}${C.sitePath}`;
  const longDate = ymd => { const [y, m, d] = String(ymd).split("-").map(Number); return new Intl.DateTimeFormat("en-US", { dateStyle: "long", timeZone: "UTC" }).format(new Date(Date.UTC(y, m - 1, d))); };
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
    if (url.origin !== "https://graph.microsoft.com" || !(url.pathname.startsWith("/v1.0/sites/") || url.pathname === "/v1.0/me/photo/$value")) throw new Error("Unexpected data destination");
    const token = await auth.acquireTokenSilent({ scopes, account: state.account });
    const response = await fetch(url, { ...options, cache: "no-store", signal: controller.signal,
      headers: { Authorization: `Bearer ${token.accessToken}`, "Content-Type": "application/json", ...options.headers } });
    if (!response.ok) { const error = new Error("Microsoft request failed"); error.status = response.status; throw error; }
    if (options.raw) return response;
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
  // Leader videos live in the ministry's Dropbox folder. Each month's video is a Dropbox share link
  // saved as a "Video" row in Monthly Content (dated the 1st of that month).
  const isVideoMonth = value => /^\d{4}-(0[1-9]|1[0-2])$/.test(value);
  const videoRevision = item => item?.eTag || item?.lastModifiedDateTime || item?.id;
  const dropboxLink = value => { try { const u = new URL(String(value).trim()); return u.protocol === "https:" && /(^|\.)dropbox\.com$/i.test(u.hostname) ? u : null; } catch { return null; } };
  const dropboxRaw = value => { const u = dropboxLink(value); if (!u) return ""; u.searchParams.delete("dl"); u.searchParams.set("raw", "1"); return u.href; };
  function loadLeaderVideos() {
    state.leaderVideos = {};
    state.videoIssue = "";
    for (const item of state.rows.content || []) {
      const data = D.read("content", item), month = String(data.date || "").slice(0, 7);
      if (data.kind === "video" && isVideoMonth(month) && dropboxLink(data.url))
        state.leaderVideos[month] = { id: String(item.id), url: data.url, eTag: item.eTag, lastModifiedDateTime: item.lastModifiedDateTime };
    }
  }
  async function addLeaderVideo(month, link) {
    if (!isVideoMonth(month)) throw new Error("Pick a month first.");
    if (!dropboxLink(link)) throw new Error("Paste a Dropbox share link for the video (it starts with https://www.dropbox.com/).");
    if (state.leaderVideos?.[month]) throw new Error("This month already has a leader video.");
    const [y, m] = month.split("-").map(Number);
    await graph(`${base("content")}/items`, { method: "POST", body: JSON.stringify({ fields: {
      Title: `${FRUIT[m - 1][0]} ${y} leader video`, Kind: "Video", PromptDate: D.toDate(month + "-01"), Link: dropboxLink(link).href } }) });
    await load("content"); loadLeaderVideos();
  }
  const watchedTitle = month => `watched-video:${month}`;
  function videoWatched(month, item) {
    return own("pilotIntake").some(row => row.fields?.Title === watchedTitle(month) && D.parse(row).revision === videoRevision(item));
  }
  function watchedEnough(video) {
    const { duration, played } = video;
    if (!Number.isFinite(duration) || duration <= 0) return false;
    let watched = 0;
    for (let i = 0; i < played.length; i++) watched += played.end(i) - played.start(i);
    return watched >= duration * 0.95;
  }
  async function startLeaderVideo(button) {
    const month = button.dataset.month, item = state.leaderVideos?.[month];
    if (!item) return;
    const card = button.closest("section");
    const video = card.querySelector("video"), status = card.querySelector("[role=status]");
    button.disabled = true;
    status.textContent = "Opening video…";
    try {
      const src = D.safeLink(dropboxRaw(item.url));
      if (!src) throw new Error("Video link missing");
      video.src = src;
      video.style.display = "block";
      status.textContent = "";
      video.addEventListener("error", () => {
        status.textContent = "The video couldn’t play. Try opening it again."; button.disabled = false; button.style.display = "";
      }, { once: true });
      button.style.display = "none";
      video.addEventListener("ended", async () => {
        if (videoWatched(month, item)) return;
        if (!watchedEnough(video)) { status.textContent = "Watch the full video to mark it done."; return; }
        if (card.dataset.saving) return;
        card.dataset.saving = "true";
        try {
          await save("pilotIntake", watchedTitle(month), { version: 1, kind: "leader-video-watched", month, revision: videoRevision(item) });
          if (!videoWatched(month, item)) throw new Error("Completion could not be confirmed");
          card.querySelector("[data-video-done]").style.display = "";
          status.textContent = "";
        } catch { status.textContent = "The video finished, but your Done status couldn’t be saved. Refresh and try again."; }
        finally { delete card.dataset.saving; }
      });
      await video.play().catch(() => { status.textContent = "Press play to watch."; });
    } catch (error) { status.textContent = error.status === 403 ? notice(error) : "Couldn’t open the video. Refresh and try again."; }
    finally { button.disabled = false; }
  }
  async function loadView() {
    state.loading = true; render();
    try {
      if (!state.site) {
        state.site = await graph(`sites/${C.siteHost}:${C.sitePath}`);
        const lists = await pages(`sites/${state.site.id}/lists?$select=id,displayName`);
        for (const [key, name] of Object.entries(C.lists)) state.lists[key] = lists.find(list => list.displayName === name)?.id;
      }
      const needed = state.view === "leaders" ? ["couples", "participation", "rsvps", "content", "mentorRequests"]
        : state.view === "mentor" ? ["couples", "participation", "rsvps", "mentorRequests"]
        : state.view === "checkin" ? ["checkins", "couples"]
        : state.view === "journal" ? ["couples"]
        : state.view === "orchard" ? ["couples", "content", "pilotIntake"]
        : state.view === "branch" ? ["couples", "messages"]
        : state.view === "couple" ? ["couples"]
        : state.view === "grow" ? ["content", "couples", "pilotIntake"]
        : ["gather", "sealed"].includes(state.view) ? ["content", "pilotIntake", "couples"] : ["content", "answers", "couples", "pilotIntake"];
      const tasks = needed.map(load);
      if (state.view === "leaders" && state.lists.onboarding) tasks.push(load("onboarding"));
      await Promise.all(tasks);
      if (["leaders", "grow"].includes(state.view)) loadLeaderVideos();
      if (["today", "checkin", "journal", "orchard", "branch"].includes(state.view)) await loadJournal();
    } catch (error) { state.message = notice(error); }
    state.loading = false; render();
  }
  // The signed-in spouse's side of the couple row, and the other spouse's details.
  const mySide = () => D.side(profile(), state.account?.username) || "spouse1";
  const spouseOf = () => { const c = profile(); if (!c) return null; const two = mySide() === "spouse2";
    return { name: two ? c.spouse1Name : c.spouse2Name, photo: two ? c.photo1 : c.photo2 }; };
  const spouseFirst = () => String(spouseOf()?.name || "").trim().split(/\s+/)[0] || "";
  function profile() {
    const email = (state.account?.username || "").toLowerCase();
    return (state.rows.couples || []).map(item => D.read("couples", item)).find(data => data.kind === "couple" && data.active !== false &&
      ((Array.isArray(data.members) && data.members.some(member => member.objectId && String(member.objectId).toLowerCase() === String(oid()).toLowerCase())) ||
        (Array.isArray(data.memberObjectIds) && data.memberObjectIds.includes(oid())) ||
        (Array.isArray(data.memberEmails) && data.memberEmails.some(value => String(value).toLowerCase() === email))));
  }
  // Match Security's opening celebration and keep the trigger within the signed-in couple.
  function celebrateOpening() {
    const couple = profile(), today = day();
    if (!couple || window.matchMedia?.("(prefers-reduced-motion: reduce)")?.matches) return;
    const monthDay = today.slice(5), year = Number(today.slice(0, 4));
    const feb28InNonLeapYear = monthDay === "02-28" && new Date(Date.UTC(year, 1, 29)).getUTCMonth() !== 1;
    const birthday = [couple.birthday1, couple.birthday2].some(date =>
      date?.slice(5) === monthDay || (feb28InNonLeapYear && date?.slice(5) === "02-29"));
    const anniversary = couple.anniversary?.slice(5) === monthDay;
    if (!birthday && !anniversary) return;
    const key = "rci-marriage:opening-celebration";
    try { if (localStorage.getItem(key) === today) return; } catch { /* Storage can be unavailable. */ }
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    try { localStorage.setItem(key, today); } catch { /* Still celebrate on this opening. */ }
    canvas.setAttribute("aria-hidden", "true");
    canvas.style.cssText = "position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:1000";
    const scale = Math.min(window.devicePixelRatio || 1, 2);
    const width = window.innerWidth, height = window.innerHeight;
    canvas.width = width * scale; canvas.height = height * scale;
    ctx.scale(scale, scale);
    document.body.appendChild(canvas);

    const colors = ["#D2B07C", "#BD9863", "#072442", "#FFFFFF", "#1E8A52", "#C62828", "#4A90D9"];
    const randomColor = () => colors[Math.floor(Math.random() * colors.length)];
    const bits = anniversary ? Array.from({ length: 160 }, () => ({
      x: width * (.35 + Math.random() * .3), y: height * .35,
      vx: (Math.random() - .5) * 16, vy: -Math.random() * 18 - 6,
      angle: Math.random() * Math.PI, spin: (Math.random() - .5) * .3,
      size: 6 + Math.random() * 6, color: randomColor()
    })) : [];
    const balloons = birthday ? Array.from({ length: 14 }, (_, i) => ({
      x: width * (.07 + (i + Math.random()) * .86 / 14),
      y: height + 24 + Math.random() * height * .45,
      speed: height * (.33 + Math.random() * .15),
      size: Math.min(18 + Math.random() * 9, width / 11), color: randomColor()
    })) : [];

    const start = performance.now(), duration = birthday ? 4400 : 3500;
    let last = start;
    function frame(now) {
      const age = now - start;
      if (age >= duration) { canvas.remove(); return; }
      const step = Math.min(2, (now - last) / 16.67);
      last = now;
      ctx.clearRect(0, 0, width, height);
      ctx.globalAlpha = Math.min(1, (duration - age) / 850);
      for (const bit of bits) {
        bit.vy += .45 * step; bit.x += bit.vx * step; bit.y += bit.vy * step; bit.angle += bit.spin * step;
        ctx.save();
        ctx.translate(bit.x, bit.y);
        ctx.rotate(bit.angle);
        ctx.fillStyle = bit.color;
        ctx.fillRect(-bit.size / 2, -bit.size / 2, bit.size, bit.size * .65);
        ctx.restore();
      }
      for (const balloon of balloons) {
        const x = balloon.x + Math.sin(age / 440 + balloon.x) * 8;
        const y = balloon.y - balloon.speed * age / 1000;
        const r = balloon.size;
        ctx.beginPath();
        ctx.moveTo(x, y + r);
        ctx.quadraticCurveTo(x + 8, y + r + 10, x - 3, y + r + 25);
        ctx.strokeStyle = "#9AA3AE";
        ctx.stroke();
        ctx.beginPath();
        ctx.ellipse(x, y, r * .75, r, 0, 0, Math.PI * 2);
        ctx.fillStyle = balloon.color;
        ctx.fill();
        ctx.beginPath();
        ctx.ellipse(x - r * .22, y - r * .33, r * .1, r * .2, .4, 0, Math.PI * 2);
        ctx.fillStyle = "rgba(255,255,255,.6)";
        ctx.fill();
      }
      requestAnimationFrame(frame);
    }
    requestAnimationFrame(frame);
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
  const FRUIT = [["January","The Vine","John 15:5"],["February","Love","1 Corinthians 13:4–7"],["March","Joy","Philippians 4:4"],["April","Peace","Colossians 3:15"],
    ["May","Patience","Ephesians 4:2"],["June","Kindness","Ephesians 4:32"],["July","Goodness","Galatians 6:10"],["August","Faithfulness","Proverbs 3:3"],
    ["September","Gentleness","Proverbs 15:1"],["October","Self-control","Proverbs 25:28"],["November","Gratitude","1 Thessalonians 5:18"],["December","Harvest","John 15:8"]];
  const FRUIT_COLOR = ["#2F5D3F","#B23A48","#D99A1E","#3E7CB1","#C4955F","#C77D3A","#5B7F3A","#6B4F9E","#4F8A8B","#8A5A44","#B8860B","#C4955F"];
  const STEPS = ["Plant","Water","Grow","Harvest"];
  const parts = () => { const [y,m,d] = day().split("-").map(Number); return { y, m, d, date: new Date(Date.UTC(y, m - 1, d)) }; };
  const monthFocus = month => (state.rows.content || []).map(D.parse)
    .find(row => row.kind === "monthly-focus" && row.month === month);
  // Monthly Content columns on the month's focus row (scripture, prayer, declaration, reminder, activities, check-in prompts).
  function monthFields(month = day().slice(0, 7)) {
    const rows = (state.rows.content || []).filter(item => D.parse(item).kind === "monthly-focus");
    const row = rows.find(item => D.parse(item).month === month) || rows.find(item => String(D.parse(item).month || "").slice(5) === month.slice(5));
    return row?.fields || {};
  }
  const lines = text => String(text || "").split(/\n+/).map(x => x.trim()).filter(Boolean);
  const paras = text => lines(text).map(x => `<p class="sp" style="margin:0 0 8px;white-space:pre-wrap">${esc(x)}</p>`).join("");
  function checkinPrompts() {
    const mf = monthFields();
    const rates = lines(mf.SundayCheckIns).filter(x => !/^rate\b/i.test(x)).slice(0, 4);
    const qs = lines(mf.SundayCheckInQuestions).slice(0, 6);
    return { rates: rates.length === 4 ? rates : ["Connected", "Heard", "Spiritually together", "Having fun"],
      qs: qs.length ? qs : ["One thing I need this week", "Where I saw God at work in you"] };
  }
  function currentFruit() {
    const fallback = FRUIT[parts().m - 1], planned = monthFocus(day().slice(0, 7));
    return planned ? [fallback[0], planned.focus, planned.verse] : fallback;
  }
  const fmt = (opts) => new Intl.DateTimeFormat("en-US", { timeZone: "UTC", ...opts }).format(parts().date);
  const weekIndex = () => Math.min(3, Math.floor((parts().d - 1) / 7));
  const firstName = () => esc(String(myName() || "").replace(/^pastor\s+/i, "").split(/[\s,]+/).filter(Boolean)[0] || "");
  const label = (text, extra = "") => `<span class="label" style="color:var(--fig)${extra}">${text}</span>`;
  const panel = (lbl, body, extra = "") => `<section class="card"${extra}>${lbl ? label(lbl) : ""}${body}</section>`;
  const tile = (view, lbl, big, small) => `<button type="button" class="card" data-view="${view}" style="gap:4px;text-align:left;cursor:pointer;font:inherit;color:var(--ink)"><span style="font-size:12px;font-weight:700;color:var(--fig)">${lbl}</span><span style="font-size:16px;font-weight:700">${big}</span><span class="sub" style="font-size:13px">${small}</span></button>`;
  const TAB_OF = { calendar: "today", today: "today", checkin: "us", journal: "us", orchard: "us", grow: "us", profile: "profile", mentor: "profile", branch: "branch", gather: "gather" };
  const TAB_ICON = {
    today: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="9" cy="13" r="5"/><circle cx="15" cy="13" r="5"/></svg>',
    us: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M12 20s-7.5-4.6-7.5-10A4.3 4.3 0 0 1 12 7.4 4.3 4.3 0 0 1 19.5 10c0 5.4-7.5 10-7.5 10z"/></svg>',
    branch: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="9" cy="8" r="3"/><circle cx="17" cy="9" r="2.5"/><path d="M3 20c0-3.3 2.7-6 6-6s6 2.7 6 6M15 14.5c3 0 6 2 6 5.5"/></svg>',
    gather: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
    profile: '<svg aria-hidden="true" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8"/></svg>'
  };
  function tabBar() {
    const cur = TAB_OF[state.view];
    const item = (id, view, text) => `<button type="button" data-view="${view}" ${cur === id ? 'aria-current="page"' : ""} ${state.loading ? "disabled" : ""}>${TAB_ICON[id]}${text}</button>`;
    return `<nav class="tabs" aria-label="Main"><div>${item("today","today","Home")}${item("us","grow","Check In")}${item("branch","branch","Branch")}${item("gather","gather","Gatherings")}${item("profile","profile","Profile")}</div></nav>`;
  }
  function usSeg(cur) {
    const a = (view, text) => `<button type="button" data-view="${view}" ${cur === view ? 'aria-current="page"' : ""}>${text}</button>`;
    return `<div class="seg seg4">${a("grow","Grow")}${a("checkin","Check-in")}${a("journal","Journal")}${a("orchard","Our Orchard")}</div>`;
  }
  function header() {
    const v = state.view, f = currentFruit();
    const T = (title, below = "", seg = "") => `<header class="head"><h1>${title}</h1>${below ? `<div class="eyebrow" style="font-weight:400">${below}</div>` : ""}${seg}</header>`;
    const H = (eyebrow, title, top = "") => top ? `<header class="head">${top}<h1 style="margin-top:6px">${title}</h1>${eyebrow ? `<div class="eyebrow" style="font-weight:400">${eyebrow}</div>` : ""}</header>` : `<header class="head">${eyebrow ? `<div class="eyebrow">${eyebrow}</div>` : ""}<h1>${title}</h1></header>`;
    if (!state.account) return T("Welcome");
    if (v === "today") { const n = firstName(); return T(`Good ${new Date().getHours() < 12 ? "morning" : new Date().getHours() < 17 ? "afternoon" : "evening"}${n ? ", " + n : ""}`, `${fmt({ weekday: "long", month: "long", day: "numeric" })} · Week ${weekIndex() + 1} of 4`); }
    if (v === "calendar") { const p = Object.values(monthPrompts()), c = p.filter(x => x.type === "Check-in").length; return H(`${esc(f[2])} · ${p.length - c} prompts · ${c} check-ins`, `${esc(f[0])} · ${esc(f[1])}`); }
    if (v === "checkin") return T("Sunday Check In", `Your responses will be shared with ${spouseFirst() ? esc(spouseFirst()) : "your spouse"}.`, usSeg("checkin"));
    if (v === "orchard") return T("Our Orchard", "Fills in automatically month by month.", usSeg("orchard"));
    if (v === "journal") return T("Our Journal", `Only visible to you and ${spouseFirst() ? esc(spouseFirst()) : "your spouse"}.`, usSeg("journal"));
    if (v === "grow") return T(`${f[1]}, Week by Week`, `${f[0]} · ${f[2]}`, usSeg("grow"));
    if (v === "branch") return T(esc(profile()?.branch ? profile().branch + " Branch" : "Branch"));
    if (v === "couple") { const c = (state.rows.couples || []).find(item => String(item.id) === String(state.coupleId)); return T(esc(c?.fields?.Title || "Couple")); }
    if (v === "gather") return T("Gatherings &amp; Special Dates");
    if (v === "profile") return T(`<span style="display:flex;justify-content:space-between;align-items:center;gap:10px">My Profile${isMentor(myCouple()) || isLeader() ? `<button type="button" class="icon-round" data-view="mentor" aria-label="Mentor dashboard" title="Mentor dashboard" style="margin-left:auto;font-size:18px;color:var(--fig)">★</button>` : ""}${isLeader() ? `<button type="button" class="icon-round" data-view="leaders" aria-label="Leader view" title="Leader view"><svg aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="7" r="3"/><circle cx="5" cy="10" r="2.2"/><circle cx="19" cy="10" r="2.2"/><path d="M7.5 20v-1.5a4.5 4.5 0 0 1 9 0V20M1.5 20v-1a3.5 3.5 0 0 1 4.6-3.3M22.5 20v-1a3.5 3.5 0 0 0-4.6-3.3"/></svg></button>` : ""}</span>`);
    if (v === "mentor") return T(`<span class="dash-title">How ${esc(myCouple()?.branch || "Your Branch")} Is Growing</span>`, `<span style="text-transform:uppercase;letter-spacing:.06em;font-weight:700;color:var(--fig)">Mentor dashboard · ${esc(fmt({ month: "long", year: "numeric" }))} · ${esc(f[1])}</span><br><em style="font-size:13px">Participation only, never what couples write.</em>`, `<button type="button" class="linkbtn" data-view="profile" style="color:var(--cream);min-height:32px">‹ Back to couple view</button>`);
    if (v === "leaders") return T(`<span style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px"><span class="dash-title">How The Ministry Is Growing</span>${state.lists.onboarding && !state.loading && !state.errors.onboarding ? button("couple-onboard-open", "+ Add Couple") : ""}</span>`, `<span style="text-transform:uppercase;letter-spacing:.06em;font-weight:700;color:var(--fig)">Leader dashboard · ${esc(fmt({ month: "long", year: "numeric" }))} · ${esc(f[1])}</span><br><em style="font-size:13px">Participation only, never what couples write.</em>`, `<button type="button" class="linkbtn" data-view="profile" style="color:var(--cream);min-height:32px">‹ Back to couple view</button>`);
    return H("", "RCI Marriage");
  }
  function today() {
    const content = (state.rows.content || []).map(item => ({ item, data: D.read("content", item) }));
    const prompt = content.find(({data}) => data.kind === "prompt" && data.date === day());
    const f = currentFruit(), w = weekIndex();
    const hero = `<section class="month" aria-label="This month"><div class="row"><span class="label">${esc(f[0])} fruit</span><span style="font-size:12px;color:#C9D2DD;letter-spacing:.08em">FRUIT OF OUR UNION</span></div>
      <div class="name">${esc(f[1])}</div><div style="font-size:15px;line-height:1.4;color:#C9D2DD">${esc(f[2])}</div>
</section>`;
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
  // Month calendar of published prompts, the same as the demo's January preview.
  const CAL_TYPES = { Reflect:["#DCE6F2","#1F4A70"], Faith:["#E3ECE2","#1F4A2E"], Fun:["#FDF1D8","#6B4800"],
    Appreciate:["#F4ECDF","#7A5526"], Plan:["#E4E1EE","#3D3366"], "Check-in":["#C4955F","#0A2440"] };
  const calPill = t => CAL_TYPES[t] ? `<span class="pill" style="background:${CAL_TYPES[t][0]};color:${CAL_TYPES[t][1]}">${esc(t)}</span>` : "";
  function monthPrompts() {
    const { y, m } = parts(), prefix = `${y}-${String(m).padStart(2, "0")}-`;
    const byDay = {};
    for (const item of state.rows.content || []) {
      const data = D.read("content", item);
      if (data.kind === "prompt" && String(data.date || "").startsWith(prefix)) byDay[Number(data.date.slice(8, 10))] = data;
    }
    return byDay;
  }
  function calendar() {
    if (state.errors.content) return panel("Prompts", empty(state.errors.content));
    const { y, m, d: todayD } = parts(), byDay = monthPrompts();
    const days = new Date(Date.UTC(y, m, 0)).getUTCDate(), lead = new Date(Date.UTC(y, m - 1, 1)).getUTCDay();
    const sel = state.calDay && state.calDay <= days ? state.calDay : todayD;
    const p = byDay[sel];
    const cells = Array.from({ length: days }, (_, i) => {
      const n = i + 1, t = byDay[n]?.type, c = CAL_TYPES[t] || ["var(--soft)", "var(--muted)"];
      return `<button type="button" data-cal="${n}" aria-pressed="${n === sel}" aria-label="${esc(FRUIT[m - 1][0])} ${n}${t ? ", " + esc(t) : ", no prompt"}" style="background:${c[0]};color:${c[1]}">${n}</button>`;
    }).join("");
    const dow = new Intl.DateTimeFormat("en-US", { timeZone: "UTC", weekday: "short" }).format(new Date(Date.UTC(y, m - 1, sel)));
    const wk = sel <= 7 ? "Week 1 · Plant" : sel <= 14 ? "Week 2 · Water" : sel <= 21 ? "Week 3 · Grow" : "Week 4 · Harvest";
    let note = "";
    if (p) {
      if (/union goal/i.test(p.text)) note = "Saved to Our Orchard as one of your Union Goals.";
      if (p.type === "Check-in") note = "No prompt on Sundays. Do the check-in instead.";
    }
    return `<div class="chips">${Object.keys(CAL_TYPES).map(calPill).join("")}</div>
      <section class="card" style="padding:12px"><div class="cal">${["S","M","T","W","T","F","S"].map(x => `<div class="dow">${x}</div>`).join("")}${"<div></div>".repeat(lead)}${cells}</div></section>
      <section class="card"><div class="row"><span class="label" style="color:var(--muted);letter-spacing:.06em">${dow}, ${esc(FRUIT[m - 1][0].slice(0, 3))} ${sel} · ${wk}</span>${p ? calPill(p.type) : ""}</div>
        <div style="font-size:16px;font-weight:600;line-height:1.4">${p ? esc(p.text) : "No prompt published for this day."}</div>
        ${note ? `<div style="font-size:13px;line-height:1.4;color:var(--fig);font-weight:600">${note}</div>` : ""}</section>`;
  }
  function grow() {
    if (state.errors.content) return panel("Grow", empty(state.errors.content));
    const content = (state.rows.content || []).map(item => ({ item, data: D.read("content", item) }));
    const order = ["video", "devotional", "challenge"];
    const schedule = content.map(({data}) => data)
      .filter(data => data.kind === "monthly-focus" && /^\d{4}-(0[1-9]|1[0-2])$/.test(data.month))
      .sort((a, b) => a.month.localeCompare(b.month));
    const currentMonth = day().slice(0, 7);
    const selected = schedule.find(data => data.month === currentMonth)
      || schedule.find(data => data.month.slice(5) === currentMonth.slice(5)) || schedule[0];
    const leaderVideo = state.leaderVideos?.[selected?.month];
    const items = content.filter(({data}) => order.includes(data.kind) && !(leaderVideo && data.kind === "video") &&
      (!selected || (data.kind !== "challenge" && (!data.date || data.date.startsWith(selected.month + "-")))))
      .sort((a, b) => order.indexOf(a.data.kind) - order.indexOf(b.data.kind));
    const w = !selected || selected.month === currentMonth ? weekIndex() : -1;
    const monthLabel = month => new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" })
      .format(new Date(month + "-01T00:00:00Z"));
    const kindLabel = { video: "Plant · Leader video", devotional: "Plant · Devotional", challenge: "Water · Couple challenge" };
    const weekOf = { video: 0, devotional: 0, challenge: 1 };
    const now = i => i === w ? ' style="position:relative;border:2px solid var(--brand-gold)"' : "";
    const tag = i => i === w ? `<span class="pill" style="position:absolute;top:10px;right:10px;background:var(--brand-gold);color:var(--navy);font-weight:700">This week</span>` : "";
    const fr = FRUIT[parts().m - 1], wk = weekIndex();
    const focus = `<section class="month" aria-label="This month’s focus"><div class="row"><span class="label">${esc(fr[0])} focus</span><span style="font-size:12px;color:#C9D2DD;letter-spacing:.08em">FRUIT OF OUR UNION</span></div>
      <div class="name">${esc(selected?.focus || fr[1])}</div><div style="font-size:15px;line-height:1.4;color:#C9D2DD">${esc(selected?.verse || fr[2])}</div>
      <div class="rhythm">${STEPS.map((step, i) => `<div class="${i < wk ? "done" : i === wk ? "now" : ""}">${step}${i < wk ? " ✓" : ""}</div>`).join("")}</div></section>`;
    const mf = monthFields(selected?.month || day().slice(0, 7));
    const week1Extra = (mf.MonthlyScripture ? `<div style="font-style:italic;line-height:1.5">${paras(mf.MonthlyScripture)}</div>` : "")
      + (mf.MonthlyReminder ? `<p class="sp" style="margin:0"><b>Monthly reminder:</b> ${esc(mf.MonthlyReminder.replace(/^remember:\s*/i, ""))}</p>` : "");
    const videoCard = leaderVideo || week1Extra ? panel("Week 1 · Plant", leaderVideo ? 
      `<span class="pill cs-live" data-video-done style="position:absolute;top:10px;right:10px${videoWatched(selected.month, leaderVideo) ? "" : ";display:none"}">Done ✓</span>
      <div style="font-size:16px;font-weight:700">Devotional Video</div>
      <button type="button" class="btn" data-live="watch-leader-video" data-month="${esc(selected.month)}">Watch devotional video</button>
      <video controls playsinline preload="none" aria-label="Leader video" style="display:none;width:100%;max-height:420px"></video><p class="sub" role="status"></p>${week1Extra}` : `${tag(0)}${week1Extra}`,
      w === 0 ? now(0) : ' style="position:relative"') : "";
    const card1 = ({item, data}) => `<section class="card"${now(weekOf[data.kind])}>${tag(weekOf[data.kind])}${label(kindLabel[data.kind])}
        ${data.kind === "video" ? `<div style="display:flex;gap:12px;align-items:center"><div class="video"><svg aria-hidden="true" width="26" height="26" viewBox="0 0 24 24" fill="#C4955F"><path d="M8 5v14l11-7z"/></svg></div><div style="font-weight:700">${esc(item.fields.Title)}</div></div>` : `<div style="font-size:16px;font-weight:700">${esc(item.fields.Title)}</div>`}
        <p class="sub" style="white-space:pre-wrap">${esc(data.text)}</p>${D.safeLink(data.url) ? `<a href="${esc(D.safeLink(data.url))}" target="_blank" rel="noopener noreferrer" style="font-weight:700">Open ${data.kind === "video" ? "video" : "resource"} ›</a>` : ""}</section>`;
    const week0 = items.filter(x => weekOf[x.data.kind] === 0).map(card1).join("");
    const week1Items = items.filter(x => weekOf[x.data.kind] === 1).map(card1).join("");
    // Weeks open on the month's Sundays: Week 1 on the 1st Sunday, Week 2 on the 2nd, and so on.
    const sundays = (() => { const { y, m } = parts(), out = []; for (let d = 1; d <= 31; d++) { const dt = new Date(Date.UTC(y, m - 1, d)); if (dt.getUTCMonth() !== m - 1) break; if (dt.getUTCDay() === 0) out.push(dt.toISOString().slice(0, 10)); } return out; })();
    const opens = i => sundays[i] || sundays[sundays.length - 1];
    // Week 4 text comes from the Monthly Content "Harvest Discussion" column (matched by name, since SharePoint may encode it).
    const harvestKey = Object.keys(mf || {}).find(k => /harvest/i.test(k) && /discussion/i.test(k)) || Object.keys(mf || {}).find(k => /harvest/i.test(k) && typeof mf[k] === "string");
    const harvestDisc = harvestKey ? mf[harvestKey] : "";
    // Temporarily show Week 4 while its content is being edited.
    const week4EditingPreview = Date.now() < Date.parse("2026-11-12T01:26:16Z");
    const gate = (i, html) => !html ? "" : (i === 3 && week4EditingPreview) || day() >= opens(i) ? html
      : panel(`Week ${i + 1} · ${STEPS[i]}`, `<p class="sub" style="margin:0">Opens Sunday, ${esc(longDate(opens(i)).replace(/, \d{4}$/, ""))}.</p>`);
    return focus + gate(0, videoCard + (week0 || (leaderVideo || week1Extra ? "" : panel("Week 1 · Plant", `${tag(0)}<div style="font-size:16px;font-weight:700">This month’s resources are on the way</div>${empty("The leader video, devotional and couple challenge appear here when the ministry publishes them.")}`, now(0)))))
      + gate(1, week1Items + (mf.InteractiveActivity || selected ? panel("Week 2 · Water", `${tag(1)}<div style="font-size:16px;font-weight:700">Interactive for couples</div>${mf.InteractiveActivity ? `<div style="line-height:1.5">${paras(mf.InteractiveActivity)}</div>` : empty(selected.water)}
          <p class="sub" style="font-style:italic;font-size:13px;margin:0">Choose at least one activity to do together this month.</p>`, now(1)) : ""))
      + gate(2, panel("Week 3 · Grow", `${tag(2)}<div style="font-size:16px;font-weight:700">Share and pray with ${spouseFirst() ? esc(spouseFirst()) : "your spouse"}</div>
          ${mf.MonthlyPrayer ? `<div><span class="label" style="color:var(--fig)">Monthly prayer</span><div style="line-height:1.5;margin-top:6px">${paras(mf.MonthlyPrayer)}</div></div>` : ""}
          ${mf.MonthlyDeclaration ? `<div><span class="label" style="color:var(--fig)">Monthly declaration</span><div style="line-height:1.5;margin-top:6px;font-weight:600">${paras(mf.MonthlyDeclaration)}</div></div>` : ""}`, now(2)))
      + gate(3, panel("Week 4 · Harvest", `${tag(3)}${harvestDisc ? `<div style="font-size:16px;font-weight:700">Harvest discussion</div><div style="line-height:1.5">${paras(harvestDisc)}</div><hr style="border:0;border-top:1px solid var(--soft);margin:4px 0">` : ""}<div style="font-size:16px;font-weight:700">Monthly gathering</div><p class="sub">Finish the month and this fruit is added to your Orchard.</p><button type="button" data-view="gather" class="linkbtn">Open the Gatherings calendar ›</button>`, now(3)));
  }
  function branch() {
    const couple = profile();
    const issue = state.errors.couples;
    if (issue) return panel("Your Branch", empty(issue));
    const initials = c => [c.spouse1Name, c.spouse2Name].map(n => String(n || "?").trim()[0] || "?").join("&");
    const names = c => [c.spouse1Name, c.spouse2Name].map(n => String(n || "").split(/\s+/)[0]).filter(Boolean).join(" & ") || "Couple";
    const mine = couple?.coupleKey;
    const branchNames = [...new Set(allCouples().map(c => c.branch).filter(Boolean))].sort();
    const viewing = isLeader() ? (state.branchView ?? "all") : (couple?.branch || "");
    const mates = viewing ? (state.rows.couples || []).map(item => D.read("couples", item)).filter(c => c.kind === "couple" && c.active !== false && (viewing === "all" || c.branch === viewing))
      .sort((a, b) => (b.coupleKey === mine) - (a.coupleKey === mine)) : [];
    const all = (state.rows.couples || []).map(item => D.read("couples", item)).filter(c => c.kind === "couple" && c.active !== false)
      .sort((a, b) => (b.coupleKey === mine) - (a.coupleKey === mine) || names(a).localeCompare(names(b)));
    const mentorIds = new Set(allCouples().filter(isMentor).map(m => m.itemId));
    const circle = c => `<button type="button" class="av-btn${c.coupleKey === mine ? " me" : ""}${mentorIds.has(c.itemId) ? " mentor" : ""}" data-couple="${esc(c.itemId)}">${mentorIds.has(c.itemId) ? STAR : ""}<span><img alt="" data-photo="${esc(c.couplePhoto || "")}" hidden><b>${esc(initials(c))}</b></span>${c.coupleKey === mine ? "You" : esc(names(c))}</button>`;
    const ministry = panel(`The ministry · ${all.length} couple${all.length === 1 ? "" : "s"}`, all.length ? `<div class="avatars big">${all.map(circle).join("")}</div>` : empty("Couples will appear here as they join."));
    const picker = isLeader() ? `<label for="branch-view" class="sub">Viewing</label><select id="branch-view"><option value="all"${viewing === "all" ? " selected" : ""}>All Branches</option>${branchNames.map(b => `<option value="${esc(b)}"${viewing === b ? " selected" : ""}>${esc(b)}</option>`).join("")}</select>` : "";
    const yours = panel(isLeader() ? "Branches" : "Your Branch", `${picker}<div style="font-size:16px;font-weight:700">${esc(viewing === "all" ? "All Branches" : viewing ? viewing + " Branch" : "Branch not assigned yet")}</div>
      ${mates.length ? `<div class="avatars big">${mates.map(circle).join("")}</div>` : empty("Your Branch couples will appear here once you’re assigned.")}
      <p class="sub">Your Branch is 4–5 couples with a mentor couple.</p>`);
    return ministry + yours
      + messagesCard(viewing)
      + panel("Your mentor couple", `<div style="font-size:16px;font-weight:700">${esc((() => { const full = allCouples().find(x => x.itemId === couple?.itemId); const m = full && mentorOf(full); return m ? pairNameOf(m) : "To be announced"; })())}</div><p class="sub">Want to talk something through? Tick “We’d like our mentor couple to reach out” on your Sunday check-in. Only they get a heads-up, with no details.</p>${state.mentorSent ? `<p class="sub" style="font-weight:600;color:var(--fig)">Request sent. Your mentor couple will reach out.</p>` : `<button class="btn" type="button" data-live="mentor-request">We’d like our mentor couple to reach out</button>`}`);
  }
  // Branch messages (like RCI Security): a "Branch Messages" list; the app shows only your Branch's posts.
  const REACTS = [["heart", "❤️"], ["up", "👍"], ["down", "👎"]];
  function messagesCard(branchName) {
    if (!state.lists.messages || !branchName) return "";
    if (state.errors.messages) return panel("Branch messages", empty(state.errors.messages));
    const all = (state.rows.messages || []).filter(item => branchName === "all" || String(item.fields?.Branch || "") === String(branchName));
    const tag = item => String(item.fields?.Title || "");
    const posts = all.filter(item => !/^(reply|react):/.test(tag(item))).sort((a, b) => String(b.createdDateTime || "").localeCompare(String(a.createdDateTime || ""))).slice(0, 30);
    const me = String(state.account?.username || "").toLowerCase();
    const mineItem = item => String(item.createdBy?.user?.email || "").toLowerCase() === me;
    const when = at => at ? new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", hour: "numeric", minute: "2-digit", timeZone: "America/New_York" }).format(new Date(at)) : "";
    const who = item => esc(item.fields?.AuthorName || item.createdBy?.user?.displayName || "Someone");
    const del = item => mineItem(item) ? `<button type="button" class="linkbtn" data-live="delete-message" data-id="${esc(item.id)}" style="font-size:13px;min-height:28px">Delete</button>` : "";
    const post = item => {
      const id = String(item.id);
      const replies = all.filter(x => tag(x) === "reply:" + id).sort((a, b) => String(a.createdDateTime).localeCompare(String(b.createdDateTime)));
      const reacts = REACTS.map(([k, emo]) => { const r = all.filter(x => tag(x) === `react:${id}:${k}`), mine = r.find(mineItem);
        return `<button type="button" data-live="react" data-id="${esc(id)}" data-kind="${k}" data-branch="${esc(item.fields?.Branch || "")}" ${mine ? `data-mine="${esc(mine.id)}"` : ""} aria-pressed="${mine ? "true" : "false"}" aria-label="${k === "heart" ? "Heart" : k === "up" ? "Thumbs up" : "Thumbs down"}">${emo}${r.length ? " " + r.length : ""}</button>`; }).join("");
      return `<article class="msg-row"><div class="row"><b>${who(item)}</b><span class="sub" style="font-size:12px">${branchName === "all" ? esc(item.fields?.Branch || "") + " · " : ""}${esc(when(item.createdDateTime))}</span></div>
        <p style="margin:4px 0 0;white-space:pre-wrap;overflow-wrap:anywhere">${esc(item.fields?.Message || "")}</p>
        <div class="react">${reacts}</div>
        ${replies.map(r => `<div class="msg-reply"><div class="row"><b style="font-size:14px">${who(r)}</b><span class="sub" style="font-size:12px">${esc(when(r.createdDateTime))}</span></div><p style="margin:2px 0 0;white-space:pre-wrap;overflow-wrap:anywhere;font-size:14px">${esc(r.fields?.Message || "")}</p>${del(r)}</div>`).join("")}
        <details style="margin-top:6px"><summary class="linkbtn" style="font-size:13px;cursor:pointer;min-height:28px">Reply</summary>
          <form data-form="branch-reply" style="display:flex;flex-direction:column;gap:6px;margin-top:6px"><input type="hidden" name="parent" value="${esc(id)}"><input type="hidden" name="branch" value="${esc(item.fields?.Branch || "")}"><textarea name="text" rows="2" maxlength="2000" required aria-label="Write a reply"></textarea><button class="btn small-btn" type="submit" style="align-self:flex-end">Reply</button><p class="form-status sub" role="status" style="margin:0"></p></form></details>
        ${del(item)}</article>`;
    };
    const canPost = branchName !== "all";
    return panel(branchName === "all" ? "All Branch messages" : `${esc(branchName)} messages`, `${canPost ? `<form data-form="branch-message" style="display:flex;flex-direction:column;gap:8px"><input type="hidden" name="branch" value="${esc(branchName)}"><textarea name="text" rows="2" maxlength="2000" required aria-label="Write a message to your Branch" placeholder="Write a message to your Branch…"></textarea><button class="btn" type="submit" style="align-self:flex-end">Post</button><p class="form-status sub" role="status" style="margin:0"></p></form>` : `<p class="sub" style="font-size:13px;margin:0">Choose a Branch above to post a message.</p>`}
      ${posts.map(post).join("") || empty("No messages yet. Say hello to your Branch.")}
      <p class="sub" style="font-size:12px;font-style:italic;margin:0">Only couples in this Branch see these messages in the app.</p>`);
  }
  function checkin() {
    const cp = checkinPrompts();
    return `<form data-form="checkin" style="display:flex;flex-direction:column;gap:12px">
      <section class="card" style="gap:14px">${label("How we’re doing this week")}<p class="sub" style="font-size:15px;font-weight:600;color:var(--ink)">How was this week?</p><p class="sub" style="font-size:13px;margin-top:-8px">1 = not at all, 5 = very</p>
      ${cp.rates.map((name, i) => `<fieldset style="border:0;margin:0;padding:0;display:flex;flex-direction:column;gap:6px"><legend style="font-size:15px;font-weight:600;padding:0;margin-bottom:6px">${esc(name)}</legend><div class="rates">${[1,2,3,4,5].map(n => `<label class="rate"><input type="radio" name="rating${i}" value="${n}" required><span>${n}</span></label>`).join("")}</div></fieldset>`).join("")}</section>
      <section class="card" style="gap:12px">${label("This week’s questions")}${cp.qs.map((q, i) => `<label for="live-q${i}">${esc(q)}</label><textarea id="live-q${i}" name="q${i}" maxlength="6000" rows="2"></textarea>`).join("")}</section>
      <button class="btn" type="submit">Send check-in</button><p role="status" class="form-status sub" style="margin:0"></p>
      <section class="card"><label class="tog-check"><input type="checkbox" name="mentor" value="yes"><span>We’d like our mentor couple to reach out</span></label><p class="sub" style="font-size:13px">Optional. Only your mentor couple gets a heads-up, with no details.</p></section>
      </form>`
      + panel("Our Check-In History", sharedList("check-in", 8, "Check-ins you and your spouse save will appear here."));
  }
  function journal() {
    if (state.errors.couples || state.errors.reveals) return card("Our Journal", empty(state.errors.reveals || state.errors.couples));
    if (!state.journal) return card("Our Journal", empty("Your couple profile has not been linked to its private journal folder yet. Ask a ministry administrator to finish setup."));
    const savedEntries = D.parse(state.journal).entries;
    const entries = Array.isArray(savedEntries) ? savedEntries : [];
    return card("Our shared journal", `<p class="sub">Notes added here are only visible to the two of you.</p>
      ${entries.map(entry=>`<article style="border-top:1px solid var(--line)"><p class="sub">${esc(entry.date)}</p><p style="white-space:pre-wrap">${esc(entry.text)}</p></article>`).join("")}
      <form data-form="journal"><textarea id="journal-note" aria-label="Write a note to each other" placeholder="Write a note to each other…" name="text" maxlength="6000" rows="4" required></textarea><button class="btn" type="submit">Save to our journal</button><p role="status" class="form-status"></p></form>`) + card("All our reveals", sharedList("reveal", 400, "No reveals yet."));
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
      return `<div class="orchard-cell${st === "now" ? " now" : ""}" role="img" aria-label="${f[0]}, ${f[1]}, ${st === "done" ? "gathered" : st === "now" ? "in progress" : "not yet"}"><svg aria-hidden="true" width="40" height="40" viewBox="0 0 40 40"><circle cx="20" cy="23" r="12" fill="${fill}" stroke="${stroke}" stroke-width="2" stroke-dasharray="${st === "done" ? "0" : "3 3"}"/><path d="M20 11c0-4.5 3.5-7 8-7 0 4.5-3.5 7-8 7z" fill="${st === "later" ? "#9AA3AE" : "#C4955F"}"/></svg><small>${f[0].slice(0,3)}</small><b>${esc(monthFocus(key)?.focus || f[1])}</b></div>`;
    }).join("")}</div><p class="sub" style="font-size:13px">A month fills in after either of you answers a prompt or checks in.</p></section>`;
    const goals = D.goals.map(goal => {
      const saved = data.unionGoals?.[goal.key] || "";
      if (day() < goal.date) return card(goal.label, empty("Opens " + longDate(goal.date) + "."));
      return card(goal.label, `<form data-form="goal"><input type="hidden" name="goal" value="${goal.key}">
        <label for="goal-${goal.key}">Our ${goal.label.toLowerCase()} goal</label>
        <textarea id="goal-${goal.key}" name="text" required maxlength="2000" rows="4">${esc(saved)}</textarea>
        <button class="btn" type="submit">Save goal</button><p class="form-status" role="status"></p></form>`);
    }).join("");
    const released = Array.isArray(data.releasedNotes) ? data.releasedNotes : [];
    const notes = released.map(note => `<article><h3>${esc(note.author)} · ${esc(note.month)}</h3><p style="white-space:pre-wrap">${esc(note.text)}</p></article>`).join("");
    return orchard;
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
  // Gatherings calendar, laid out like the Security app's calendar: tap a day to see details and RSVP.
  function gatherings() {
    return (state.rows.content || []).map(item => D.read("content", item)).filter(row => row.kind === "gathering" && /^\d{4}-\d{2}-\d{2}$/.test(String(row.date || "")))
      .map(row => ({ ...row, end: /^\d{4}-\d{2}-\d{2}$/.test(String(row.end || "")) && row.end > row.date ? row.end : row.date })).sort((a, b) => a.date.localeCompare(b.date));
  }
  const ymdAdd = (ymd, n) => { const [y, m, d] = ymd.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); };
  function rsvpForm(event) {
    const details = (state.rows.content || []).map(D.parse).find(row => row.kind === "event" && row.event === event && row.season === "2027");
    if (!details) return empty("RSVPs open closer to the date.");
    const saved = own("pilotIntake").find(item => item.fields?.Title === "rsvp:2027:" + event);
    const selection = saved ? D.parse(saved).response : "";
    return `<form data-form="rsvp"><input type="hidden" name="event" value="${event}">
      <fieldset style="border:0;margin:0;padding:0"><legend class="sub" style="padding:0;margin-bottom:6px">Your RSVP${selection ? ` · saved: ${selection === "yes" ? "We’re going" : "Can’t make it"}` : ""}</legend><div class="grid2">${[["yes","We’re going"],["no","Can’t make it"]].map(([value,text]) =>
        `<label class="tog-check"><input type="radio" name="response" value="${value}" required ${selection === value ? "checked" : ""}><span>${text}</span></label>`).join("")}</div></fieldset><button class="btn" type="submit" style="margin-top:10px">Save RSVP</button><p class="form-status sub" role="status"></p></form>`;
  }
  function gather() {
    if (state.errors.content || state.errors.pilotIntake) return card("Gather", empty(state.errors.content || state.errors.pilotIntake));
    const list = gatherings(), t = day();
    const people = (state.rows.couples || []).map(item => D.read("couples", item)).filter(c => c.kind === "couple" && c.active !== false);
    const first = n => String(n || "").trim().split(/\s+/)[0];
    const yearly = ymd => people.flatMap(c => [
      c.birthday1 && c.birthday1.slice(5) === ymd.slice(5) ? { kind: "birthday", text: `${first(c.spouse1Name) || "Spouse"}’s birthday` } : null,
      c.birthday2 && c.birthday2.slice(5) === ymd.slice(5) ? { kind: "birthday", text: `${first(c.spouse2Name) || "Spouse"}’s birthday` } : null,
      c.anniversary && c.anniversary.slice(5) === ymd.slice(5) ? { kind: "anniversary", text: `${[first(c.spouse1Name), first(c.spouse2Name)].filter(Boolean).join(" & ") || "Couple"}’s anniversary`, years: Number(ymd.slice(0, 4)) - Number(c.anniversary.slice(0, 4)) } : null
    ]).filter(Boolean);
    const on = ymd => list.filter(g => g.date <= ymd && ymd <= g.end);
    if (!state.gcal) state.gcal = { month: t.slice(0, 7), sel: t };
    const { month, sel } = state.gcal, [y, m] = month.split("-").map(Number);
    const lead = new Date(Date.UTC(y, m - 1, 1)).getUTCDay(), dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
    const start = ymdAdd(`${month}-01`, -lead), weeks = Math.ceil((lead + dim) / 7);
    let cells = "";
    for (let n = 0; n < weeks * 7; n++) {
      const k = ymdAdd(start, n), ev = on(k), yr = yearly(k);
      const marks = (ev.length ? "<i></i>" : "") + (yr.some(x => x.kind === "birthday") ? '<span aria-hidden="true">🎂</span>' : "") + (yr.some(x => x.kind === "anniversary") ? '<span aria-hidden="true">💍</span>' : "");
      cells += `<button type="button" class="dcell${k.slice(0, 7) !== month ? " out" : ""}${k === t ? " today" : ""}${k === sel ? " sel" : ""}" data-gday="${k}" aria-label="${esc(longDate(k))}${[...ev, ...yr].length ? ", " + esc([...ev, ...yr].map(g => g.text).join(", ")) : ""}"><span>${Number(k.slice(8))}</span><span class="dots">${marks}</span></button>`;
    }
    const title = new Intl.DateTimeFormat("en-US", { month: "long", year: "numeric", timeZone: "UTC" }).format(new Date(Date.UTC(y, m - 1, 1)));
    const today = on(sel);
    const yrSel = yearly(sel);
    const celebrate = yrSel.length ? `<section class="card" style="gap:8px">${yrSel.map(x => `<div style="display:flex;align-items:center;gap:10px"><span aria-hidden="true" style="font-size:20px">${x.kind === "birthday" ? "🎂" : "💍"}</span><span style="font-weight:600">${esc(x.text)}${x.kind === "anniversary" && x.years > 0 ? ` <span class="sub" style="font-weight:400">· ${x.years} year${x.years === 1 ? "" : "s"}</span>` : ""}</span></div>`).join("")}</section>` : "";
    const detail = celebrate + (today.length ? today.map(g => `<section class="card"><span class="label" style="color:var(--fig)">${esc(g.when || longDate(g.date))}</span>
        <div style="font-size:17px;font-weight:700">${esc(g.text)}</div>
        ${g.details ? `<p class="sub" style="white-space:pre-wrap">${esc(g.details)}</p>` : `<p class="sub">Time and place to be announced.</p>`}
        ${D.events.includes(g.event) ? rsvpForm(g.event) : `<p class="sub" style="font-size:13px">RSVPs open closer to the date.</p>`}</section>`).join("")
      : celebrate ? "" : `<section class="card"><p class="sub" style="text-align:center;margin:0">Nothing on this date.</p></section>`);
    const upcoming = list.map(g => `<button type="button" class="card" data-gday="${g.date}" style="flex-direction:row;justify-content:space-between;align-items:center;cursor:pointer;font:inherit;color:var(--ink);text-align:left;padding:12px 16px"><span style="font-weight:700">${esc(g.text)}</span><span class="sub">${esc(g.when || longDate(g.date))}</span></button>`).join("");
    return `<section class="card cal-panel"><div class="cal-head"><h2>${esc(title)}</h2>
        <button type="button" class="btn ghost" data-gcal="today" style="padding:6px 10px;min-height:0">Today</button>
        <button type="button" class="iconbtn" data-gcal="prev" aria-label="Previous month">‹</button><button type="button" class="iconbtn" data-gcal="next" aria-label="Next month">›</button></div>
        <div class="grid7">${["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].map(x => `<div class="dowh">${x}</div>`).join("")}${cells}</div></section>
      <h2 class="label" style="margin:4px 0 0;color:var(--muted)">${esc(longDate(sel))}</h2>${detail}
      <h2 class="label" style="margin:4px 0 0;color:var(--muted)">2027 Gatherings</h2>
      <p class="sub" style="font-size:13px;font-style:italic;margin:0">These are projected dates, pending the church calendar.</p>${upcoming}`;
  }
  // Profile: photos and preferences, saved to the couple's row in the Couples list.
  const photos = new Map();
  function photoUrl(raw) {
    const path = D.assetPath(raw, C.sitePath);
    const key = path || (raw === "me" ? "me" : "");
    if (!key) return Promise.resolve(null);
    if (!photos.has(key)) photos.set(key, (async () => {
      try {
        const res = key === "me" ? await graph("me/photo/$value", { raw: true })
          : await graph(`sites/${state.site.id}/lists/${encodeURIComponent("Site Assets")}/drive/root:/${path.split("/").map(encodeURIComponent).join("/")}:/content`, { raw: true });
        return URL.createObjectURL(await res.blob());
      } catch { return null; }
    })());
    return photos.get(key);
  }
  function hydratePhotos() {
    document.querySelectorAll("img[data-photo]").forEach(img => {
      const chain = img.dataset.photo.split("|").filter(Boolean);
      (async () => { for (const raw of chain) { const url = await photoUrl(raw); if (url) { img.src = url; img.hidden = false; img.nextElementSibling?.setAttribute("hidden", ""); return; } } })();
    });
  }
  const PENCIL = '<svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>';
  const initialsOf = name => String(name || "").trim().split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0].toUpperCase()).join("") || "?";
  const face = (raws, name, size) => `<span class="face" style="width:${size}px;height:${size}px;font-size:${Math.round(size / 2.8)}px"><img alt="" data-photo="${esc(raws)}" hidden><b>${esc(initialsOf(name))}</b></span>`;
  function myPhotoChain() { const c = profile(); return [c ? (mySide() === "spouse2" ? c.photo2 : c.photo1) : "", "me"].filter(Boolean).join("|"); }
  function myName() { const c = profile(); return (c ? (mySide() === "spouse2" ? c.spouse2Name : c.spouse1Name) : "") || state.account?.name || ""; }
  // Mentor couples are marked in each couple's AppData (isMentor); assignment uses the MentorCouple column (mentor emails).
  const allCouples = () => (state.rows.couples || []).map(item => ({ ...D.read("couples", item), app: D.parse(item), title: item.fields?.Title || "" }))
    .filter(c => c.kind === "couple" && c.active !== false);
  const isMentor = c => !!c?.app?.isMentor;
  const isLeader = () => (C.leaderIds || []).includes(String(oid() || ""));
  const leaderCouple = (list = allCouples()) => list.find(c => (c.app?.members || []).some(m => (C.leaderIds || []).includes(String(m.objectId))));
  const emailsOf = c => (c?.memberEmails || []).map(e => String(e).toLowerCase());
  const mentorOf = (c, list = allCouples()) => { if (isMentor(c)) { const lc = leaderCouple(list); return lc && lc.itemId !== c.itemId ? lc : null; } const want = String(c?.mentorCouple || "").toLowerCase().split(/[;,\s]+/).filter(Boolean);
    return want.length ? list.find(m => m.itemId !== c.itemId && emailsOf(m).some(e => want.includes(e))) || null : null; };
  const menteesOf = (m, list = allCouples()) => list.filter(c => c.itemId !== m.itemId && mentorOf(c, list)?.itemId === m.itemId);
  const pairNameOf = c => [c?.spouse1Name, c?.spouse2Name].map(n => String(n || "").trim().split(/\s+/)[0]).filter(Boolean).join(" & ") || c?.title || "Couple";
  const STAR = "";
  // A person's photo: their SharePoint photo, else their Microsoft photo when it's the signed-in person.
  const personChain = (raw, email) => [raw, String(email || "").toLowerCase() === String(state.account?.username || "").toLowerCase() ? "me" : ""].filter(Boolean).join("|");
  // Gives (or removes) a mentor couple's read access to the lists behind their dashboard numbers.
  // Uses the SharePoint permission already approved for this app registration (AllSites.FullControl), as RCI Productions does.
  async function spToken() {
    const request = { scopes: [`https://${C.siteHost}/AllSites.FullControl`], account: state.account };
    try { return (await auth.acquireTokenSilent(request)).accessToken; } catch { return (await auth.acquireTokenPopup(request)).accessToken; }
  }
  async function sp(path, opts = {}) {
    const url = new URL(`https://${C.siteHost}${C.sitePath}/_api/${path}`);
    if (url.hostname !== C.siteHost) throw new Error("Unexpected data destination");
    const res = await fetch(url, { ...opts, headers: { Authorization: "Bearer " + await spToken(), Accept: "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata", ...(opts.headers || {}) } });
    if (!res.ok) { const body = await res.json().catch(() => null); const e = new Error(body?.["odata.error"]?.message?.value || `SharePoint returned ${res.status}.`); e.status = res.status; throw e; }
    return res.status === 204 ? null : res.json().catch(() => null);
  }
  async function mentorAccess(c, on) {
    const READ = 1073741826;
    for (const email of emailsOf(c)) {
      const user = await sp("web/ensureuser", { method: "POST", body: JSON.stringify({ logonName: `i:0#.f|membership|${email}` }) });
      for (const key of ["participation", "rsvps", "mentorRequests"]) {
        const list = `web/lists/getbytitle('${C.lists[key].replace(/'/g, "''")}')`;
        const info = await sp(`${list}?$select=HasUniqueRoleAssignments`);
        if (!info?.HasUniqueRoleAssignments) continue; // already inherits site access
        if (on) await sp(`${list}/roleassignments/addroleassignment(principalid=${user.Id},roledefid=${READ})`, { method: "POST" });
        else await sp(`${list}/roleassignments/removeroleassignment(principalid=${user.Id},roledefid=${READ})`, { method: "POST" }).catch(() => null);
      }
    }
  }
  async function patchCouple(c, fields) {
    await graph(`${base("couples")}/items/${encodeURIComponent(c.itemId)}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
    await load("couples");
  }
  function leaderControls(c) {
    if (!state.leaderMode || !isLeader()) return "";
    const mentors = allCouples().filter(m => isMentor(m) && m.itemId !== c.itemId), cur = mentorOf(c);
    return `<section class="card" style="gap:10px">${label("Leader tools")}
      <button type="button" class="btn ghost" data-live="toggle-mentor" data-id="${esc(c.itemId)}" style="align-self:flex-start">${isMentor(c) ? "Mentor couple · Remove" : "Mark as mentor couple"}</button>
      ${isMentor(c) ? `<label for="mentor-branch">Branch name</label><div class="row" style="gap:8px"><input type="text" id="mentor-branch" maxlength="60" value="${esc(c.branch || "")}" placeholder="e.g. Olive" style="flex:1"><button type="button" class="btn small-btn" data-live="mentor-branch" data-id="${esc(c.itemId)}">Save</button></div>
        <p class="sub" style="font-size:13px">Saving also moves every couple assigned to them into this Branch.</p>` : `<label for="assign-mentor">Mentor couple</label><div class="row" style="gap:8px"><select id="assign-mentor" style="flex:1">
        <option value="">Not assigned</option>${mentors.map(m => `<option value="${esc(m.itemId)}"${cur?.itemId === m.itemId ? " selected" : ""}>${esc(pairNameOf(m))}${m.branch ? " · " + esc(m.branch) : ""}</option>`).join("")}</select>
        <button type="button" class="btn small-btn" data-live="assign-mentor" data-id="${esc(c.itemId)}">Save</button></div>
        ${mentors.length ? "" : `<p class="sub" style="font-size:13px">Mark a couple as a mentor couple first.</p>`}`}</section>`;
  }
  function coupleView() {
    const c = (state.rows.couples || []).map(item => D.read("couples", item)).find(x => x.itemId === state.coupleId);
    const from = state.coupleFrom || "branch";
    const back = `<button type="button" class="linkbtn" data-view="${from}" style="min-height:32px">‹ Back to ${from === "leaders" ? "Leader dashboard" : from === "mentor" ? "Mentor dashboard" : "Branch"}</button>`;
    if (!c) return back + card("Couple", empty("This couple isn’t available."));
    return back + `<section class="card profile-card">
        <div class="cover-wrap">${c.couplePhoto ? `<img alt="Couples photo" data-photo="${esc(c.couplePhoto)}" hidden class="cover"><div class="cover cover-empty">Loading photo…</div>` : `<div class="cover cover-empty">No couples photo yet</div>`}</div>
        <div class="pair">
          <div class="person">${face(personChain(c.photo1, c.memberEmails?.[0]), c.spouse1Name, 92)}<b>${esc(c.spouse1Name || "—")}</b></div>
          <div class="person">${face(personChain(c.photo2, c.memberEmails?.[1]), c.spouse2Name, 92)}<b>${esc(c.spouse2Name || "—")}</b></div>
        </div>
        ${(() => { const full = allCouples().find(x => x.itemId === c.itemId); const viewerMentor = full && mentorOf(full)?.itemId === myCouple()?.itemId;
          const icons = (isLeader() || viewerMentor) && (c.contact || c.gathering) ? `<div class="pref-icons">${c.contact ? `<span title="Preferred contact method">${c.contact === "Call" ? "📞" : c.contact === "Email" ? "✉️" : "💬"} ${esc(c.contact)}</span>` : ""}${c.gathering ? `<span title="Preferred gathering method">${c.gathering === "Virtual" ? "💻" : "🤝"} ${esc(c.gathering)}</span>` : ""}</div>` : "";
          return (isMentor(full) ? `<p style="text-align:center;margin:0;font-weight:700;color:var(--fig)">Mentor couple${c.branch ? " · " + esc(c.branch) : ""}</p>` : "") + icons; })()}
        <p class="sub" style="text-align:center;margin:0">${c.anniversary ? `💍 Anniversary: ${esc(longDate(c.anniversary).replace(/, \d{4}$/, ""))}` : "Anniversary not added yet"}</p>
      </section>` + leaderControls(allCouples().find(x => x.itemId === c.itemId) || c);
  }
  function profileView() {
    const c = profile();
    if (state.errors.couples) return card("Profile", empty(state.errors.couples));
    if (!c) return card("Profile", empty("Your couple profile hasn’t been set up yet. Ask a ministry administrator to add you to the Couples list."));
    const two = mySide() === "spouse2", sp = spouseOf(), mine = two ? "Spouse2Photo" : "Spouse1Photo";
    const myPhone = two ? c.phone2 : c.phone1, myBday = two ? c.birthday2 : c.birthday1;
    const pills = (name, value, opts) => `<div class="pills">${opts.map(o => `<label class="pill-check"><input type="radio" name="${name}" value="${o}" ${value === o ? "checked" : ""} ${state.profileEdit ? "" : "disabled"}><span>${o}</span></label>`).join("")}</div>`;
    const cover = c.couplePhoto ? `<img alt="Your couples photo" data-photo="${esc(c.couplePhoto)}" hidden class="cover"><div class="cover cover-empty">Loading photo…</div>`
      : `<div class="cover cover-empty">Add a couples photo</div>`;
    return `<section class="card profile-card">
        <div class="cover-wrap">${cover}<label class="edit-btn cover-edit" for="pf-couple" aria-label="${c.couplePhoto ? "Change couples photo" : "Add couples photo"}">${PENCIL}</label><input id="pf-couple" type="file" accept="image/*" data-field="CouplePhoto" hidden></div>
        <div class="pair">
          <div class="person"><span class="face-wrap">${face(myPhotoChain(), myName(), 92)}<label class="edit-btn face-edit" for="pf-photo" aria-label="Change your photo">${PENCIL}</label></span><b>${esc(myName())}</b><span class="sub">You</span>
            <input id="pf-photo" type="file" accept="image/*" data-field="${mine}" hidden></div>
          <div class="person">${face(sp?.photo || "", sp?.name, 92)}<b>${esc(sp?.name || "Not added yet")}</b><span class="sub">Your spouse</span></div>
        </div>
        ${c.anniversary ? `<p class="sub" style="text-align:center;margin:0">💍 Married ${esc(longDate(c.anniversary))}</p>` : ""}
      </section>
      <form data-form="profile" class="card" style="gap:12px;position:relative">${label("About me")}
        <span class="form-icons"><button class="icon-round${state.profileEdit ? " on" : ""}" type="button" data-live="profile-edit" aria-pressed="${state.profileEdit ? "true" : "false"}" aria-label="${state.profileEdit ? "Stop editing" : "Edit profile"}" title="Edit">${PENCIL}</button><button class="icon-round" type="submit" aria-label="Save profile" title="Save" ${state.profileEdit ? "" : "disabled"}><svg aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/></svg></button></span>
        <span style="font-size:16px;word-break:break-all;color:var(--muted)">${esc(state.account?.username || "")}</span>
        <label for="pf-phone">Phone number</label><input id="pf-phone" type="tel" name="phone" maxlength="40" ${state.profileEdit ? "" : "disabled"} value="${esc(myPhone || "")}" placeholder="(555) 555-5555" autocomplete="tel">
        <label for="pf-bday">Birthday</label><input id="pf-bday" type="date" name="birthday" ${state.profileEdit ? "" : "disabled"} value="${esc(myBday || "")}">
        ${label("About us", ";margin-top:6px")}
        <label for="pf-anniv">Anniversary date</label><input id="pf-anniv" type="date" name="anniversary" ${state.profileEdit ? "" : "disabled"} value="${esc(c.anniversary || "")}">
        <span style="font-weight:600">Preferred contact method</span>${pills("contact", c.contact, ["Text", "Call", "Email"])}
        <span style="font-weight:600">Preferred gathering method</span>${pills("gathering", c.gathering, ["In person", "Virtual"])}
        <p class="form-status sub" role="status" style="margin:0"></p></form>
      <button class="btn ghost" type="button" data-live="signout">Sign out</button>`;
  }
  const HEIC_MESSAGE = "HEIC photos can’t be uploaded. Please choose a JPG or PNG. On iPhone, go to Settings → Camera → Formats and pick “Most Compatible,” or take a screenshot of the photo and upload that.";
  function photoError(text) { const error = new Error(text); error.userMessage = text; return error; }
  function isHeic(file) {
    const type = String(file?.type || "").toLowerCase(), name = String(file?.name || "").toLowerCase();
    return /^image\/hei[cf](-sequence)?$/.test(type) || /\.hei[cf]$/.test(name);
  }
  async function resizeImage(file, max = 800) {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(photoError("We couldn’t read that photo. Please choose a JPG or PNG photo and try again.")); i.src = URL.createObjectURL(file); });
    const k = Math.min(1, max / Math.max(img.width, img.height));
    const cv = document.createElement("canvas"); cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
    cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
    return new Promise(r => cv.toBlob(r, "image/jpeg", 0.86));
  }
  async function updateCouple(fields) {
    const c = profile(); if (!c?.itemId) throw new Error("No couple profile");
    await graph(`${base("couples")}/items/${encodeURIComponent(c.itemId)}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
    await load("couples");
  }
  async function uploadPhoto(field, file) {
    const c = profile(); if (!c?.itemId || !["Spouse1Photo", "Spouse2Photo", "CouplePhoto"].includes(field)) throw new Error("Invalid photo");
    if (field !== "CouplePhoto" && field !== (mySide() === "spouse2" ? "Spouse2Photo" : "Spouse1Photo")) throw new Error("Invalid photo");
    if (isHeic(file)) throw photoError(HEIC_MESSAGE);
    const blob = await resizeImage(file, field === "CouplePhoto" ? 1400 : 800);
    const fileName = `couple_${c.itemId}_${field}_${Date.now()}.jpg`;
    const up = await graph(`sites/${state.site.id}/lists/${encodeURIComponent("Site Assets")}/drive/root:/Lists/${encodeURIComponent(state.lists.couples)}/${fileName}:/content`,
      { method: "PUT", headers: { "Content-Type": "image/jpeg" }, body: blob });
    const u = new URL(up.webUrl);
    await updateCouple({ [field]: JSON.stringify({ type: "thumbnail", fileName, serverUrl: u.origin, serverRelativeUrl: decodeURIComponent(u.pathname) }) });
  }
  // Signed-in photo next to the light/dark button; opens the profile.
  function headerPhoto() {
    const bar = document.querySelector("#brandbar > div"); if (!bar) return;
    let slot = document.getElementById("me-slot");
    const toggle = document.getElementById("theme-toggle");
    if (!slot && toggle) { slot = document.createElement("span"); slot.id = "me-slot"; slot.style.cssText = "display:flex;align-items:center;gap:8px"; toggle.replaceWith(slot); slot.append(toggle); }
    if (!slot) return;
    slot.querySelector(".me-btn")?.remove();
    if (!state.account) return;
    const btn = document.createElement("button"); btn.type = "button"; btn.className = "me-btn"; btn.dataset.view = "profile"; btn.setAttribute("aria-label", "My profile");
    btn.innerHTML = face(myPhotoChain(), myName(), 44); slot.append(btn);
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
  function coupleOnboarding() {
    if (!state.lists.onboarding || state.errors.onboarding) return "";
    const requests = own("onboarding").map(item => ({ ...D.parse(item), id: Number(item.id) }))
      .filter(request => request.kind === "couple-onboarding").sort((a, b) => b.id - a.id).slice(0, 6);
    const statusLabel = { queued: "Queued", processing: "Setting up access", ready: "Ready", blocked: "Needs administrator review" };
    const status = requests.length ? panel("Recent onboarding requests", requests.map(request =>
      `<div class="row"><strong>${esc(request.coupleName)}</strong><span class="pill">${esc(statusLabel[request.status] || "Queued")}</span></div>`).join("")
      + button("couple-onboard-refresh", "Refresh onboarding status")) : "";
    if (!state.addCouple) return status;
    const person = (key, name) => `<fieldset style="border:0;margin:0;padding:0;display:grid;gap:7px"><legend>${name}</legend>
      <label for="new-${key}-name">Full name</label><input id="new-${key}-name" name="${key}Name" required maxlength="120" autocomplete="off">
      <label for="new-${key}-email">Microsoft sign-in email</label><input id="new-${key}-email" name="${key}Email" type="email" required maxlength="200" autocomplete="off"></fieldset>`;
    return `<form class="card" data-form="couple-onboarding" style="gap:12px"><h2 style="margin:0">Add a couple</h2>
      <p class="sub">Use the two existing Microsoft accounts they’ll sign in with. The couple name will be used for their private journal folder.</p>
      <label for="new-couple-name">Couple name</label><input id="new-couple-name" name="coupleName" required maxlength="120" autocomplete="off">
      ${person("husband", "Husband")}${person("wife", "Wife")}
      <div class="row"><button class="btn" type="submit">Save couple</button>${button("couple-onboard-close", "Cancel")}</div><p role="status" class="form-status sub"></p></form>` + status;
  }
  async function queueCouple(values) {
    const clean = key => String(values.get(key) || "").trim();
    const coupleName = clean("coupleName");
    const spouses = ["husband", "wife"].map(key => ({ name: clean(key + "Name"), email: clean(key + "Email").toLowerCase() }));
    if (!coupleName || coupleName.length > 120 || /[\\/:*?"<>|#%\u0000-\u001f]/.test(coupleName) || /[. ]$/.test(coupleName))
      throw new Error("Enter a couple name without folder punctuation or a trailing period.");
    if (spouses.some(spouse => !spouse.name || spouse.name.length > 120 || spouse.email.length > 200 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(spouse.email)) || spouses[0].email === spouses[1].email)
      throw new Error("Enter a full name and a different Microsoft sign-in email for each spouse.");
    const roster = (state.rows.couples || []).flatMap(item => D.read("couples", item).memberEmails || []).map(email => String(email).toLowerCase());
    if (spouses.some(spouse => roster.includes(spouse.email))) throw new Error("One of these accounts is already on the marriage roster.");
    const signature = new TextEncoder().encode(spouses.map(spouse => spouse.email).sort().join("\n"));
    const bytes = new Uint8Array(await crypto.subtle.digest("SHA-256", signature));
    const Title = "onboard:" + Array.from(bytes, byte => byte.toString(16).padStart(2, "0")).join("");
    if (own("onboarding").some(item => item.fields?.Title === Title)) throw new Error("This couple already has an onboarding request.");
    const fields = { Title, AppData: JSON.stringify({ kind: "couple-onboarding", version: 1, coupleName, spouses, status: "queued" }) };
    await graph(`${base("onboarding")}/items`, { method: "POST", body: JSON.stringify({ fields }) });
    await load("onboarding");
  }
  const leaders = () => isLeader() ? dashboard("all") : card("Leader dashboard", empty("The Leader dashboard is only for the ministry leaders."));
  const mentorDash = () => dashboard("mentor");
  const myCouple = () => allCouples().find(c => c.itemId === profile()?.itemId) || null;
  const myBranchCouples = () => { const me = myCouple(); if (!me) return [];
    const ids = new Set(menteesOf(me).map(c => c.itemId));
    return allCouples().filter(c => c.itemId !== me.itemId && (ids.has(c.itemId) || (me.branch && c.branch === me.branch && !isMentor(c)))); };
  const circleOf = c => `<button type="button" class="av-btn${isMentor(c) ? " mentor" : ""}" data-couple="${esc(c.itemId)}">${isMentor(c) ? STAR : ""}<span><img alt="" data-photo="${esc(c.couplePhoto || "")}" hidden><b>${esc([c.spouse1Name, c.spouse2Name].map(n => String(n || "?").trim()[0] || "?").join("&"))}</b></span>${esc(pairNameOf(c))}</button>`;
  function dashboard(scope) {
    const mentorView = scope === "mentor";
    if (state.errors.couples) return card("Dashboard", empty(state.errors.couples));
    if (mentorView && !isMentor(myCouple()) && !isLeader()) return card("Mentor dashboard", empty("This dashboard is for mentor couples."));
    const everyone = allCouples();
    const roster = (mentorView ? myBranchCouples() : everyone).map(c => ({ ...c, id: String(c.itemId) }));
    const parts = (state.rows.participation || []).map(D.parse).filter(r => r.kind === "participation");
    const ids = new Set(roster.map(c => c.id));
    const month = day().slice(0, 7), monthName = FRUIT[Number(month.slice(5)) - 1];
    const week = parts.filter(r => r.week === state.week);
    const did = (rows, c, flag) => rows.some(r => String(r.coupleId) === c.id && (!flag || r[flag]));
    const tookPart = roster.filter(c => parts.some(r => String(r.coupleId) === c.id && String(r.week || "").slice(0, 7) === month && (r.answered || r.checkedIn))).length;
    const first = n => String(n || "").trim().split(/\s+/)[0];
    const pairName = c => [first(c.spouse1Name), first(c.spouse2Name)].filter(Boolean).join(" & ") || c.title || "Couple";
    const mentorName = c => { const emails = String(c.mentorCouple || "").toLowerCase().split(/[;,\s]+/).filter(Boolean);
      const m = roster.find(x => (x.memberEmails || []).some(e => emails.includes(String(e).toLowerCase()))); return m ? pairName(m) : (emails.length ? "Assigned" : "—"); };
    const rsvps = (state.rows.rsvps || []).map(D.parse).filter(r => r.kind === "rsvp" && r.response === "yes");
    const going = ev => new Set(rsvps.filter(r => r.event === ev && (!mentorView || ids.has(String(r.coupleId)))).map(r => String(r.coupleId))).size;
    const keys = new Set(roster.map(c => String(c.coupleKey)));
    const requests = (state.rows.mentorRequests || []).map(item => ({ id: item.id, f: item.fields || {}, at: item.createdDateTime || "" }))
      .filter(r => !mentorView || keys.has(String(r.f.CoupleKey || ""))).sort((a, b) => b.at.localeCompare(a.at));
    const open = requests.filter(r => (r.f.Status || "New") !== "Contacted");
    const branches = [...new Set(roster.map(c => c.branch || "Unassigned"))].sort();
    const status = pct => pct >= .8 ? ["Healthy", "good"] : pct >= .4 ? ["Watch", "watch"] : ["Reach out", "reach"];
    const bar = (n, d) => `<span class="meter"><i style="width:${d ? Math.round(n / d * 100) : 0}%"></i></span>`;
    const tile = (lbl, big, small, hot) => `<div class="tile${hot ? " hot" : ""}"><span class="sub">${lbl}</span><b>${big}</b><span class="sub" style="font-size:13px">${small}</span></div>`;
    const tiles = `<div class="tiles">${mentorView ? tile("Couples in your Branch", roster.length, esc(myCouple()?.branch || "Your Branch")) : tile("Couples enrolled", roster.length, `${branches.length} Branch${branches.length === 1 ? "" : "es"}`)}
      ${tile(`Took part in ${monthName[0]}`, `${tookPart} of ${roster.length}`, "Answered or checked in")}
      ${tile("Retreat interest", going("retreat"), "couples · RSVP’d yes")}
      ${tile("Open mentor requests", open.length, "Reach out within a week", open.length > 0)}</div>`;
    const perCouple = `<section class="card dash"><h2>Couple health</h2><div class="scroll"><table class="health"><thead><tr><th>Couple</th><th>Sunday <span class="br-sm">Check In</span></th><th>Reveal <span class="br-sm">Together</span></th><th>Status</th></tr></thead><tbody>${roster.map(c => {
        const ci = did(week, c, "checkedIn"), an = did(week, c, "answered"), [st, cls] = status(ci || an ? (ci && an ? 1 : .5) : 0);
        return `<tr><td><b>${esc(pairNameOf(c))}</b></td><td>${ci ? "✓" : "—"}</td><td>${an ? "✓" : "—"}</td><td><span class="pill st-${cls}">${st}</span></td></tr>`;
      }).join("") || `<tr><td colspan="4">No couples are assigned to you yet.</td></tr>`}</tbody></table></div>
      <label for="leader-week" class="sub" style="margin-top:6px">Week containing</label><input id="leader-week" type="date" value="${state.week}" style="max-width:220px"></section>`;
    const health = `<section class="card dash"><h2>Branch health</h2><div class="scroll"><table class="health"><thead><tr><th>Branch</th><th class="hide-sm">Mentor couple</th><th>Sunday <span class="br-sm">Check Ins</span></th><th>Reveal <span class="br-sm">Together</span></th><th>Status</th></tr></thead><tbody>${branches.map(b => {
        const cs = roster.filter(c => (c.branch || "Unassigned") === b), n = cs.length;
        const ci = cs.filter(c => did(week, c, "checkedIn")).length, an = cs.filter(c => did(week, c, "answered")).length;
        const [st, cls] = status(n ? Math.max(ci, an) / n : 0);
        return `<tr><td><b>${esc(b)}</b></td><td class="hide-sm">${esc(mentorName(cs[0] || {}))}</td><td><span class="meter-row">${bar(ci, n)}${ci} of ${n}</span></td><td>${an} of ${n}</td><td><span class="pill st-${cls}">${st}</span></td></tr>`;
      }).join("") || `<tr><td colspan="5">No active couples are listed.</td></tr>`}</tbody></table></div>
      <label for="leader-week" class="sub" style="margin-top:6px">Week containing</label><input id="leader-week" type="date" value="${state.week}" style="max-width:220px"><p class="sub" style="font-size:13px">Week of ${esc(longDate(state.week))} · Monday through Sunday, Eastern time</p></section>`;
    const coupleOf = key => roster.find(c => c.coupleKey === key);
    const mentor = `<section class="card dash"><h2>Mentor requests</h2>${state.errors.mentorRequests ? empty(state.errors.mentorRequests) : requests.slice(0, 8).map(r => {
        const c = coupleOf(r.f.CoupleKey), when = r.at ? new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", timeZone: "America/New_York" }).format(new Date(r.at)) : "";
        const done = (r.f.Status || "New") === "Contacted";
        return `<div class="req"><div><b>${esc(c ? pairName(c) : r.f.CoupleKey || "Couple")}</b><span class="sub">${esc(r.f.Branch || c?.branch || "")}${when ? " · asked " + esc(when) : ""}</span></div>
          <button type="button" class="btn ${done ? "" : "ghost"} small-btn" data-live="mentor-contacted" data-id="${esc(r.id)}" ${done ? "disabled" : ""}>${done ? "Contacted ✓" : "Mark contacted"}</button></div>`;
      }).join("") || empty("No mentor requests yet.")}</section>`;
    const events = [["harvest", "Harvest Celebration"], ["retreat", "Marriage Retreat"]];
    const signups = `<section class="card dash"><h2>Sign-ups</h2>${state.errors.rsvps ? empty(state.errors.rsvps) : events.map(([k, name]) => `<div class="signup"><div class="row"><b>${name}</b><span>${going(k)} going</span></div>${bar(going(k), roster.length)}</div>`).join("")}</section>`;
    const content = (state.rows.content || []).map(item => D.read("content", item));
    const cols = [["video", "Leader Video"], ["devotional", "Devotional"]];
    const curM = Number(month.slice(5));
    const cell = (m, kind, y = 2027) => {
      const ym = `${y}-${String(m).padStart(2, "0")}`;
      const water = kind === "challenge" && monthFocus(ym)?.water;
      const uploaded = kind === "video" && state.leaderVideos?.[ym];
      const has = uploaded || water || content.some(r => r.kind === kind && String(r.date || "").startsWith(ym));
      const live = has && day() >= ym + (water ? "-08" : "-01");
      if (kind === "video" && !has && !state.videoIssue) {
        return `<button class="pill cs-none" type="button" data-live="upload-leader-video" data-month="${ym}" aria-label="Add ${FRUIT[m - 1][0]} leader video" style="border:0;cursor:pointer">Add video</button>`;
      }
      return `<span class="pill cs-${live ? "live" : has ? "ready" : "none"}">${live ? "Live" : has ? "Ready" : "Not started"}</span>`;
    };
    const schedule = `<section class="card dash"><div class="row" style="flex-wrap:wrap"><h2>Content schedule</h2><span style="display:flex;gap:6px;flex-wrap:wrap"><span class="pill cs-live">Live</span><span class="pill cs-ready">Ready</span><span class="pill cs-none">Not started</span></span></div>
      <div class="scroll"><table><thead><tr><th>Month</th><th>Fruit</th>${cols.map(c => `<th>${c[1]}</th>`).join("")}</tr></thead><tbody>${[[11, 2026], [12, 2026]].map(([m, y]) => { const f = FRUIT[m - 1], ym = `${y}-${m}`;
          return `<tr${day().startsWith(ym) ? ' class="cur"' : ""}><td><b>${f[0].slice(0, 3)} ${y}</b></td><td>${esc(monthFocus(ym)?.focus || f[1])}</td>${cols.map(c => `<td>${cell(m, c[0], y)}</td>`).join("")}</tr>`; }).join("")}${FRUIT.map((f, i) =>
        `<tr${day().startsWith("2027-") && i + 1 === curM ? ' class="cur"' : ""}><td><b>${f[0].slice(0, 3)} 2027</b></td><td>${esc(monthFocus(`2027-${String(i + 1).padStart(2, "0")}`)?.focus || f[1])}</td>${cols.map(c => `<td>${cell(i + 1, c[0])}</td>`).join("")}</tr>`).join("")}</tbody></table></div>
      <p class="sub" style="font-size:13px">${state.videoIssue ? "Leader video folder: " + esc(state.videoIssue) : "Upload each month’s MP4 to the ministry’s leader video Dropbox folder, copy that video’s share link, then tap Add video for the month and paste it."}</p></section>`;
    if (mentorView) return `<section class="card dash"><h2>Your Branch couples</h2>${roster.length ? `<div class="avatars big">${roster.map(circleOf).join("")}</div>` : empty("Couples assigned to you will appear here.")}</section>
      ${state.errors.participation ? card("Participation", empty(state.errors.participation)) : ""}${tiles}
      <div class="cols"><div class="main">${perCouple}</div><div class="side">${mentor}${signups}</div></div>`;
    const mentors = everyone.filter(isMentor), unassigned = everyone.filter(c => !isMentor(c) && !mentorOf(c, everyone));
    const mentorCard = `<section class="card dash"><h2>Mentor couples · ${mentors.length}</h2>${mentors.map(m => `<div class="req"><div><b>${esc(pairNameOf(m))}</b><span class="sub">${esc(m.branch || "No Branch")} · ${menteesOf(m, everyone).length} couple${menteesOf(m, everyone).length === 1 ? "" : "s"}</span></div><button type="button" class="btn ghost small-btn" data-couple="${esc(m.itemId)}">Open</button></div>`).join("") || empty("No mentor couples yet. Open a couple below and tap “Mark as mentor couple.”")}</section>`;
    const unassignedCard = `<section class="card dash"><h2>Not assigned to a mentor · ${unassigned.length}</h2>${unassigned.length ? `<div class="avatars big">${unassigned.map(circleOf).join("")}</div><p class="sub" style="font-size:13px">Tap a couple to assign them to a mentor couple.</p>` : empty("Every couple has a mentor couple.")}</section>`;
    const allCard = `<section class="card dash"><h2>All couples · ${everyone.length}</h2><div class="avatars big">${everyone.map(circleOf).join("")}</div><p class="sub" style="font-size:13px">Tap a couple to mark them as a mentor couple or change their mentor.</p></section>`;
    return `${coupleOnboarding()}${allCard}${unassignedCard}${mentorCard}${state.errors.participation ? card("Participation", empty(state.errors.participation)) : ""}${tiles}
      <div class="cols"><div class="main">${health}</div><div class="side">${mentor}${signups}</div></div>${schedule}`;
  }
  function render() {
    const account = state.account, wide = ["leaders", "mentor"].includes(state.view) && !!account;
    root.className = "app" + (wide ? " wide" : "");
    const bar = document.getElementById("brandbar"); if (bar && bar !== root) bar.className = "brandbar" + (wide ? " wide" : "");
    const views = { couple: coupleView, profile: profileView, calendar, today, checkin, journal, orchard: pilotJournal, sealed, grow, branch, gather, leaders, mentor: mentorDash };
    root.innerHTML = header() + `<main>
      ${!account ? card("Sign in", `<p class="sub">Sign in with your provided email address. A “Sign in with Microsoft” screen is to be expected.</p>${C.clientId ? button("signin", "Sign in with Microsoft") : empty("Live sign-in is being configured.")}${state.message ? `<p role="status">${esc(state.message)}</p>` : ""}`) :
        `${state.message ? `<div class="banner" role="status">${esc(state.message)}</div>` : ""}${state.site && !state.loading && state.view === "today" ? notifications() : ""}${state.loading ? "" : !state.site ? empty("The connection is not ready. Pull down to refresh.") : (views[state.view] || today)()}`}
      </main>${account && !wide ? tabBar() : ""}`;
    if (typeof document.querySelector === "function") { headerPhoto(); hydratePhotos(); document.body?.classList.toggle("signed-out", !account); }
    if (account) { try { localStorage.setItem("rci-marriage-view", JSON.stringify({ view: state.view, coupleId: state.coupleId || null })); } catch {} }
  }
  window.renderLive = render;
  // Pull down at the top of any page to refresh (same as RCI Security).
  if (typeof document.createElement === "function" && typeof addEventListener === "function") (function () {
    const ind = document.createElement("div"); ind.id = "ptr"; ind.innerHTML = "<span>↓ Pull to refresh</span>";
    if (document.body) document.body.appendChild(ind); else document.addEventListener("DOMContentLoaded", () => document.body.appendChild(ind));
    let y0 = null, h = 0;
    const editing = t => t && t.closest && t.closest("input,textarea,select,[contenteditable='true'],.grid7,.cal");
    addEventListener("touchstart", e => { y0 = (window.scrollY <= 0 && e.touches.length === 1 && !editing(e.target) && !editing(document.activeElement)) ? e.touches[0].clientY : null; h = 0; }, { passive: true });
    addEventListener("touchmove", e => {
      if (y0 == null) return;
      const dy = e.touches[0].clientY - y0;
      if (dy <= 0 || window.scrollY > 0) { ind.style.height = "0"; h = 0; return; }
      h = Math.min(dy * 0.5, 90); ind.style.height = h + "px";
      const ready = h >= 60; ind.classList.toggle("ready", ready); ind.firstChild.textContent = ready ? "↻ Release to refresh" : "↓ Pull to refresh";
    }, { passive: true });
    addEventListener("touchend", async () => {
      if (y0 == null) return; y0 = null;
      if (h >= 60 && state.account && !state.loading) { ind.firstChild.textContent = "Refreshing…"; ind.style.height = "48px"; photos.clear(); await loadView(); }
      ind.style.height = "0"; h = 0;
    });
  })();
  document.addEventListener("click", async event => {
    const pick = event.target.closest("[data-couple]");
    if (pick) { state.coupleFrom = ["leaders", "mentor"].includes(state.view) ? state.view : "branch"; state.leaderMode = state.view === "leaders"; state.coupleId = pick.dataset.couple; state.view = "couple"; render(); window.scrollTo?.(0, 0); return; }
    const gday = event.target.closest("[data-gday]");
    if (gday) { state.gcal = { month: gday.dataset.gday.slice(0, 7), sel: gday.dataset.gday }; render(); return; }
    const gnav = event.target.closest("[data-gcal]");
    if (gnav && state.gcal) {
      const [y, m] = state.gcal.month.split("-").map(Number), t = day();
      const to = gnav.dataset.gcal === "today" ? t.slice(0, 7) : new Date(Date.UTC(y, m - 1 + (gnav.dataset.gcal === "next" ? 1 : -1), 1)).toISOString().slice(0, 7);
      state.gcal = { month: to, sel: gnav.dataset.gcal === "today" ? t : to + "-01" }; render(); return;
    }
    const calDay = event.target.closest("[data-cal]");
    if (calDay) { state.calDay = Number(calDay.dataset.cal); render(); return; }
    const view = event.target.closest("[data-view]");
    if (view && !state.loading) { if (view.dataset.view === "gather") state.gcal = null; state.view = view.dataset.view; state.message = ""; await loadView(); return; }
    const target = event.target.closest("[data-live]"); if (!target) return;
    if (target.dataset.live === "watch-leader-video") { await startLeaderVideo(target); return; }
    if (target.dataset.live === "upload-leader-video" && state.view === "leaders" && !state.videoUploading) {
      const link = typeof prompt === "function" ? prompt(`Paste the Dropbox share link for the ${target.dataset.month} leader video`) : "";
      if (!link) return;
      state.videoUploading = true; state.message = "Saving the video link…"; render();
      try { await addLeaderVideo(target.dataset.month, link); state.message = "Leader video added."; }
      catch (error) { state.message = error.status ? "Couldn’t save the video link. " + notice(error) : error.message; }
      finally { state.videoUploading = false; render(); }
      return;
    }
    try {
      if (target.dataset.live === "signin") await auth.loginRedirect({ scopes, redirectUri: C.redirectUri });
      if (target.dataset.live === "refresh" && !state.loading) await loadView();
      if (target.dataset.live === "push-check") await checkPush();
      if (target.dataset.live === "push-on") await changePush(true);
      if (target.dataset.live === "couple-onboard-open" && state.view === "leaders" && state.lists.onboarding && !state.errors.onboarding) {
        state.addCouple = true; render(); document.getElementById("new-couple-name")?.focus();
      }
      if (target.dataset.live === "couple-onboard-close") { state.addCouple = false; render(); }
      if (target.dataset.live === "couple-onboard-refresh" && state.view === "leaders") { await load("onboarding"); await load("couples"); render(); }
      if (target.dataset.live === "profile-edit") { state.profileEdit = !state.profileEdit; render(); }
      if (target.dataset.live === "toggle-mentor" && state.leaderMode && isLeader()) {
        const c = allCouples().find(x => x.itemId === target.dataset.id); if (!c) return;
        target.disabled = true;
        const on = !isMentor(c);
        await patchCouple(c, { AppData: JSON.stringify({ ...c.app, isMentor: on }) });
        try { await mentorAccess(c, on); state.message = on ? "Marked as a mentor couple. They can now see their Branch numbers." : "Mentor couple removed, and their dashboard access was taken away."; }
        catch (error) { state.message = (on ? "Marked as a mentor couple, but their list access couldn’t be set: " : "Removed, but their list access couldn’t be changed: ") + (error.message || notice(error)); }
        render();
      }
      if (target.dataset.live === "mentor-branch" && state.leaderMode && isLeader()) {
        const c = allCouples().find(x => x.itemId === target.dataset.id), name = String(document.getElementById("mentor-branch")?.value || "").trim().slice(0, 60);
        if (!c || !name) return;
        target.disabled = true;
        for (const m of [c, ...menteesOf(c)]) await graph(`${base("couples")}/items/${encodeURIComponent(m.itemId)}/fields`, { method: "PATCH", body: JSON.stringify({ Branch: name }) });
        await load("couples"); state.message = `Branch set to ${name}.`; render();
      }
      if (target.dataset.live === "assign-mentor" && state.leaderMode && isLeader()) {
        const c = allCouples().find(x => x.itemId === target.dataset.id), sel = document.getElementById("assign-mentor");
        const m = allCouples().find(x => x.itemId === sel?.value);
        if (!c) return;
        target.disabled = true;
        await patchCouple(c, m ? { MentorCouple: emailsOf(m).join(";"), ...(m.branch ? { Branch: m.branch } : {}) } : { MentorCouple: null });
        state.message = m ? `Assigned to ${pairNameOf(m)}.` : "Mentor assignment removed."; render();
      }
      if (target.dataset.live === "react") {
        const b = String(target.dataset.branch || ""), kind = target.dataset.kind, id = String(target.dataset.id || "").replace(/\D/g, "");
        if (!["heart", "up", "down"].includes(kind) || !id || (!isLeader() && b !== profile()?.branch)) return;
        target.disabled = true;
        if (target.dataset.mine) await graph(`${base("messages")}/items/${encodeURIComponent(target.dataset.mine)}`, { method: "DELETE" });
        else await graph(`${base("messages")}/items`, { method: "POST", body: JSON.stringify({ fields: { Title: `react:${id}:${kind}`, Branch: b, Message: "", AuthorName: myName() } }) });
        await load("messages"); render();
      }
      if (target.dataset.live === "delete-message") {
        target.disabled = true;
        await graph(`${base("messages")}/items/${encodeURIComponent(target.dataset.id)}`, { method: "DELETE" });
        await load("messages"); render();
      }
      if (target.dataset.live === "mentor-contacted") {
        target.disabled = true;
        await graph(`${base("mentorRequests")}/items/${encodeURIComponent(target.dataset.id)}/fields`, { method: "PATCH", body: JSON.stringify({ Status: "Contacted" }) });
        await load("mentorRequests"); render();
      }
      if (target.dataset.live === "mentor-request") {
        const c = profile(); if (!c) throw new Error("No couple profile");
        target.disabled = true;
        await graph(`${base("mentorRequests")}/items`, { method: "POST", body: JSON.stringify({ fields: { Title: "Mentor request", CoupleKey: String(c.coupleKey || ""), Branch: String(c.branch || ""), Status: "New" } }) });
        state.mentorSent = true; render();
      }
      if (target.dataset.live === "push-off") await changePush(false);
      if (target.dataset.live === "signout") {
        if (push.busy) return;
        await changePush(false);
        controller.abort(); controller = new AbortController(); const account = state.account;
        state.account = null; state.rows = {}; state.journal = null; state.shared = []; state.site = null; render();
        await auth.logoutRedirect({ account, postLogoutRedirectUri: C.redirectUri });
      }
    } catch (error) { state.message = notice(error); render(); }
  });
  document.addEventListener("change", event => {
    if (event.target.id === "grow-month") {
      state.growMonth = event.target.value;
      render();
      return;
    }
    if (event.target.id === "branch-view" && isLeader()) { state.branchView = event.target.value; render(); return; }
    if (event.target.id !== "leader-week" || !/^\d{4}-\d{2}-\d{2}$/.test(event.target.value)) return;
    state.week = D.weekStart(event.target.value);
    render();
  });
  document.addEventListener("change", async event => {
    const input = event.target;
    if (!["pf-photo", "pf-couple"].includes(input.id) || !input.files?.[0]) return;
    state.message = "Uploading photo…"; render();
    try { await uploadPhoto(input.dataset.field, input.files[0]); photos.clear(); state.message = "Photo updated."; }
    catch (error) { state.message = error.userMessage || "Couldn’t update the photo. " + notice(error); }
    input.value = "";
    render();
  });
  document.addEventListener("submit", async event => {
    const form = event.target.closest("[data-form]"); if (!form) return;
    event.preventDefault(); const submit = form.querySelector('[type="submit"]');
    if (submit.disabled) return; submit.disabled = true;
    const status = form.querySelector(".form-status"), values = new FormData(form);
    if (form.dataset.form === "couple-onboarding") {
      if (state.view !== "leaders" || !state.lists.onboarding || state.errors.onboarding) return;
      status.textContent = "Queueing couple setup…";
      try {
        await queueCouple(values);
        state.addCouple = false;
        state.message = "Couple setup queued. Check the onboarding status before asking them to sign in.";
        render();
      } catch (error) {
        status.textContent = error.status ? notice(error) : error.message;
        submit.disabled = false;
      }
      return;
    }
    status.textContent = "Saving to SharePoint…";
    try {
      if (form.dataset.form === "answer") await save("answers", `prompt:${values.get("promptId")}:${day()}`, { version: 1, kind: "prompt-answer", promptId: values.get("promptId"), date: day(), answer: values.get("answer") });
      if (form.dataset.form === "checkin") {
        const cp = checkinPrompts(), ratings = [0,1,2,3].map(i=>Number(values.get(`rating${i}`)));
        const answers = cp.qs.map((q, i) => ({ q, a: String(values.get(`q${i}`) || "").trim() }));
        const summary = [...cp.rates.map((r, i) => `${r}: ${ratings[i] || "-"}/5`), "", ...answers.flatMap(x => [x.q, x.a || "-", ""])].join("\n").trim();
        await save("checkins", `check-in:${day()}`, { version: 1, kind: "check-in", date: day(), ratings, prompts: cp.rates, answers, summary,
          need: answers[0]?.a || "", saw: answers[answers.length - 1]?.a || "", mentor: values.get("mentor") === "yes" });
      }
      if (form.dataset.form === "branch-message" || form.dataset.form === "branch-reply") {
        const c = profile(), text = String(values.get("text") || "").trim().slice(0, 2000), b = String(values.get("branch") || "");
        if (!text || !b || (!isLeader() && b !== c?.branch)) throw new Error("Write a message first.");
        const parent = form.dataset.form === "branch-reply" ? String(values.get("parent") || "").replace(/\D/g, "") : "";
        await graph(`${base("messages")}/items`, { method: "POST", body: JSON.stringify({ fields: { Title: parent ? "reply:" + parent : b, Branch: b, Message: text, AuthorName: myName() } }) });
        await load("messages");
      }
      if (form.dataset.form === "profile") {
        const anniv = String(values.get("anniversary") || ""), contact = String(values.get("contact") || ""), gathering = String(values.get("gathering") || "");
        if (anniv && !/^\d{4}-\d{2}-\d{2}$/.test(anniv)) throw new Error("Invalid date");
        const bday = String(values.get("birthday") || ""), phone = String(values.get("phone") || "").trim().slice(0, 40);
        if (bday && !/^\d{4}-\d{2}-\d{2}$/.test(bday)) throw new Error("Invalid date");
        const n = mySide() === "spouse2" ? "2" : "1";
        await updateCouple({ [`Spouse${n}Phone`]: phone || null, [`Spouse${n}Birthday`]: bday ? D.toDate(bday) : null, AnniversaryDate: anniv ? D.toDate(anniv) : null,
          PreferredContact: ["Text", "Call", "Email"].includes(contact) ? contact : null,
          ...(state.profileEdit = false, {}), PreferredGathering: ["In person", "Virtual"].includes(gathering) ? gathering : null });
      }
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
      state.message = ["branch-message", "branch-reply"].includes(form.dataset.form) ? "" : "Saved to SharePoint."; render();
    } catch (error) { status.textContent = notice(error); submit.disabled = false; }
  });
  render();
  if (!C.clientId) return;
  try {
    auth = new msal.PublicClientApplication({ auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`, redirectUri: C.redirectUri, navigateToLoginRequestUrl: false }, cache: { cacheLocation: "sessionStorage" } });
    await auth.initialize(); const result = await auth.handleRedirectPromise();
    state.account = result?.account || auth.getAllAccounts().find(account => account.tenantId === C.tenantId) || null;
    if (state.account) { auth.setActiveAccount(state.account); await checkPush(); await loadView(); celebrateOpening(); }
  } catch (error) { state.message = notice(error); render(); }
})();
