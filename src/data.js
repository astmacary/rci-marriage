// Data helpers contain no live records. SharePoint is the authorization boundary.
(function (root) {
  const parse = item => {
    try {
      const data = JSON.parse(item.fields?.AppData || "{}");
      return data && !Array.isArray(data) && typeof data === "object" ? data : {};
    } catch { return {}; }
  };
  const own = (items, oid) => items.filter(item =>
    String(item.createdBy?.user?.id || "").toLowerCase() === String(oid).toLowerCase());
  const escape = value => String(value ?? "").replace(/[&<>"']/g,
    c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const safeLink = value => {
    try { const url = new URL(value); return url.protocol === "https:" ? url.href : ""; }
    catch { return ""; }
  };
  const validate = data => {
    if (!data || data.version !== 1 || !["prompt-answer", "check-in", "journal-entry", "rsvp"].includes(data.kind))
      throw new Error("Please refresh the app before saving.");
    if (JSON.stringify(data).length > 24000) throw new Error("Please shorten your response before saving.");
    return data;
  };
  // SharePoint date-only columns: write noon UTC so the day never shifts; read back as YYYY-MM-DD.
  const toDate = day => day ? `${day}T12:00:00Z` : null;
  const fromDate = value => value ? String(value).slice(0, 10) : "";
  const str = v => v == null ? "" : String(v);
  // Read a row from its named columns, falling back to the older AppData JSON.
  const read = (key, item) => {
    const f = item?.fields || {}, old = parse(item);
    if (key === "content") return { ...old,
      kind: f.Kind ? str(f.Kind).toLowerCase() : old.kind,
      date: f.PromptDate ? fromDate(f.PromptDate) : old.date,
      type: f.PromptType || old.type, week: f.Week || old.week,
      text: f.Kind ? str(f.Title) : (old.text ?? str(f.Title)),
      url: f.Link || old.url };
    if (key === "couples") {
      const emails = [f.Spouse1Email, f.Spouse2Email].filter(Boolean);
      return { ...old, kind: old.kind || (emails.length ? "couple" : undefined),
        memberEmails: emails.length ? emails : old.memberEmails,
        branch: f.Branch || old.branch, mentorCouple: f.MentorCouple || old.mentorCouple,
        coupleKey: f.CoupleKey || old.coupleKey || str(item?.id),
        journalItemId: f.JournalItemId || old.journalItemId };
    }
    if (key === "answers") return { ...old,
      date: f.PromptDate ? fromDate(f.PromptDate) : old.date, answer: f.Answer ?? old.answer };
    if (key === "checkins") return { ...old,
      date: f.WeekOf ? fromDate(f.WeekOf) : old.date,
      ratings: f.Connected != null ? [f.Connected, f.Heard, f.Spiritual, f.Fun] : old.ratings,
      need: f.Need ?? old.need, saw: f.SawInYou ?? old.saw,
      mentor: f.MentorRequest ?? old.mentor };
    return old;
  };
  // Turn validated app data into the list's named columns.
  const columns = (data, coupleKey) => {
    if (data.kind === "prompt-answer") return { PromptDate: toDate(data.date), Answer: str(data.answer), CoupleKey: str(coupleKey) };
    if (data.kind === "check-in") {
      const r = data.ratings.map(n => (Number.isInteger(n) && n >= 1 && n <= 5) ? n : null);
      return { WeekOf: toDate(data.date), Connected: r[0], Heard: r[1], Spiritual: r[2], Fun: r[3],
        Need: str(data.need), SawInYou: str(data.saw), MentorRequest: !!data.mentor, CoupleKey: str(coupleKey) };
    }
    throw new Error("Please refresh the app before saving.");
  };
  // Items the flows write into a couple's journal folder: reveals and shared check-ins.
  // SharePoint permissions already limit what a spouse can read; this also skips the folder
  // itself and anything tagged for another couple (site owners can see every folder).
  const shared = (items, couple) => (items || []).map(item => {
    const f = item?.fields || {}, kind = str(f.Kind).toLowerCase();
    if (!["reveal", "check-in"].includes(kind)) return null;
    if (couple?.journalItemId && str(item.id) === str(couple.journalItemId)) return null;
    if (f.CoupleKey && couple?.coupleKey && str(f.CoupleKey) !== str(couple.coupleKey)) return null;
    return { id: str(item.id), kind, title: str(f.Title),
      date: fromDate(f.PromptDate || f.WeekOf || item.createdDateTime),
      spouse1: str(f.Spouse1Answer), spouse2: str(f.Spouse2Answer) };
  }).filter(Boolean).sort((a, b) => b.date.localeCompare(a.date) || Number(b.id) - Number(a.id));
  // Which side of a shared item belongs to the signed-in spouse.
  const side = (couple, email) => {
    const emails = (couple?.memberEmails || []).map(v => String(v).toLowerCase());
    const i = emails.indexOf(String(email || "").toLowerCase());
    return i === 1 ? "spouse2" : i === 0 ? "spouse1" : "";
  };
  const api = { shared, side, parse, own, escape, safeLink, validate, read, columns, toDate, fromDate };
  root.RCIData = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window === "undefined" ? globalThis : window);
