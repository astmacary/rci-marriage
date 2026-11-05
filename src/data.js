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
  const api = { parse, own, escape, safeLink, validate };
  root.RCIData = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window === "undefined" ? globalThis : window);
