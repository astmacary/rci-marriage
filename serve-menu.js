/* RCI Serve menu for the Marriage app.
   When the app is opened from RCI Serve, a teams button appears beside the light/dark button.
   It opens a sheet (like the team apps' Switch Team) with All Teams Home and the person's teams,
   read from the hub's localStorage "rci-serve:state". */
(function(){
  "use strict";
  var TEAMS = {
    productions: { name: "Productions", path: "/rci-productions/", c: "#0B3A6E", icon: '<rect x="2" y="6" width="14" height="12" rx="2"/><path d="m22 8-6 4 6 4V8z"/>' },
    security: { name: "Security", path: "/rci-security/", c: "#7A1F1F", icon: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>' },
    ushers: { name: "Ushers & Altar Workers", path: "/rci-ushers/", c: "#1E5E43", icon: '<path d="M3 21h18"/><path d="M5 21V7l7-4 7 4v14"/><path d="M10 21v-6h4v6"/>' },
    worship: { name: "Worship & Arts", path: "/rci-worship-arts/", c: "#4A2F7A", icon: '<path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/>' }
  };
  var RINGS = '<circle cx="9" cy="14" r="5"/><circle cx="15" cy="14" r="5"/><path d="m10 5 2-2 2 2"/>';
  var GRID = '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>';
  var inHub = false; try { inHub = sessionStorage.getItem("rci-serve:hub") === "1"; } catch(e){}
  if (!inHub && !/[?&]from=serve\b/.test(location.search)) return;
  var toggle = document.getElementById("theme-toggle"); if (!toggle) return;
  var svg = function(inner, w){ w = w || 18; return '<svg viewBox="0 0 24 24" width="' + w + '" height="' + w + '" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + inner + '</svg>'; };
  var esc = function(s){ return String(s == null ? "" : s).replace(/[&<>"']/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]; }); };
  var state = function(){ try { return JSON.parse(localStorage.getItem("rci-serve:state") || "null"); } catch(e){ return null; } };

  var css = document.createElement("style");
  css.textContent =
    ".sm-wrap{display:flex;align-items:center;gap:8px;flex-shrink:0}" +
    ".sm-btn{position:relative}" +
    ".sm-dot{position:absolute;top:-1px;right:-1px;width:10px;height:10px;border-radius:50%;background:#C62828;border:2px solid #0A2440}" +
    ".sm-back{position:fixed;inset:0;z-index:90;background:rgba(3,17,31,.55)}" +
    ".sm-sheet{position:fixed;left:0;right:0;bottom:0;z-index:91;max-width:560px;margin:0 auto;background:var(--surface,#fff);color:var(--ink,#0A2440);border-radius:20px 20px 0 0;padding:10px 16px calc(24px + env(safe-area-inset-bottom,0px));display:flex;flex-direction:column;gap:6px;box-shadow:0 -8px 30px rgba(3,17,31,.3);max-height:85vh;overflow:auto}" +
    ".sm-grab{width:40px;height:5px;border-radius:3px;background:var(--soft,#ddd);align-self:center;margin-bottom:6px}" +
    ".sm-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:4px}.sm-head h2{margin:0;font-size:22px;letter-spacing:.04em;text-transform:uppercase}" +
    ".sm-x{width:40px;height:40px;border-radius:50%;border:0;background:#0A2440;color:#F3EEE6;display:grid;place-items:center;cursor:pointer}" +
    ".sm-row{display:flex;align-items:center;gap:12px;min-height:58px;padding:0 12px;border-radius:12px;border:1px solid var(--line2,#ccc);text-decoration:none;color:inherit}" +
    ".sm-row.sm-home{border:2px solid #C4955F}.sm-row.sm-here{border:2px solid #0A2440}:root[data-theme=dark] .sm-row.sm-here{border-color:#D4A66E}" +
    ".sm-ic{width:36px;height:36px;border-radius:10px;color:#fff;display:grid;place-items:center;flex:none}" +
    ".sm-info{flex:1;min-width:0}.sm-info b{display:block}.sm-info small{display:block;font-size:12px;opacity:.7}" +
    ".sm-count{min-width:22px;height:22px;padding:0 6px;border-radius:11px;background:#C62828;color:#fff;font-size:12px;font-weight:700;display:inline-flex;align-items:center;justify-content:center}" +
    ".sm-label{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;opacity:.65;margin:10px 0 2px 4px}";
  document.head.appendChild(css);

  var wrap = document.createElement("span"); wrap.className = "sm-wrap";
  toggle.parentNode.insertBefore(wrap, toggle);
  var btn = document.createElement("button");
  btn.type = "button"; btn.className = "theme-toggle sm-btn"; btn.setAttribute("aria-label", "RCI Serve teams");
  wrap.appendChild(btn); wrap.appendChild(toggle);
  function pending(s){ return s && Array.isArray(s.teams) ? s.teams.reduce(function(a, t){ return a + (Number(t.count) || 0); }, 0) : 0; }
  function paint(){ var n = pending(state()); btn.innerHTML = svg(GRID, 15) + (n ? '<span class="sm-dot" aria-hidden="true"></span>' : ""); }
  paint();

  var back = null, sheet = null;
  function close(){ if (back) back.remove(); if (sheet) sheet.remove(); back = sheet = null; btn.focus(); }
  function row(href, color, icon, name, sub, extra, cls){
    return '<a class="sm-row' + (cls ? " " + cls : "") + '" href="' + href + '"><span class="sm-ic" style="background:' + color + '">' + svg(icon) + '</span><span class="sm-info"><b>' + esc(name) + '</b><small>' + esc(sub) + '</small></span>' + (extra || "") + '</a>';
  }
  btn.addEventListener("click", function(){
    var s = state(), n = pending(s);
    var teams = (s && Array.isArray(s.teams) ? s.teams : []).map(function(st){
      var t = TEAMS[st.key]; if (!t) return "";
      var c = Number(st.count) || 0;
      return row(t.path, t.c, t.icon, t.name, c ? c + (c === 1 ? " needs a reply" : " need a reply") : (st.admin ? "Admin · Nothing waiting" : "Nothing waiting"), c ? '<span class="sm-count">' + c + '</span>' : "");
    }).join("");
    back = document.createElement("div"); back.className = "sm-back"; back.addEventListener("click", close);
    sheet = document.createElement("section"); sheet.className = "sm-sheet"; sheet.setAttribute("role", "dialog"); sheet.setAttribute("aria-modal", "true"); sheet.setAttribute("aria-label", "RCI Serve");
    sheet.innerHTML = '<div class="sm-grab" aria-hidden="true"></div><div class="sm-head"><h2>RCI Serve</h2><button type="button" class="sm-x" aria-label="Close">' + svg('<path d="M18 6 6 18M6 6l12 12"/>') + '</button></div>' +
      row("/", "#0B1D33", GRID, "RCI Serve", n ? n + (n === 1 ? " request" : " requests") + " on your teams" : "All Teams Home", '<span>' + svg('<path d="m9 6 6 6-6 6"/>') + '</span>', "sm-home") +
      (teams ? '<div class="sm-label">Your teams</div>' + teams : "") +
      '<div class="sm-label">Your ministries</div>' + row("/rci-marriage/", "#B8864B", RINGS, "Marriage", "You're here", '<span>' + svg('<path d="M20 6 9 17l-5-5"/>', 20) + '</span>', "sm-here");
    sheet.querySelector(".sm-x").addEventListener("click", close);
    sheet.addEventListener("keydown", function(ev){ if (ev.key === "Escape") close(); });
    document.body.appendChild(back); document.body.appendChild(sheet);
    sheet.querySelector(".sm-x").focus();
  });
  window.addEventListener("storage", function(e){ if (e.key === "rci-serve:state") paint(); });
})();
