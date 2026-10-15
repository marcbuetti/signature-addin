/*
 * Signature library — shared by the event handler (commands.js) and the task pane (taskpane.js).
 * Each builder returns EMAIL-SAFE HTML with inline styles only (CSS classes don't survive in mail).
 * In production the company data + which templates a user gets would come from the SharePoint profile.
 */
(function () {
  var BRAND = "#2f9e5f", DEEP = "#1f7a47", INK = "#17201b", MUT = "#5d6b64", META = "#6b7a72";
  var CO = "Vegilife AG", WEB = "www.vegilife.ch", ADDR = "Solothurnstrasse 12 · CH-2544 Bettlach";
  var PHONE = "+41 32 512 85 00", INFO = "info@vegilife.ch", TAG = "Pflanzlich. Schweizer Qualität.";
  var FONT = "'IBM Plex Sans',Arial,Helvetica,sans-serif";

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function wrap(inner) {
    return '<div style="font-family:' + FONT + ';font-size:13px;line-height:1.5;color:#2a332e">' + inner + "</div>";
  }
  function mail(addr) {
    return '<a href="mailto:' + esc(addr) + '" style="color:' + DEEP + ';text-decoration:none">' + esc(addr) + "</a>";
  }
  function web() {
    return '<a href="https://' + WEB + '" style="color:' + DEEP + ';text-decoration:none">' + WEB + "</a>";
  }
  function lab(t) {
    return '<span style="display:inline-block;width:16px;color:#9aa8a0;font-weight:700">' + t + "</span>";
  }

  // Standard — personal signature with contact block (accent bar)
  function standard(p) {
    var name = esc(p.displayName), email = p.emailAddress || INFO;
    return wrap(
      '<div style="border-left:3px solid ' + BRAND + ';padding-left:14px">' +
      '<div style="font-weight:700;font-size:15px;color:' + INK + '">' + name + "</div>" +
      '<div style="color:' + DEEP + ';font-weight:700;margin-top:2px">' + CO + "</div>" +
      '<div style="margin-top:6px">' + lab("T") + esc(PHONE) + "</div>" +
      "<div>" + lab("E") + mail(email) + "</div>" +
      "<div>" + lab("W") + web() + "</div>" +
      '<div style="color:' + META + ';font-size:12px;margin-top:6px">' + esc(ADDR) + "</div>" +
      "</div>"
    );
  }

  // Team — generic company signature, no personal name
  function team() {
    return wrap(
      '<div style="border-left:3px solid ' + BRAND + ';padding-left:14px">' +
      '<div style="font-weight:700;font-size:15px;color:' + INK + '">' + CO + "</div>" +
      '<div style="color:' + MUT + '">' + esc(TAG) + "</div>" +
      '<div style="margin-top:6px">' + lab("T") + esc(PHONE) + "</div>" +
      "<div>" + lab("E") + mail(INFO) + "</div>" +
      "<div>" + lab("W") + web() + "</div>" +
      '<div style="color:' + META + ';font-size:12px;margin-top:6px">' + esc(ADDR) + "</div>" +
      "</div>"
    );
  }

  // Kurz — compact one-liner
  function kurz(p) {
    var name = esc(p.displayName), email = p.emailAddress || INFO;
    return wrap(
      '<div><span style="font-weight:700;color:' + INK + '">' + name +
      '</span><span style="color:#9aa8a0;margin:0 7px">·</span><span style="color:' + DEEP + ';font-weight:700">' + CO + "</span></div>" +
      '<div style="color:' + META + ';font-size:12px;margin-top:3px">' + esc(PHONE) + " · " + mail(email) + " · " + web() + "</div>"
    );
  }

  window.SignatureLib = {
    list: [
      { id: "standard", label: "Standard", desc: "Persönlich, mit Kontaktblock" },
      { id: "team", label: "Team", desc: "Allgemein, ohne Namen" },
      { id: "kurz", label: "Kurz", desc: "Kompakt, zweizeilig" }
    ],
    defaultId: "standard",
    build: function (id, profile) {
      profile = profile || {};
      if (id === "team") return team();
      if (id === "kurz") return kurz(profile);
      return standard(profile);
    }
  };
})();
