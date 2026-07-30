/**
 * PULSE IPT Dashboard — v1.0.0
 * AEWTTR IPT DB46200
 *
 * Developer:  Eli Scalere (Contractor)
 * Primary:    elijah.t.scalere.ctr@us.navy.mil
 * Secondary:  eli.scalere@scaleredesign.com  |  (516) 265-2636
 *
 * Built for NAVAIR AEWTTR IPT. Single-file SharePoint SPA.
 * All source under apps/PULSE/. Release builds in releases/.
 */

/* AEWTTR-PULSE — single source of app-level configuration.
   Do not duplicate these values anywhere else in the codebase. */

const APP_VERSION = "1.0.0";

/* ---------- boot diagnostic (remove once Firepit loading is confirmed stable) ---------- */
(function () {
  function showBootError(msg) {
    var root = document.getElementById("aewttr-root");
    if (!root) return;
    root.innerHTML = '<div style="padding:28px;font-family:sans-serif;max-width:640px;">'
      + '<h2 style="color:#b00;margin:0 0 10px;font-size:18px;">PULSE boot error</h2>'
      + '<pre style="background:#f5f5f5;padding:12px;border-radius:4px;font-size:12px;'
      + 'white-space:pre-wrap;word-break:break-all;">' + String(msg).replace(/</g, "&lt;") + '</pre>'
      + '<p style="font-size:12px;color:#666;margin:8px 0 0;">Report this to your app admin.</p></div>';
  }

  window.addEventListener("error", function (e) {
    if (window.AEWTTR && window.AEWTTR.bootComplete) {
      console.error("PULSE runtime error:", e.message, e.filename, e.lineno, e.error);
      if (typeof toast === "function") toast("Something went wrong on this page. Try again or refresh.", "error");
      return;
    }
    showBootError((e.filename ? e.filename.split("/").pop() : "?") + ":" + e.lineno + " — " + e.message);
  });
  window.addEventListener("unhandledrejection", function (e) {
    if (window.AEWTTR && window.AEWTTR.bootComplete) {
      // Background failures (polls, prefetches) already surface their own
      // targeted toasts where it matters — a generic toast here just spams
      // "a background action failed" on every transient network blip.
      console.error("PULSE unhandled rejection:", e.reason);
      return;
    }
    showBootError("Unhandled promise rejection: " + (e.reason && e.reason.message ? e.reason.message : String(e.reason)));
  });

  // If boot never renders, show a diagnostic.  Firepit can inject a flat FS
  // package after the host page has already fired DOMContentLoaded, so start
  // this watchdog immediately in that environment rather than attaching a
  // listener that will never run.
  function scheduleBootWatchdog() {
    setTimeout(function () {
      if (window.AEWTTR && (window.AEWTTR.bootComplete || window.AEWTTR.db)) return; // app initialized OK
      if (window.AEWTTR && window.AEWTTR.bootStarted && !window.AEWTTR.bootFailed) return; // boot still in progress
      var missing = [];
      if (typeof normalizeStoreShape === "undefined") missing.push("data.js");
      if (typeof sharePointAdapter === "undefined") missing.push("sharepoint-adapter.js");
      if (typeof aewttrLoadStore === "undefined") missing.push("data.js/aewttrLoadStore");
      if (typeof renderShell === "undefined") missing.push("app.js/renderShell");
      showBootError("App did not initialize within 15 s.\n"
        + "Scripts that appear missing: " + (missing.length ? missing.join(", ") : "(all present — check app.js DOMContentLoaded)")
        + "\nAEWTTR global: " + JSON.stringify(window.AEWTTR));
    }, 15000);
  }

  if (document.readyState === "loading") {
    window.addEventListener("DOMContentLoaded", scheduleBootWatchdog, { once: true });
  } else {
    scheduleBootWatchdog();
  }
}());

const APP_CONFIG = {
  // Leave blank to auto-detect the SharePoint site URL from _spPageContextInfo
  // or the current page location. Only set this if auto-detection fails in
  // your environment (e.g. the app is hosted in a Firepit web part on a
  // page whose URL doesn't directly expose the host site).
  manualSharePointSiteUrl: "",

  // SharePoint list display names start with this prefix.
  listPrefix: "PULSE",

  // Legacy file-store settings kept only for backward compatibility with
  // older diagnostics. The active backend is SharePoint Lists.
  dataLibraryTitle: "PULSE App Data",
  dataFolderName: "AEWTTR-PULSE",
  dataFiles: {
    db: "db.json",
    roles: "roles.json",
    peopleDirectory: "people-directory.json"
  },

  // SharePoint Site Page that hosts the PULSE app (Teams notification links).
  pulseAppPageFileName: "AEWTTR-PULSE.aspx",
  // Used only when site URL auto-detection is unavailable (local dev).
  pulseAppPagePath: "/sites/AEWTTRTest/SitePages/AEWTTR-PULSE.aspx",

  // Absolute URL always used for the "Open PULSE" button in emails. Emails
  // are opened outside any live page/iframe context, so they can't rely on
  // auto-detecting the current site the way in-app links do — this fixed
  // URL is what every notification email's CTA button links to. Leave
  // blank to fall back to site auto-detection (see pulseAppUrl in notify.js).
  pulseAppFullUrl: "https://flankspeed.sharepoint-mil.us/sites/AEWTTRTest/SitePages/AEWTTR-PULSE.aspx",

  // Role assigned when an admin assigns someone to PULSE (creates App Roles).
  defaultUserRole: "Member",

  // SharePoint AssociatedMemberGroup (+ Owners) still feed the bulk directory
  // pull when permission expansion is empty. These exact group Title fallbacks
  // are tried only if AssociatedMemberGroup lookup fails.
  siteMemberGroupNames: [
    "AEWTTRTest Members",
    "AEWTTR IPT Members",
    "AEWTTR Members",
    "PULSE IPT Members"
  ],

  // Microsoft Graph base URL — MUST match your tenant's actual cloud, or
  // every Graph call (tenant-wide people search, M365 group expansion)
  // fails silently with an auth/network error while SharePoint's own REST
  // API keeps working fine (it's a different hostname entirely, on your
  // own tenant's SharePoint domain, so a wrong Graph endpoint doesn't break
  // it — this is why "site users" can load while "search the whole
  // tenant" doesn't). A *.sharepoint-mil.us site collection is the DoD
  // (IL4/IL5) cloud, not GCC High — set to the DoD Graph endpoint.
  //   Commercial / GCC : https://graph.microsoft.com
  //   GCC High (Flank Speed/NMCI, *.sharepoint.us) : https://graph.microsoft.us
  //   DoD (IL4/IL5, *.sharepoint-mil.us) : https://dod-graph.microsoft.us
  graphApiBase: "https://dod-graph.microsoft.us",

  // ---------- AI grammar review (Document Review page) ----------
  // Talks to CapraGPT's OpenAI-compatible chat completions endpoint.
  // >>> PUT YOUR CAPRAGPT API KEY IN apiKey BELOW. <<<
  // If CapraGPT's actual endpoint URL differs from the placeholder, update
  // `endpoint` to match what CapraGPT's own docs specify — everything else
  // (request/response shape, Bearer auth) follows the standard OpenAI
  // chat-completions contract and shouldn't need to change.
  aiReview: {
    enabled: true,
    endpoint: "https://api.capragpt.mil/v1/chat/completions", // CONFIRM against CapraGPT's docs
    apiKey: "", // <-- put your CapraGPT API key here, as a plain string
    model: "gpt-4.1"
  }
};
