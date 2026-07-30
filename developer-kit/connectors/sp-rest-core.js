/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-rest-core.js — session auth, request digest, and the one fetch wrapper
                     everything else in this folder is built on.

   THE WHOLE AUTHENTICATION STORY: your page is served from the SharePoint
   origin, so the browser already carries the user's session cookie. Pass
   `credentials: "same-origin"` and you are authenticated. There is no Azure AD
   app registration, no MSAL, no token cache, no refresh flow, no consent screen.

   If you are reaching for an auth library, you have misread the platform.

   Dependency-free. Paste into any app, or load it as a module in a dev setup.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SP = (() => {
  /* ── Where are we? ────────────────────────────────────────────────────── */

  /* SharePoint sets _spPageContextInfo on its own pages. Inside a Firepit web
     part your app runs from a blob: URL, so that global is NOT present and
     document.referrer / a configured site URL is what you have. Always allow an
     explicit override — it is the difference between an app that can be tested
     and one that can only be debugged in production. */
  let siteUrl = null;

  function candidateSiteUrls(configured) {
    const out = [];
    if (configured) out.push(String(configured).replace(/\/+$/, ""));
    if (typeof window !== "undefined") {
      const ctx = window._spPageContextInfo;
      if (ctx && ctx.webAbsoluteUrl) out.push(ctx.webAbsoluteUrl.replace(/\/+$/, ""));
      /* /sites/Foo/SitePages/x.aspx → https://host/sites/Foo */
      for (const source of [window.location.href, document.referrer]) {
        const m = String(source || "").match(/^(https?:\/\/[^/]+(?:\/sites\/[^/]+|\/personal\/[^/]+))/i);
        if (m) out.push(m[1]);
      }
    }
    return [...new Set(out.filter(Boolean))];
  }

  async function fetchWithTimeout(url, options, ms) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), ms || 8000);
    try {
      return await fetch(url, { ...options, signal: controller.signal });
    } finally {
      clearTimeout(timer);
    }
  }

  /* Probe each candidate until one answers. Returns null when none do, which is
     a legitimate state — see the local-fallback discussion in notes/06. Do not
     throw here; a hard failure at this point means a blank page. */
  async function detectSite(configured) {
    for (const candidate of candidateSiteUrls(configured)) {
      try {
        const r = await fetchWithTimeout(
          `${candidate}/_api/web/currentuser`,
          { credentials: "same-origin", headers: { Accept: "application/json;odata=nometadata" } },
          6000
        );
        if (r.ok) {
          siteUrl = candidate;
          return candidate;
        }
      } catch (err) {
        /* try the next candidate */
      }
    }
    return null;
  }

  const site = () => {
    if (!siteUrl) throw new Error("SP.detectSite() has not resolved a site URL yet.");
    return siteUrl;
  };

  /* ── The request digest ───────────────────────────────────────────────── */

  /* Every POST/MERGE/DELETE needs a fresh X-RequestDigest. GETs need none.
     Digests expire in roughly 30 minutes, so cache one and refresh on expiry
     rather than fetching one per write — a bulk save that requests a digest per
     record is both slow and rude to the server. */
  let digest = null;
  let digestFetchedAt = 0;
  const DIGEST_TTL_MS = 20 * 60 * 1000; // refresh well inside the ~30 min window

  async function getDigest(force) {
    const fresh = digest && Date.now() - digestFetchedAt < DIGEST_TTL_MS;
    if (fresh && !force) return digest;

    const r = await fetch(`${site()}/_api/contextinfo`, {
      method: "POST",
      credentials: "same-origin",
      headers: { Accept: "application/json;odata=nometadata" },
    });
    if (!r.ok) throw new Error(`contextinfo ${r.status}: ${await r.text()}`);
    const body = await r.json();
    digest = body.FormDigestValue;
    digestFetchedAt = Date.now();
    return digest;
  }

  /* ── The one request function ─────────────────────────────────────────── */

  /* Keep every raw fetch in one place. Nothing else in the app should call
     fetch() against SharePoint directly — when a call starts failing in
     production you want exactly one function to instrument, and one place
     holding the last URL, status, and error text for the diagnostics screen. */
  const lastCall = { url: null, status: null, error: null };

  async function request(apiPath, options) {
    const opts = options || {};
    const method = (opts.method || "GET").toUpperCase();
    const url = /^https?:/i.test(apiPath) ? apiPath : `${site()}${apiPath}`;

    /* Most endpoints prefer nometadata: smaller responses, no __metadata noise.
       A few — createfieldasxml, SendEmail — REQUIRE odata=verbose instead. Those
       pass `verbose: true`. Getting this backwards yields a 400 that reads like
       a payload problem. */
    const flavor = opts.verbose ? "verbose" : "nometadata";
    const headers = {
      Accept: `application/json;odata=${flavor}`,
      ...(opts.headers || {}),
    };

    let body = opts.body;
    if (body !== undefined && typeof body !== "string" && !(body instanceof ArrayBuffer) && !(body instanceof Blob)) {
      headers["Content-Type"] = headers["Content-Type"] || `application/json;odata=${flavor}`;
      body = JSON.stringify(body);
    }

    if (method !== "GET" && method !== "HEAD") {
      headers["X-RequestDigest"] = await getDigest();
    }

    lastCall.url = url;
    lastCall.status = null;
    lastCall.error = null;

    let r = await fetch(url, { method, credentials: "same-origin", headers, body });

    /* A 403 on a write is usually a stale digest, not a permissions problem.
       Retry exactly once with a fresh one before surfacing it — and only once,
       so a genuine permission error still fails fast. */
    if (r.status === 403 && method !== "GET" && !opts._retried) {
      headers["X-RequestDigest"] = await getDigest(true);
      r = await fetch(url, { method, credentials: "same-origin", headers, body });
    }

    lastCall.status = r.status;
    if (!r.ok) {
      lastCall.error = await r.text();
      throw new Error(`${method} ${apiPath} → ${r.status}: ${lastCall.error.slice(0, 500)}`);
    }

    if (r.status === 204) return null;
    const contentType = r.headers.get("content-type") || "";
    if (opts.raw) return r;
    if (!contentType.includes("json")) return r.text();
    return r.json();
  }

  const get = (apiPath, opts) => request(apiPath, { ...opts, method: "GET" });
  const post = (apiPath, body, opts) => request(apiPath, { ...opts, method: "POST", body });

  /* SharePoint has no PATCH. An update is a POST carrying an X-HTTP-Method
     override plus an IF-MATCH. "*" means "whatever the current version is";
     pass a real ETag when you need optimistic concurrency (see sp-files.js). */
  const merge = (apiPath, body, etag) =>
    request(apiPath, {
      method: "POST",
      body,
      headers: { "X-HTTP-Method": "MERGE", "IF-MATCH": etag || "*" },
    });

  const del = (apiPath, etag) =>
    request(apiPath, {
      method: "POST",
      headers: { "X-HTTP-Method": "DELETE", "IF-MATCH": etag || "*" },
    });

  /* ── Paging ───────────────────────────────────────────────────────────── */

  /* SharePoint pages list items at ~100 by default. Code that reads .value and
     stops silently works perfectly until the list crosses that boundary in
     production, then quietly shows partial data. Always follow the nextLink. */
  async function getAllPages(apiPath) {
    const all = [];
    let next = apiPath;
    let guard = 0;
    while (next) {
      if (++guard > 500) throw new Error("getAllPages: refusing to follow more than 500 pages");
      const page = await get(next);
      if (Array.isArray(page && page.value)) all.push(...page.value);
      next = (page && page["odata.nextLink"]) || null;
    }
    return all;
  }

  return {
    detectSite,
    site,
    setSite: (url) => {
      siteUrl = String(url).replace(/\/+$/, "");
    },
    getDigest,
    request,
    get,
    post,
    merge,
    del,
    getAllPages,
    /* Capture these three in your bug reports. They answer most "it doesn't
       work" questions without a repro. */
    lastCall,
  };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SP;

/* UNCLASSIFIED */
