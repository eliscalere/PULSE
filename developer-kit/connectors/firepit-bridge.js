/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   firepit-bridge.js — talking to the Firepit host web part from inside the frame.

   Your app runs in a sandboxed iframe loaded from a blob: URL. A few things are
   therefore impossible for you and possible for the host, and the host exposes
   them over postMessage. Protocol read from Firepit 7.1.0
   (firepit-webpart/src/webparts/htmlFileRenderer/HtmlFileRendererWebPart.ts).

   | Message type                          | Direction    | Purpose                        |
   |---------------------------------------|--------------|--------------------------------|
   | firepit:open-edit-panel               | child → host | Open the SharePoint property pane |
   | firepit:dataverse-request             | child → host | Proxy a Dataverse call            |
   | firepit:dataverse-response            | host → child | Correlated reply, by requestId    |
   | firepit:destructive-command-approved  | child → host | Record session approval           |

   The host validates event.source === its own iframe before acting, so these only
   work from the hosted frame — which also means they no-op harmlessly when you
   run the same app standalone. Guard on isHosted() rather than feature-detecting.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const Firepit = (() => {
  const MSG = {
    OPEN_EDIT_PANEL: "firepit:open-edit-panel",
    DATAVERSE_REQUEST: "firepit:dataverse-request",
    DATAVERSE_RESPONSE: "firepit:dataverse-response",
    DESTRUCTIVE_APPROVED: "firepit:destructive-command-approved",
  };

  /* We are hosted if we are in a frame at all. There is no handshake in 7.1.0, so
     this is the honest test — do not infer more than it tells you. */
  const isHosted = () => {
    try {
      return typeof window !== "undefined" && window.parent && window.parent !== window;
    } catch (err) {
      return true; // cross-origin access threw, which means we are framed
    }
  };

  const post = (payload) => {
    if (!isHosted()) return false;
    try {
      window.parent.postMessage(payload, "*");
      return true;
    } catch (err) {
      return false;
    }
  };

  /* ── The CSP nonce ────────────────────────────────────────────────────── */

  /* If the host page runs a nonce-based CSP, Firepit forwards the nonce into your
     frame and rewrites your static <script> tags to carry it. Scripts you create
     at RUNTIME are yours to stamp. */
  const cspNonce = () => (typeof window !== "undefined" ? window.__firepitCspNonce || "" : "");

  function createScript(code) {
    const el = document.createElement("script");
    const nonce = cspNonce();
    if (nonce) el.setAttribute("nonce", nonce);
    el.textContent = code;
    return el;
  }

  /* ── Property pane ────────────────────────────────────────────────────── */

  const openEditPanel = () => post({ type: MSG.OPEN_EDIT_PANEL });

  /* ── Dataverse proxy ──────────────────────────────────────────────────── */

  /* You cannot obtain an AAD token inside the sandboxed frame. The host can,
     through AadHttpClient and the tenant-approved "Microsoft Dataverse /
     user_impersonation" scope declared in the .sppkg. So you describe the request
     and the host performs it.

     Requires the web part's `dataverseEnvironmentUrl` property to be set — that
     property is also what adds the Dataverse origin to the frame's connect-src.

     Remember the host's throttle: destructive operations are limited to 5 per
     60-second window per target, plus a session approval gate. Design bulk work
     around that rather than discovering it at record six. */
  const pending = new Map();
  let listening = false;
  let sequence = 0;

  function startListening() {
    if (listening || typeof window === "undefined") return;
    listening = true;
    window.addEventListener("message", (event) => {
      const data = event && event.data;
      if (!data || typeof data !== "object" || data.type !== MSG.DATAVERSE_RESPONSE) return;
      const entry = pending.get(data.requestId);
      if (!entry) return;
      pending.delete(data.requestId);
      clearTimeout(entry.timer);

      if (data.ok) {
        entry.resolve({
          ok: true,
          status: data.status,
          headers: data.headers || {},
          body: data.body,
          json: () => {
            try {
              return JSON.parse(data.body);
            } catch (err) {
              return null;
            }
          },
        });
      } else {
        entry.reject(
          new Error(data.error || `Dataverse ${data.status} ${data.statusText || ""}`.trim())
        );
      }
    });
  }

  function dataverse(url, init, options) {
    const opts = options || {};
    if (!isHosted()) {
      return Promise.reject(new Error("Dataverse proxy requires the Firepit host; this app is not framed."));
    }
    startListening();

    const requestId = `dv-${++sequence}-${Date.now()}`;
    return new Promise((resolve, reject) => {
      /* Always time out. A host that never replies — wrong Firepit version, the
         dataverseEnvironmentUrl property unset — would otherwise leave the promise
         pending forever and the calling UI stuck on a spinner. */
      const timer = setTimeout(() => {
        pending.delete(requestId);
        reject(
          new Error(
            `Dataverse request timed out after ${opts.timeoutMs || 30000}ms. ` +
              `Check that the web part's dataverseEnvironmentUrl property is set.`
          )
        );
      }, opts.timeoutMs || 30000);

      pending.set(requestId, { resolve, reject, timer });

      const sent = post({
        type: MSG.DATAVERSE_REQUEST,
        requestId,
        url,
        init: {
          method: (init && init.method) || "GET",
          headers: (init && init.headers) || {},
          body: init && init.body,
        },
      });

      if (!sent) {
        pending.delete(requestId);
        clearTimeout(timer);
        reject(new Error("Could not post to the Firepit host."));
      }
    });
  }

  /* ── Destructive-operation approval ───────────────────────────────────── */

  /* Tell the host the user has approved destructive commands for this session.
     Send it only in response to an actual, explicit user confirmation — it is a
     safety gate, and firing it on load defeats the whole mechanism. */
  const approveDestructiveCommands = () => post({ type: MSG.DESTRUCTIVE_APPROVED });

  return { MSG, isHosted, cspNonce, createScript, openEditPanel, dataverse, approveDestructiveCommands };
})();

if (typeof module !== "undefined" && module.exports) module.exports = Firepit;

/* Example:

     if (Firepit.isHosted()) {
       const res = await Firepit.dataverse(
         "/api/data/v9.2/accounts?$select=name&$top=5",
         { headers: { Accept: "application/json" } }
       );
       console.log(res.json().value);
     } else {
       // Standalone: fall back to local data rather than failing.
     }
*/

/* UNCLASSIFIED */
