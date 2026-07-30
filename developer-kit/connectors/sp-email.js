/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   sp-email.js — sending mail with no server, no Graph, no Power Automate.

   SharePoint exposes SP.Utilities.Utility.SendEmail, which sends as the site
   from the user's session. Two hard rules:

     1. This endpoint requires odata=VERBOSE. nometadata returns a 400 that reads
        like a malformed body.
     2. Recipients must already have access to the site. It will not mail
        arbitrary external addresses.

   And one design rule that is not optional:

     TREAT A FAILED NOTIFICATION AS NON-FATAL. Some tenants disable outbound mail
     from this endpoint by policy. Save the data first, notify second, and never
     let a missed email roll back or block real work. A user whose travel request
     vanished because the confirmation email failed will not forgive the design.

   Requires sp-rest-core.js.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const SPEmail = (() => {
  const asArray = (v) => (Array.isArray(v) ? v : v == null ? [] : [v]);

  /* ── Send ─────────────────────────────────────────────────────────────── */

  /* Returns { sent: true } or { sent: false, reason } — deliberately does not
     throw, so a caller cannot accidentally make mail failure fatal by forgetting
     a try/catch. If you genuinely want the throw, use sendOrThrow(). */
  async function send(message) {
    try {
      await sendOrThrow(message);
      return { sent: true };
    } catch (err) {
      const reason = String(err.message || err);
      /* Worth distinguishing in your logs: a tenant policy block is permanent and
         should stop you retrying, a 5xx is worth one retry. */
      const policyBlocked = /→ 40[0-9]/.test(reason);
      return { sent: false, reason, policyBlocked };
    }
  }

  async function sendOrThrow(message) {
    const to = asArray(message.to).filter(Boolean);
    if (!to.length) throw new Error("sendEmail: no recipients");

    const properties = {
      __metadata: { type: "SP.Utilities.EmailProperties" },
      To: { results: to },
      Subject: message.subject || "(no subject)",
      /* Body is HTML. SharePoint does not sanitize it for you and it is rendered
         in a mail client, so escape anything user-supplied — see html() below. */
      Body: message.body || "",
    };
    if (message.cc) properties.CC = { results: asArray(message.cc).filter(Boolean) };
    if (message.bcc) properties.BCC = { results: asArray(message.bcc).filter(Boolean) };
    if (message.from) properties.From = message.from;

    /* verbose: true is mandatory here. */
    await SP.request("/_api/SP.Utilities.Utility.SendEmail", {
      method: "POST",
      verbose: true,
      body: { properties },
    });
  }

  /* ── Composing safely ─────────────────────────────────────────────────── */

  const esc = (s) =>
    String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  /* Tagged template that escapes every interpolated value:
       html`<p>Hello ${userName}</p>`
     Anything you want to inject as raw markup must be passed through raw(). */
  function html(strings, ...values) {
    return strings.reduce((out, chunk, i) => {
      if (i === 0) return chunk;
      const value = values[i - 1];
      const rendered = value && value.__raw ? value.value : esc(value);
      return out + rendered + chunk;
    }, "");
  }
  const raw = (value) => ({ __raw: true, value });

  /* A plain, legible layout. Mail clients strip most CSS; inline styles and
     tables survive, flexbox and external stylesheets do not. */
  function layout(opts) {
    const rows = (opts.facts || [])
      .map(
        (f) =>
          `<tr><td style="padding:4px 12px 4px 0;color:#555;white-space:nowrap">${esc(f.label)}</td>` +
          `<td style="padding:4px 0"><strong>${esc(f.value)}</strong></td></tr>`
      )
      .join("");

    return (
      `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222;line-height:1.45">` +
      `<p style="margin:0 0 12px">${esc(opts.lead || "")}</p>` +
      (rows ? `<table style="border-collapse:collapse;margin:0 0 12px">${rows}</table>` : "") +
      (opts.bodyHtml || "") +
      (opts.linkUrl
        ? `<p style="margin:16px 0 0"><a href="${esc(opts.linkUrl)}">${esc(opts.linkText || "Open")}</a></p>`
        : "") +
      (opts.footer
        ? `<p style="margin:20px 0 0;color:#777;font-size:12px">${esc(opts.footer)}</p>`
        : "") +
      `</div>`
    );
  }

  /* ── Notifying a group without spamming it ────────────────────────────── */

  /* Two habits worth keeping: honour a per-user preference record, and de-duplicate
     recipients — the same person often matches two rules (assignee AND approver)
     and getting the mail twice reads as a broken app. */
  async function notify(recipients, message, options) {
    const opts = options || {};
    const wanted = [...new Set(asArray(recipients).map((r) => String(r || "").trim().toLowerCase()).filter(Boolean))];
    const allowed = wanted.filter((email) => (opts.isSubscribed ? opts.isSubscribed(email) : true));

    if (!allowed.length) return { sent: false, reason: "no subscribed recipients", skipped: wanted.length };

    /* One message with many recipients, not many messages — this endpoint is
       rate-sensitive and a loop over 40 addresses will start failing partway. */
    const result = await send({ ...message, to: allowed });
    return { ...result, recipients: allowed.length, skipped: wanted.length - allowed.length };
  }

  return { send, sendOrThrow, notify, html, raw, esc, layout };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SPEmail;

/* Example — save first, notify second, never the reverse:

     const saved = await SPList.createItem("Travel Requests", fields);   // must succeed

     const outcome = await SPEmail.notify(approverEmails, {
       subject: `Travel request from ${user.displayName}`,
       body: SPEmail.layout({
         lead: `${user.displayName} submitted a travel request for your approval.`,
         facts: [
           { label: "Destination", value: fields.Destination },
           { label: "Departs",     value: fields.DepartDate },
           { label: "Estimate",    value: `$${fields.EstimatedCost}` },
         ],
         linkUrl: `${SP.site()}/SitePages/PULSE.aspx#/travel/${saved.Id}`,
         linkText: "Review in PULSE",
         footer: "Sent by PULSE. Reply to the requester, not to this message.",
       }),
     }, {
       isSubscribed: (email) => prefs[email]?.travelApprovals !== false,
     });

     if (!outcome.sent) console.warn("Notification not delivered:", outcome.reason);
     // ...and the request is still saved. That is the point.
*/

/* UNCLASSIFIED */
