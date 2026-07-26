/* AEWTTR-PULSE — unified outbound notifications (email + Teams queue). */

/* Matches the "PULSE Notifications" list's NotificationArea choices. */
const NOTIFICATION_AREAS = ["Travel", "Documents", "Projects", "Weekly", "Tickets", "Admin", "General"];

const NOTIFICATION_AREA_META = {
  Travel: { label: "Travel & Leave", emoji: "✈️", cardStyle: "accent" },
  Documents: { label: "Document Review", emoji: "📄", cardStyle: "warning" },
  Projects: { label: "Projects", emoji: "📌", cardStyle: "good" },
  Weekly: { label: "Weekly Meeting", emoji: "🪨", cardStyle: "emphasis" },
  Tickets: { label: "Support Tickets", emoji: "🎫", cardStyle: "attention" },
  Admin: { label: "Admin", emoji: "🛠️", cardStyle: "default" },
  General: { label: "PULSE", emoji: "🔔", cardStyle: "default" }
};

/* Purpose / event-kind tones — drive email accent, Teams card style, and
   purpose-specific intro/headline copy. Separate from the user-preference
   "voice" tones below (friendly/formal/…). Call sites pass `kind`. */
const NOTIFICATION_KINDS = {
  success: {
    key: "success",
    label: "Confirmed",
    headline: "All set",
    intro: "Good news — here's the confirmation.",
    accent: "#15803d",
    accentSoft: "#E8F5EE",
    teamsStyle: "good",
    teamsColor: "Good",
    icon: "✅"
  },
  action: {
    key: "action",
    label: "Action needed",
    headline: "Needs your attention",
    intro: "Something is waiting on you — open it when you can.",
    accent: "#2451C4",
    accentSoft: "#E8ECF6",
    teamsStyle: "accent",
    teamsColor: "Accent",
    icon: "👉"
  },
  denied: {
    key: "denied",
    label: "Not approved",
    headline: "Changes needed",
    intro: "This was not approved. Review the details and next steps below.",
    accent: "#B45309",
    accentSoft: "#FEF3C7",
    teamsStyle: "attention",
    teamsColor: "Warning",
    icon: "⚠️"
  },
  cancelled: {
    key: "cancelled",
    label: "Cancelled",
    headline: "Cancelled",
    intro: "This item is no longer active.",
    accent: "#64748B",
    accentSoft: "#F1F5F9",
    teamsStyle: "default",
    teamsColor: "Default",
    icon: "∅"
  },
  info: {
    key: "info",
    label: "Briefing",
    headline: "For your information",
    intro: "A short summary of what happened.",
    accent: "#0F766E",
    accentSoft: "#E6F4F1",
    teamsStyle: "emphasis",
    teamsColor: "Accent",
    icon: "📋"
  },
  comment: {
    key: "comment",
    label: "Update",
    headline: "New activity",
    intro: "A quick update on something you're involved with.",
    accent: "#5B21B6",
    accentSoft: "#F3E8FF",
    teamsStyle: "emphasis",
    teamsColor: "Default",
    icon: "💬"
  }
};
const NOTIFICATION_KIND_KEYS = Object.keys(NOTIFICATION_KINDS);
const DEFAULT_NOTIFICATION_KIND = "info";

function resolveNotificationKind(kind) {
  return NOTIFICATION_KINDS[NOTIFICATION_KIND_KEYS.includes(kind) ? kind : DEFAULT_NOTIFICATION_KIND];
}

/* User voice preferences — only reframes greeting / sign-off. "Robotic" is
   the default and keeps copy lean; purpose `kind` still shapes visuals. */
const NOTIFICATION_TONES = [
  { key: "robotic", label: "Robotic", sample: "✈️ Travel & Leave\nPULSE Travel: TR-0042 — Approved\nStatus is now Approved.", blurb: "Plain and efficient — no small talk. This is how PULSE sounds today." },
  { key: "friendly", label: "Friendly", sample: "👋 Hey Jordan!\nPULSE Travel: TR-0042 — Approved\nStatus is now Approved.", blurb: "Warm and casual, like a coworker pinging you." },
  { key: "formal", label: "Formal", sample: "Dear Jordan,\nPULSE Travel: TR-0042 — Approved\nStatus is now Approved.\n\nRegards, AEWTTR PULSE", blurb: "Professional and buttoned-up — dear-and-regards style." },
  { key: "direct", label: "Direct", sample: "Status: Approved\nPULSE Travel: TR-0042 — Approved", blurb: "Facts first, nothing extra, no greeting — for the no-nonsense reader." },
  { key: "funny", label: "Funny", sample: "Hiya, champ! 🎉\nPULSE Travel: TR-0042 — Approved\nStatus is now Approved.\n\nBeep boop, — PULSE 🤖", blurb: "A little playful, and it doesn't repeat itself — a new bit every time." }
];
const NOTIFICATION_TONE_KEYS = NOTIFICATION_TONES.map((t) => t.key);
const DEFAULT_NOTIFICATION_TONE = "robotic";

const FUNNY_TEAMS_INTROS = [
  (name) => `Hiya${name ? " " + name : ""}! 🎉 Beep boop, incoming update.`,
  (name) => `📣 Attention${name ? ` ${name}` : ""}! The PULSE hamsters have news.`,
  (name) => `🤖 *whirr* Update acquired${name ? `, ${name}` : ""}. Deploying…`,
  (name) => `Ding ding! 🔔${name ? ` ${name},` : ""} something happened.`,
  (name) => `🚨 Not a drill${name ? `, ${name}` : ""} (ish). PULSE has thoughts.`,
  (name) => `Yo${name ? ` ${name}` : ""}! 👀 PULSE spotted something.`,
  (name) => `🎬 And… action${name ? `, ${name}` : ""}! Here's the scoop.`,
  (name) => `Psst${name ? `, ${name}` : ""}… 🤫 over here.`
];
const FUNNY_EMAIL_GREETINGS = [
  (name) => `Hiya${name ? " " + name : " champ"}! 🎉`,
  (name) => `Well hello${name ? " " + name : " there"}! 👋`,
  (name) => `Greetings, ${name || "human"}. 🤖`,
  (name) => `${name || "Friend"}! Great, you're here. 🙌`,
  (name) => `Knock knock, ${name || "friend"}. 🚪`,
  (name) => `Ahoy${name ? ", " + name : ""}! ⚓`,
  (name) => `Breaking news${name ? `, ${name}` : ""}! 📰`,
  (name) => `${name || "Hey you"}, got a sec? 😎`
];
const FUNNY_SIGNOFFS = [
  "Beep boop,<br>— PULSE 🤖",
  "Stay awesome,<br>— PULSE 🚀",
  "That is all. 🫡<br>— PULSE",
  "Catch you later,<br>— PULSE 👋",
  "*mic drop*<br>— PULSE 🎤",
  "No cap,<br>— PULSE 🧢",
  "Over and out,<br>— PULSE 📡",
  "Powered by hamsters,<br>— PULSE 🐹"
];
function randomFrom(list) { return list[Math.floor(Math.random() * list.length)]; }
function funnyToneFrame(firstName) {
  return {
    teamsIntro: randomFrom(FUNNY_TEAMS_INTROS)(firstName),
    emailGreeting: randomFrom(FUNNY_EMAIL_GREETINGS)(firstName),
    emailSignoff: randomFrom(FUNNY_SIGNOFFS),
    factsFirst: false
  };
}
function randomFunnySampleText() {
  const frame = funnyToneFrame("Jordan");
  return `${frame.teamsIntro}\nPULSE Travel: TR-0042 — Approved\nStatus is now Approved.\n\n${frame.emailSignoff.replace(/<br\s*\/?>/gi, " ")}`;
}

const DOC_REVIEW_DELIVERY_MODES = ["immediate", "digest"];
const DOC_REVIEW_DIGEST_FREQUENCIES = ["daily", "weekly"];
const DEFAULT_DOC_REVIEW_DELIVERY = "immediate";
const DEFAULT_DOC_REVIEW_DIGEST_FREQUENCY = "daily";

const DIGEST_FREQUENCIES = ["daily", "weekly"];

function defaultNotificationPrefs() {
  return {
    areas: NOTIFICATION_AREAS.slice(),
    tone: DEFAULT_NOTIFICATION_TONE,
    channels: { email: true, teams: true },
    everything: true,
    documents: {
      delivery: DEFAULT_DOC_REVIEW_DELIVERY,
      digestFrequency: DEFAULT_DOC_REVIEW_DIGEST_FREQUENCY
    },
    digest: { enabled: false, frequency: "daily" }
  };
}

function normalizeNotificationPrefs(prefs) {
  if (!prefs || typeof prefs !== "object") return defaultNotificationPrefs();
  const areas = Array.isArray(prefs.areas) ? prefs.areas.filter((area) => NOTIFICATION_AREAS.includes(area)) : NOTIFICATION_AREAS.slice();
  const tone = NOTIFICATION_TONE_KEYS.includes(prefs.tone) ? prefs.tone : DEFAULT_NOTIFICATION_TONE;
  const channels = {
    email: !(prefs.channels && prefs.channels.email === false),
    teams: !(prefs.channels && prefs.channels.teams === false)
  };
  // Legacy payloads without `everything` treated all areas checked as "everything".
  const everything = typeof prefs.everything === "boolean"
    ? prefs.everything
    : areas.length === NOTIFICATION_AREAS.length;
  const docRaw = prefs.documents && typeof prefs.documents === "object" ? prefs.documents : {};
  const documents = {
    delivery: DOC_REVIEW_DELIVERY_MODES.includes(docRaw.delivery) ? docRaw.delivery : DEFAULT_DOC_REVIEW_DELIVERY,
    digestFrequency: DOC_REVIEW_DIGEST_FREQUENCIES.includes(docRaw.digestFrequency)
      ? docRaw.digestFrequency
      : DEFAULT_DOC_REVIEW_DIGEST_FREQUENCY
  };
  const digestRaw = prefs.digest && typeof prefs.digest === "object" ? prefs.digest : {};
  const digest = {
    enabled: !!digestRaw.enabled,
    frequency: DIGEST_FREQUENCIES.includes(digestRaw.frequency) ? digestRaw.frequency : "daily"
  };
  return { areas, tone, channels, everything, documents, digest };
}

function isNotificationAreaEnabledForMember(member, area) {
  if (!member || !member.notificationPrefs) return true;
  const prefs = normalizeNotificationPrefs(member.notificationPrefs);
  if (prefs.everything) return true;
  return prefs.areas.includes(area || "General");
}

function isDocReviewDigestMode(prefs) {
  const normalized = normalizeNotificationPrefs(prefs);
  return normalized.documents.delivery === "digest";
}

function findMemberByEmail(email) {
  const db = window.AEWTTR && window.AEWTTR.db;
  const members = (db && db.members) || [];
  const normalized = String(email || "").trim().toLowerCase();
  if (!normalized) return null;
  return members.find((m) => String(m.email || "").trim().toLowerCase() === normalized) || null;
}

function filterRecipientsByAreaPreference(emails, area) {
  return emails.filter((email) => isNotificationAreaEnabledForMember(findMemberByEmail(email), area));
}

function notificationConfig() {
  const db = window.AEWTTR && window.AEWTTR.db;
  if (!db) return { teamsEnabled: true, alsoSendEmail: true };
  if (!db.notificationConfig) db.notificationConfig = { teamsEnabled: true, alsoSendEmail: true };
  return db.notificationConfig;
}

function plainTextFromHtml(html) {
  if (!html) return "";
  const node = document.createElement("div");
  node.innerHTML = html;
  return (node.textContent || "").replace(/\s+/g, " ").trim();
}

function isTeamsNotificationsEnabled() {
  return !!(notificationConfig().teamsEnabled);
}

function firstNameOf(fullName) {
  if (!fullName) return "";
  const commaIdx = fullName.indexOf(",");
  if (commaIdx >= 0) {
    const rest = fullName.slice(commaIdx + 1).trim();
    return (rest.split(/\s+/)[0] || "").trim();
  }
  return (String(fullName).trim().split(/\s+/)[0] || "").trim();
}

function pulseAppUrl() {
  const pageName = (typeof APP_CONFIG !== "undefined" && APP_CONFIG.pulseAppPageFileName) || "AEWTTR-PULSE.aspx";
  const siteUrl = window.AEWTTR && window.AEWTTR.siteUrl;
  if (siteUrl) {
    return `${String(siteUrl).replace(/\/$/, "")}/SitePages/${pageName}`;
  }
  const fallbackPath = (typeof APP_CONFIG !== "undefined" && APP_CONFIG.pulseAppPagePath)
    || "/sites/AEWTTRTst/SitePages/AEWTTR-PULSE.aspx";
  return `${location.origin}${fallbackPath.startsWith("/") ? fallbackPath : `/${fallbackPath}`}`;
}

function pulseAppRouteUrl(path, query) {
  const route = String(path || "").replace(/^#?\/?/, "");
  const queryString = query && typeof query === "object"
    ? Object.keys(query)
      .filter((key) => query[key] != null && query[key] !== "")
      .map((key) => `${encodeURIComponent(key)}=${encodeURIComponent(String(query[key]))}`)
      .join("&")
    : "";
  const hash = route ? `#/${route}${queryString ? `?${queryString}` : ""}` : "";
  return `${pulseAppUrl()}${hash}`;
}

function notifyEscape(s) {
  if (typeof escapeHtml === "function") return escapeHtml(s);
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* User voice frame — greeting/sign-off only. */
function toneFrame(tone, opts) {
  opts = opts || {};
  const area = opts.area || "General";
  const name = opts.firstName || "";
  const areaMeta = NOTIFICATION_AREA_META[area] || NOTIFICATION_AREA_META.General;
  switch (tone) {
    case "friendly":
      return {
        tone: "friendly",
        teamsIntro: `👋 Hey${name ? " " + name : ""}!`,
        emailGreeting: `Hi${name ? " " + name : " there"},`,
        emailSignoff: "Thanks!<br>— PULSE",
        factsFirst: false
      };
    case "formal":
      return {
        tone: "formal",
        teamsIntro: "**Notification**",
        emailGreeting: `Dear ${name || "Team Member"},`,
        emailSignoff: "Regards,<br>AEWTTR PULSE",
        factsFirst: false
      };
    case "direct":
      return { tone: "direct", teamsIntro: "", emailGreeting: "", emailSignoff: "", factsFirst: true };
    case "funny": {
      const fFrame = funnyToneFrame(name);
      fFrame.tone = "funny";
      return fFrame;
    }
    case "robotic":
    default:
      return {
        tone: "robotic",
        teamsIntro: `${areaMeta.emoji} ${areaMeta.label}`,
        emailGreeting: "",
        emailSignoff: "",
        factsFirst: false
      };
  }
}

function purposeIntroFor(kindMeta, voiceTone) {
  if (voiceTone === "direct") return "";
  if (voiceTone === "funny") return kindMeta.intro;
  if (voiceTone === "formal") {
    if (kindMeta.key === "action") return "Your attention is requested for the following matter.";
    if (kindMeta.key === "success") return "Please find confirmation of the following.";
    if (kindMeta.key === "denied") return "Please review the outcome and any requested follow-up below.";
    if (kindMeta.key === "cancelled") return "Please note that the following item has been cancelled.";
    if (kindMeta.key === "comment") return "Please see the activity noted below.";
    return "Please find the following briefing.";
  }
  if (voiceTone === "friendly") {
    if (kindMeta.key === "action") return "Heads-up — this one's waiting on you.";
    if (kindMeta.key === "success") return "Nice one — you're all set.";
    if (kindMeta.key === "denied") return "Not quite — take a look at what to do next.";
    if (kindMeta.key === "cancelled") return "Quick note: this one's been cancelled.";
    if (kindMeta.key === "comment") return "Someone left an update for you.";
    return "Thought you'd want to know.";
  }
  return kindMeta.intro;
}

function wrapEmailChrome({ subject, area, kind, preview, facts, actionUrl, actionTitle, frame, innerHtml }) {
  const kindMeta = resolveNotificationKind(kind);
  const areaMeta = NOTIFICATION_AREA_META[area] || NOTIFICATION_AREA_META.General;
  const purposeIntro = purposeIntroFor(kindMeta, frame && frame.tone);
  const fontStack = "-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif";
  const factList = (facts || []).filter((fact) => fact && fact.title != null);
  const factRows = factList.map((fact, index) => {
    const isLast = index === factList.length - 1;
    const rule = isLast ? "none" : "1px solid #ECEEF2";
    return `
    <tr>
      <td style="padding:11px 0;border-bottom:${rule};color:#6B7280;font-size:12px;line-height:1.4;width:132px;vertical-align:top;font-family:${fontStack};">${notifyEscape(fact.title)}</td>
      <td style="padding:11px 0 11px 16px;border-bottom:${rule};color:#111827;font-size:14px;line-height:1.45;vertical-align:top;font-family:${fontStack};">${notifyEscape(String(fact.value == null || fact.value === "" ? "—" : fact.value))}</td>
    </tr>`;
  }).join("");

  const greeting = frame && frame.emailGreeting
    ? `<p style="margin:0 0 14px;font-size:14px;line-height:1.5;color:#111827;font-family:${fontStack};">${frame.emailGreeting}</p>`
    : "";
  const signoff = frame && frame.emailSignoff
    ? `<p style="margin:24px 0 0;font-size:13px;line-height:1.5;color:#6B7280;font-family:${fontStack};">${frame.emailSignoff}</p>`
    : "";
  const previewBlock = preview
    ? `<p style="margin:0 0 18px;font-size:14px;line-height:1.6;color:#374151;font-family:${fontStack};">${notifyEscape(preview)}</p>`
    : "";
  const introBlock = purposeIntro
    ? `<p style="margin:0 0 12px;font-size:14px;line-height:1.55;color:#4B5563;font-family:${fontStack};">${notifyEscape(purposeIntro)}</p>`
    : "";
  const factsTable = factRows
    ? `<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin:4px 0 8px;">${factRows}</table>`
    : "";
  const cta = actionUrl
    ? `<table role="presentation" cellpadding="0" cellspacing="0" style="margin:22px 0 0;"><tr><td style="border-radius:6px;background:${kindMeta.accent};">
        <a href="${notifyEscape(actionUrl)}" style="display:inline-block;padding:11px 20px;font-size:13px;font-weight:600;line-height:1.2;color:#FFFFFF;text-decoration:none;font-family:${fontStack};">${notifyEscape(actionTitle || "Open in PULSE")}</a>
      </td></tr></table>`
    : "";
  const bodyContent = innerHtml
    ? `${introBlock}<div style="margin:0 0 18px;font-size:14px;line-height:1.6;color:#374151;font-family:${fontStack};">${innerHtml}</div>${factsTable}`
    : `${introBlock}${previewBlock}${factsTable}`;

  return `<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<meta name="color-scheme" content="light only" />
<meta name="supported-color-schemes" content="light only" />
<title>${notifyEscape(subject || kindMeta.headline)}</title>
<!--[if mso]><style>table,td{font-family:Arial,Helvetica,sans-serif !important;}</style><![endif]-->
</head>
<body style="margin:0;padding:0;background:#EEF0F3;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="width:100%;background:#EEF0F3;">
  <tr>
    <td align="center" style="padding:28px 16px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="580" style="width:100%;max-width:580px;background:#FFFFFF;border:1px solid #E2E5EA;">
        <tr>
          <td style="height:3px;line-height:3px;font-size:0;background:${kindMeta.accent};">&nbsp;</td>
        </tr>
        <tr>
          <td style="padding:22px 28px 0;background:#FFFFFF;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="font-family:${fontStack};font-size:15px;font-weight:700;letter-spacing:0.04em;color:#1B3A6B;line-height:1.2;">PULSE</td>
                <td align="right" style="font-family:${fontStack};font-size:11px;color:#9CA3AF;line-height:1.2;">AEWTTR</td>
              </tr>
            </table>
            <p style="margin:6px 0 0;font-family:${fontStack};font-size:12px;color:#9CA3AF;line-height:1.3;">${notifyEscape(areaMeta.label)}</p>
          </td>
        </tr>
        <tr>
          <td style="padding:20px 28px 8px;background:#FFFFFF;">
            <p style="margin:0 0 8px;font-family:${fontStack};font-size:11px;font-weight:600;letter-spacing:0.06em;text-transform:uppercase;color:${kindMeta.accent};line-height:1.2;">${notifyEscape(kindMeta.label)}</p>
            <h1 style="margin:0;font-family:${fontStack};font-size:20px;font-weight:700;line-height:1.35;color:#111827;">${notifyEscape(subject || kindMeta.headline)}</h1>
          </td>
        </tr>
        <tr>
          <td style="padding:8px 28px 28px;background:#FFFFFF;">
            ${greeting}
            ${bodyContent}
            ${cta}
            ${signoff}
          </td>
        </tr>
        <tr>
          <td style="padding:14px 28px 18px;background:#FFFFFF;border-top:1px solid #ECEEF2;">
            <p style="margin:0;font-family:${fontStack};font-size:11px;line-height:1.45;color:#9CA3AF;">Sent by AEWTTR PULSE · Manage preferences in Settings</p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>`;
}

function buildEmailBodyHtml({ subject, area, kind, preview, facts, actionUrl, actionTitle, frame, body }) {
  frame = frame || {};
  // If the caller passed a custom HTML body (minutes, comments), still wrap
  // it in the branded chrome so every email shares one design system.
  const looksLikeFragment = body && !/^\s*<!DOCTYPE/i.test(body) && !/<html[\s>]/i.test(body);
  if (body && !looksLikeFragment) return body;
  return wrapEmailChrome({
    subject,
    area,
    kind,
    preview: body ? "" : preview,
    facts: facts || [],
    actionUrl,
    actionTitle,
    frame,
    innerHtml: body || ""
  });
}

function buildTeamsMarkdownText({ subject, area, kind, preview, facts, actionUrl, actionTitle, frame }) {
  frame = frame || {};
  const kindMeta = resolveNotificationKind(kind);
  const areaMeta = NOTIFICATION_AREA_META[area] || NOTIFICATION_AREA_META.General;
  const tone = frame.tone || "robotic";
  const purposeIntro = purposeIntroFor(kindMeta, tone);
  const blocks = [];
  const factLines = (facts || []).map((f) => `**${f.title}:** ${f.value == null || f.value === "" ? "—" : f.value}`).join("  \n");
  const subjectBlock = `**${subject || "PULSE notification"}**`;
  const kindLine = `${kindMeta.icon} **${kindMeta.label}** · ${areaMeta.label}`;

  if (tone === "friendly") {
    if (frame.teamsIntro) blocks.push(frame.teamsIntro);
    blocks.push(kindLine);
    if (purposeIntro) blocks.push(purposeIntro);
    blocks.push(subjectBlock);
    if (preview) blocks.push(preview);
    if (factLines) blocks.push(factLines);
    blocks.push("Thanks,  \nPULSE");
  } else if (tone === "formal") {
    blocks.push(frame.teamsIntro || "**Official Notification**");
    blocks.push(kindLine);
    if (purposeIntro) blocks.push(purposeIntro);
    blocks.push(subjectBlock);
    if (preview) blocks.push(preview);
    if (factLines) blocks.push(factLines);
    blocks.push("Regards,  \nAEWTTR PULSE");
  } else if (tone === "funny") {
    if (frame.teamsIntro) blocks.push(frame.teamsIntro);
    blocks.push(kindLine);
    blocks.push(subjectBlock);
    if (preview) blocks.push(preview);
    if (factLines) blocks.push(`The nitty gritty:\n${factLines}`);
    if (frame.emailSignoff) blocks.push(frame.emailSignoff.replace(/<br\s*\/?>/gi, " "));
  } else if (tone === "direct") {
    blocks.push(kindLine);
    if (factLines) blocks.push(factLines);
    blocks.push(subjectBlock);
    if (preview) blocks.push(preview);
  } else {
    blocks.push(kindLine);
    if (purposeIntro) blocks.push(purposeIntro);
    blocks.push(subjectBlock);
    if (preview) blocks.push(preview);
    if (factLines) blocks.push(factLines);
  }

  if (actionUrl) blocks.push(`[${actionTitle || "Open in PULSE"}](${actionUrl})`);
  return blocks.filter(Boolean).join("\n\n");
}

function buildTeamsAdaptiveCard({ subject, area, kind, preview, facts, actionUrl, actionTitle, frame }) {
  frame = frame || {};
  const tone = frame.tone || "robotic";
  const kindMeta = resolveNotificationKind(kind);
  const areaMeta = NOTIFICATION_AREA_META[area] || NOTIFICATION_AREA_META.General;
  const purposeIntro = purposeIntroFor(kindMeta, tone);
  const cardStyle = kindMeta.teamsStyle || areaMeta.cardStyle || "default";

  const card = {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.4",
    body: [],
    msteams: { width: "Full" }
  };

  const items = [];

  if (tone === "funny" && frame.teamsIntro) {
    items.push({ type: "TextBlock", text: frame.teamsIntro, wrap: true, weight: "Bolder", color: "Attention", spacing: "None" });
  } else if (tone === "friendly" && frame.teamsIntro) {
    items.push({ type: "TextBlock", text: frame.teamsIntro, wrap: true, weight: "Bolder", spacing: "None" });
  } else if (tone === "formal") {
    items.push({ type: "TextBlock", text: "NOTICE", size: "Small", weight: "Bolder", isSubtle: true, spacing: "None" });
  }

  items.push({
    type: "ColumnSet",
    spacing: items.length ? "Small" : "None",
    columns: [
      {
        type: "Column",
        width: "stretch",
        items: [
          {
            type: "TextBlock",
            text: `${kindMeta.icon} ${kindMeta.label}`,
            weight: "Bolder",
            color: kindMeta.teamsColor || "Accent",
            size: "Small",
            spacing: "None"
          },
          {
            type: "TextBlock",
            text: areaMeta.label,
            isSubtle: true,
            size: "Small",
            spacing: "None",
            wrap: true
          }
        ]
      }
    ]
  });

  items.push({
    type: "TextBlock",
    text: subject || kindMeta.headline,
    weight: "Bolder",
    size: "Large",
    wrap: true,
    spacing: "Medium"
  });

  if (purposeIntro && tone !== "direct") {
    items.push({ type: "TextBlock", text: purposeIntro, wrap: true, isSubtle: true, spacing: "Small" });
  }

  if (preview) {
    items.push({ type: "TextBlock", text: preview, wrap: true, spacing: "Small" });
  }

  if (facts && facts.length) {
    items.push({
      type: "FactSet",
      spacing: "Medium",
      facts: facts.map((fact) => ({
        title: String(fact.title || ""),
        value: String(fact.value == null || fact.value === "" ? "—" : fact.value)
      }))
    });
  }

  if (tone === "funny" && frame.emailSignoff) {
    items.push({
      type: "TextBlock",
      text: frame.emailSignoff.replace(/<br\s*\/?>/gi, " "),
      spacing: "Large",
      isSubtle: true,
      wrap: true
    });
  } else if (tone === "formal") {
    items.push({
      type: "TextBlock",
      text: "Regards,\nAEWTTR PULSE",
      wrap: true,
      spacing: "Large",
      isSubtle: true
    });
  } else if (tone === "friendly") {
    items.push({
      type: "TextBlock",
      text: "Thanks — PULSE",
      spacing: "Large",
      isSubtle: true
    });
  }

  card.body.push({
    type: "Container",
    style: cardStyle,
    bleed: true,
    items
  });

  if (actionUrl) {
    card.actions = [{ type: "Action.OpenUrl", title: actionTitle || "Open in PULSE", url: actionUrl }];
  }
  return card;
}

function buildNotificationPayload({ subject, area, kind, preview, facts, actionUrl, actionTitle, body, teamsText, tone, recipientFirstName }) {
  const resolvedTone = NOTIFICATION_TONE_KEYS.includes(tone) ? tone : DEFAULT_NOTIFICATION_TONE;
  const resolvedKind = NOTIFICATION_KIND_KEYS.includes(kind) ? kind : DEFAULT_NOTIFICATION_KIND;
  const resolvedPreview = preview || teamsText || plainTextFromHtml(body || "") || subject || "";
  const resolvedFacts = facts || [];
  const frame = toneFrame(resolvedTone, { area, firstName: recipientFirstName });
  const emailBodyHtml = buildEmailBodyHtml({
    subject,
    area,
    kind: resolvedKind,
    preview: resolvedPreview,
    facts: resolvedFacts,
    actionUrl,
    actionTitle,
    frame,
    body
  });
  const adaptiveCard = buildTeamsAdaptiveCard({
    subject,
    area,
    kind: resolvedKind,
    preview: resolvedPreview,
    facts: resolvedFacts,
    actionUrl,
    actionTitle,
    frame
  });
  return {
    emailBodyHtml,
    bodyHtml: JSON.stringify(adaptiveCard),
    teamsText: teamsText || buildTeamsMarkdownText({
      subject,
      area,
      kind: resolvedKind,
      preview: resolvedPreview,
      facts: resolvedFacts,
      actionUrl,
      actionTitle,
      frame
    })
  };
}

async function notifyUsers({ to, subject, area, kind, preview, facts, actionUrl, actionTitle, body, teamsText, forceEmail, forceTeams, toneOverride, skipDocDigestGate }) {
  const rawEmails = (Array.isArray(to) ? to : [to]).map((email) => String(email || "").trim()).filter(Boolean);
  const emails = (forceEmail || forceTeams) ? rawEmails : filterRecipientsByAreaPreference(rawEmails, area || "General");
  if (!emails.length || typeof isSharePointMode !== "function" || !isSharePointMode()) {
    return { email: { sent: false }, teams: { queued: false } };
  }

  const resolvedArea = area || "General";
  const immediateEmails = [];
  let digestHeld = 0;
  emails.forEach((email) => {
    if (forceEmail || forceTeams || skipDocDigestGate || resolvedArea !== "Documents") {
      immediateEmails.push(email);
      return;
    }
    const member = findMemberByEmail(email);
    const prefs = normalizeNotificationPrefs(member && member.notificationPrefs);
    if (isDocReviewDigestMode(prefs)) {
      // Digest preference is stored (incl. daily/weekly), but there is no
      // shared SharePoint digest queue / Power Automate cron yet — so we
      // hold back the immediate ping rather than misfire a per-browser queue.
      digestHeld += 1;
      return;
    }
    immediateEmails.push(email);
  });
  if (!immediateEmails.length) {
    return { email: { sent: false }, teams: { queued: false }, digestHeld };
  }

  const groups = new Map();
  immediateEmails.forEach((email) => {
    const member = findMemberByEmail(email);
    const prefs = normalizeNotificationPrefs(member && member.notificationPrefs);
    const tone = toneOverride && NOTIFICATION_TONE_KEYS.includes(toneOverride) ? toneOverride : prefs.tone;
    if (!groups.has(tone)) groups.set(tone, { emailTo: [], teamsTo: [] });
    const g = groups.get(tone);
    if (forceEmail || prefs.channels.email) g.emailTo.push(email);
    if (forceTeams || prefs.channels.teams) g.teamsTo.push(email);
  });

  const cfg = notificationConfig();
  const siteUrl = typeof currentSiteUrl === "function" ? currentSiteUrl() : "";
  let emailResult = { sent: false };
  let teamsResult = { queued: false };

  for (const [tone, group] of groups) {
    const anchorEmail = group.emailTo[0] || group.teamsTo[0];
    const firstMember = anchorEmail ? findMemberByEmail(anchorEmail) : null;
    const payload = buildNotificationPayload({
      subject, area, kind, preview, facts, actionUrl, actionTitle, body, teamsText,
      tone, recipientFirstName: firstMember ? firstNameOf(firstMember.name) : ""
    });

    if (group.emailTo.length && (forceEmail || cfg.alsoSendEmail !== false)) {
      try {
        emailResult = await sharePointAdapter.sendEmail(siteUrl, { to: group.emailTo, subject, body: payload.emailBodyHtml });
      } catch (e) {
        emailResult = { sent: false, error: (e && (e.friendly || e.message)) || String(e) };
        console.warn("PULSE: email notification failed.", e);
      }
    }

    if (group.teamsTo.length && (forceTeams || isTeamsNotificationsEnabled())) {
      try {
        teamsResult = await sharePointAdapter.queueTeamsNotification(siteUrl, {
          to: group.teamsTo,
          subject,
          bodyHtml: payload.bodyHtml,
          emailBodyHtml: payload.emailBodyHtml,
          teamsText: payload.teamsText,
          area: area || "General"
        });
      } catch (e) {
        teamsResult = { queued: false, error: (e && (e.friendly || e.message)) || String(e) };
        console.warn("PULSE: Teams notification queue failed.", e);
      }
    }
  }

  return { email: emailResult, teams: teamsResult, digestHeld };
}

function travelRequestsForCurrentUser() {
  const db = window.AEWTTR && window.AEWTTR.db;
  if (!db) return [];
  const meEmail = String((db.user && db.user.email) || "").trim().toLowerCase();
  const meName = String((db.user && db.user.name) || "").trim().toLowerCase();
  return (db.travelRequests || []).filter((request) => {
    const requesterEmail = String(request.requesterEmail || "").trim().toLowerCase();
    const requesterName = String(request.requester || "").trim().toLowerCase();
    if (meEmail && requesterEmail === meEmail) return true;
    if (meName && requesterName === meName) return true;
    return (request.travelers || []).some((traveler) => {
      if (traveler && (traveler.type === "group" || traveler.isGroup)) {
        const expanded = typeof expandPickerPeople === "function" ? expandPickerPeople([traveler]) : [];
        return expanded.some((person) => {
          const travelerEmail = String(person.email || "").trim().toLowerCase();
          const travelerName = String(person.name || "").trim().toLowerCase();
          return (meEmail && travelerEmail === meEmail) || (meName && travelerName === meName);
        });
      }
      const travelerEmail = String(traveler.email || "").trim().toLowerCase();
      const travelerName = String(traveler.name || "").trim().toLowerCase();
      return (meEmail && travelerEmail === meEmail) || (meName && travelerName === meName);
    });
  });
}

function travelStatusPreview(request) {
  if (!request) return "Your travel request has a new status.";
  const isLeave = String(request.formMode || request.requestType || request.category || "").toLowerCase().includes("leave")
    || request.requestType === "Personal Leave"
    || request.category === "Leave";
  switch (request.status) {
    case "Pending Finance":
      return "Admin approved — sent to Finance for charge object assignment";
    case "Approved":
      if (isLeave) return "Customer concurrence recorded for your leave request.";
      return request.chargeObject
        ? `Fully approved with C/O ${request.chargeObject}`
        : "Travel request fully approved";
    case "Denied":
      return request.denyReason ? `Denied: ${request.denyReason}` : "Travel request was denied";
    case "Revoked":
      return "Request withdrawn";
    case "Cancelled":
      if (isLeave) return "Leave request was cancelled";
      return request.cancelledByAdmin
        ? `Travel cancelled by ${request.cancelledBy || "an admin"}`
        : "Travel request was cancelled";
    case "Pending":
      return isLeave ? "Submitted and awaiting customer concurrence" : "Submitted and awaiting admin review";
    default:
      return `Status is now ${request.status}`;
  }
}

function travelNotificationFacts(request) {
  const isLeave = String(request.formMode || request.requestType || request.category || "").toLowerCase().includes("leave")
    || request.requestType === "Personal Leave"
    || request.category === "Leave";
  const travelerNames = (request.travelers || []).map((t) => t.name).join(", ") || request.requester;
  const facts = [
    { title: "Request", value: `${request.id} — ${request.tripTitle || (isLeave ? "Leave Request" : "Travel Request")}` }
  ];
  if (isLeave) {
    facts.push({ title: "Requester", value: request.requester || travelerNames });
  } else {
    facts.push({ title: "Travelers", value: travelerNames });
    facts.push({ title: "Destination", value: request.destination || "—" });
  }
  facts.push({
    title: "Dates",
    value: (request.allDay === false || (request.allDay == null && (request.startTime || request.endTime)))
      ? `${request.start || ""}${request.startTime ? ` ${request.startTime}` : ""} to ${request.end || ""}${request.endTime ? ` ${request.endTime}` : ""}`
      : `${request.start || ""} to ${request.end || ""}`
  });
  if (request.status) {
    facts.push({
      title: "Status",
      value: isLeave && request.status === "Approved" ? "Customer Concurrence" : request.status
    });
  }
  if (request.notes && isLeave) facts.push({ title: "Notes", value: request.notes });
  if (request.chargeObject) facts.push({ title: "Charge object", value: request.chargeObject });
  return facts;
}

async function sendTestNotificationForCurrentUser(toneOverride) {
  if (typeof isSharePointMode !== "function" || !isSharePointMode()) {
    throw new Error("Test messages only work in SharePoint mode.");
  }
  const db = window.AEWTTR.db;
  const user = db.user || {};
  const email = String(user.email || (window.AEWTTR.currentSpUser && window.AEWTTR.currentSpUser.email) || "").trim();
  if (!email) throw new Error("No email found for the current user.");

  return notifyUsers({
    to: [email],
    subject: "PULSE Travel: TR-0042 — Approved",
    area: "Travel",
    kind: "success",
    preview: "Status is now Approved.",
    facts: [
      { title: "Request", value: "TR-0042 — Sample Travel Request" },
      { title: "Destination", value: "Washington, DC" },
      { title: "Status", value: "Approved" }
    ],
    actionUrl: pulseAppUrl(),
    actionTitle: "Open PULSE",
    forceEmail: true,
    forceTeams: true,
    toneOverride
  });
}

async function sendTestTravelStatusNotification() {
  if (typeof isSharePointMode !== "function" || !isSharePointMode()) {
    throw new Error("Test messages only work in SharePoint mode.");
  }
  const db = window.AEWTTR.db;
  const user = db.user || {};
  const email = String(user.email || (window.AEWTTR.currentSpUser && window.AEWTTR.currentSpUser.email) || "").trim();
  if (!email) throw new Error("No email found for the current user.");

  const mine = travelRequestsForCurrentUser();
  const request = mine[0] || {
    id: "TR-TEST",
    tripTitle: "Sample Travel Request",
    destination: "Washington, DC",
    start: new Date().toISOString().slice(0, 10),
    end: new Date(Date.now() + 3 * 86400000).toISOString().slice(0, 10),
    status: "Pending Finance"
  };
  const statusLine = travelStatusPreview(request);
  const subject = `PULSE Travel test: ${request.id} status update`;

  return notifyUsers({
    to: [email],
    subject,
    area: "Travel",
    kind: "info",
    preview: `Test notification — ${statusLine}`,
    facts: travelNotificationFacts(request),
    actionUrl: pulseAppUrl(),
    actionTitle: "Open PULSE",
    forceEmail: true,
    forceTeams: true
  });
}

window.NOTIFICATION_AREAS = NOTIFICATION_AREAS;
window.NOTIFICATION_AREA_META = NOTIFICATION_AREA_META;
window.NOTIFICATION_KINDS = NOTIFICATION_KINDS;
window.NOTIFICATION_TONES = NOTIFICATION_TONES;
window.DOC_REVIEW_DELIVERY_MODES = DOC_REVIEW_DELIVERY_MODES;
window.DOC_REVIEW_DIGEST_FREQUENCIES = DOC_REVIEW_DIGEST_FREQUENCIES;
window.defaultNotificationPrefs = defaultNotificationPrefs;
window.normalizeNotificationPrefs = normalizeNotificationPrefs;
window.randomFunnySampleText = randomFunnySampleText;
window.notifyUsers = notifyUsers;
window.isTeamsNotificationsEnabled = isTeamsNotificationsEnabled;
window.notificationConfig = notificationConfig;
window.sendTestTravelStatusNotification = sendTestTravelStatusNotification;
window.sendTestNotificationForCurrentUser = sendTestNotificationForCurrentUser;
window.buildTeamsAdaptiveCard = buildTeamsAdaptiveCard;
window.buildTeamsMarkdownText = buildTeamsMarkdownText;
window.pulseAppUrl = pulseAppUrl;
window.pulseAppRouteUrl = pulseAppRouteUrl;
window.travelNotificationFacts = travelNotificationFacts;
window.firstNameOf = firstNameOf;
window.findMemberByEmail = findMemberByEmail;
window.resolveNotificationKind = resolveNotificationKind;
window.isDocReviewDigestMode = isDocReviewDigestMode;
