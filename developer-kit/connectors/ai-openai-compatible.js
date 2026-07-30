/* ═══════════════════════════════════════════════════════════════════════════
   UNCLASSIFIED

   ai-openai-compatible.js — calling an LLM from a browser-only app.

   PULSE CODE talks to Ask Sage, which speaks the OpenAI chat-completions shape,
   so this connector works against any endpoint with that contract.

   ── READ THIS BEFORE YOU SHIP AN API KEY ──────────────────────────────────

   There is no server in this architecture, which means there is nowhere to hide a
   secret. Any key your app can send, a user can read — out of localStorage, out of
   the network tab, or out of the shipped HTML itself. Consequences:

     · NEVER hardcode a key in source. It ends up in a .html file in a document
       library, in git, and in every decompile of that package.
     · A key in a SharePoint config file is readable by everyone with read access
       to that library. That may be acceptable for a shared team key on an internal
       site — it is a deliberate decision to make with whoever owns the key, not a
       default to fall into.
     · Prefer a per-user key entered in settings and kept in that user's
       localStorage, so a leak is scoped to one person and revocable.
     · Assume any key shipped this way is compromised eventually. Know how to
       rotate it before you need to.

   ── AND THE CSP ──────────────────────────────────────────────────────────

   Inside a Firepit web part your frame gets `connect-src 'self' <sharepoint>
   <mcas…> <dataverse?>`. An arbitrary AI host is NOT on that list, so this call
   will be blocked outright.

   And the CSP is only half of it. Firepit also injects a NETWORK GUARD that
   monkeypatches fetch/XHR/WebSocket/EventSource/sendBeacon and checks the origin
   itself (notes/02 §5). Both layers must permit a host, and a block by the second
   one produces a rejected promise with NO CSP violation logged — which sends
   people debugging their API key for an hour.

   Getting an origin allow-listed is a platform request, not a code change. In this
   tenant the allow-list is reported to include government AI endpoints
   (api.genai.mil and similar) — confirm with your administrator rather than
   assuming. Verify in a real web part early: this is the most common way an AI
   feature works on a desktop and dies in production.

   Requires nothing. Optionally uses sp-files.js for shared-config storage.
═══════════════════════════════════════════════════════════════════════════ */

"use strict";

const AI = (() => {
  const defaults = {
    endpoint: "https://api.asksage.ai/server/openai/v1/chat/completions",
    model: "gpt-4.1-mini",
    maxTokens: 2048,
    temperature: 0.2,
    timeoutMs: 90000,
  };

  let config = { ...defaults, apiKey: "" };

  const configure = (patch) => {
    config = { ...config, ...(patch || {}) };
    return { ...config, apiKey: config.apiKey ? "(set)" : "" };
  };

  const isConfigured = () => !!(config.apiKey && config.endpoint && config.model);

  /* ── The call ─────────────────────────────────────────────────────────── */

  async function chat(messages, options) {
    const opts = options || {};
    if (!isConfigured()) throw new Error("AI is not configured — no API key.");

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), opts.timeoutMs || config.timeoutMs);

    try {
      const r = await fetch(opts.endpoint || config.endpoint, {
        method: "POST",
        signal: controller.signal,
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${config.apiKey}`,
        },
        /* Deliberately NOT credentials:'include'. Sending SharePoint session
           cookies to a third-party host is the kind of mistake that is invisible
           until someone reviews the traffic. */
        body: JSON.stringify({
          model: opts.model || config.model,
          messages,
          max_tokens: opts.maxTokens || config.maxTokens,
          temperature: opts.temperature != null ? opts.temperature : config.temperature,
        }),
      });

      if (!r.ok) {
        const text = await r.text();
        /* Distinguish the failures a user can act on from the ones they cannot. */
        if (r.status === 401 || r.status === 403) {
          throw new Error(`AI rejected the key (${r.status}). Re-enter it in Settings.`);
        }
        if (r.status === 429) {
          throw new Error("AI rate limit reached. Wait a moment and retry.");
        }
        throw new Error(`AI ${r.status}: ${text.slice(0, 400)}`);
      }

      const body = await r.json();
      const content = body && body.choices && body.choices[0] && body.choices[0].message;
      return {
        text: (content && content.content) || "",
        usage: body && body.usage,
        raw: body,
      };
    } catch (err) {
      if (err.name === "AbortError") throw new Error("AI request timed out.");
      /* A platform block surfaces as an opaque TypeError with no status. Say so,
         because the default message ("Failed to fetch") sends people debugging
         their key. Name both layers: the absence of a CSP violation in the console
         does NOT mean the request was allowed to leave. */
      if (err instanceof TypeError) {
        throw new Error(
          "AI request could not leave the page. Inside a Firepit web part this is " +
            "almost always the platform blocking the AI host — either the injected " +
            "connect-src CSP (look for a violation in the console) or the network " +
            "guard that patches window.fetch (no violation logged). Both must " +
            "allow-list the origin. See notes/02 §4–5."
        );
      }
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  /* ── Conversation helper ──────────────────────────────────────────────── */

  /* Keeps a bounded history. An unbounded one grows until the model rejects the
     request, and the error arrives at the worst moment — mid-conversation, on the
     user's most complex question. */
  function conversation(systemPrompt, options) {
    const opts = options || {};
    const maxTurns = opts.maxTurns || 12;
    const history = [];

    return {
      history,
      async ask(userText, askOptions) {
        history.push({ role: "user", content: userText });
        while (history.length > maxTurns * 2) history.shift();

        const messages = systemPrompt
          ? [{ role: "system", content: systemPrompt }, ...history]
          : [...history];

        const reply = await chat(messages, askOptions);
        history.push({ role: "assistant", content: reply.text });
        return reply;
      },
      reset() {
        history.length = 0;
      },
    };
  }

  /* ── Shared configuration in SharePoint ───────────────────────────────── */

  /* PULSE CODE stores its key in SiteAssets/pulse-code-config.json so a team
     shares one key without each person pasting it. Convenient, and a deliberate
     trade: anyone with read access to that library can read the key. Do this only
     with the key owner's agreement, and never for a key with real cost or blast
     radius attached. */
  async function loadSharedConfig(fileServerRelativeUrl) {
    if (typeof SPFiles === "undefined") throw new Error("loadSharedConfig requires sp-files.js");
    try {
      const text = await SPFiles.readText(fileServerRelativeUrl);
      const parsed = JSON.parse(text);
      configure(parsed);
      return parsed;
    } catch (err) {
      return null; // absent config is a normal first-run state
    }
  }

  async function saveSharedConfig(folderServerRelativeUrl, fileName) {
    if (typeof SPFiles === "undefined") throw new Error("saveSharedConfig requires sp-files.js");
    const payload = { endpoint: config.endpoint, model: config.model, apiKey: config.apiKey };
    await SPFiles.upload(folderServerRelativeUrl, fileName || "ai-config.json", JSON.stringify(payload, null, 2), {
      contentType: "application/json",
    });
  }

  return { configure, isConfigured, chat, conversation, loadSharedConfig, saveSharedConfig, defaults };
})();

if (typeof module !== "undefined" && module.exports) module.exports = AI;

/* Example:

     AI.configure({ apiKey: localStorage.getItem("ai-key") || "" });

     const chat = AI.conversation(
       "You are a concise assistant embedded in an internal SharePoint dashboard."
     );
     const { text } = await chat.ask("Summarize this project status: " + JSON.stringify(project));
*/

/* UNCLASSIFIED */
