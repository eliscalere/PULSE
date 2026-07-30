<!-- UNCLASSIFIED -->

# 06 — Postmortems

**UNCLASSIFIED**

Bugs that cost real time on this program, in roughly the order they were hit. Each
one is here because the *symptom pointed somewhere other than the cause* — which
is exactly the kind of thing a document can save you from and a code review
cannot.

Read this before building the equivalent piece yourself.

---

## 1. `Options: 0` on column creation — the root of a dozen symptoms

**Symptoms, all at once:** columns named `field_1`, `field_7`; columns missing from
the add/edit form; lists opening with only a `Title` column; and the nastiest —
writes returning **201 Created** while the SharePoint form stayed empty.

**Cause:** `createfieldasxml` called without `Options: 28`. Without flag `8`,
SharePoint ignores your `Name` attribute and invents an internal name. Your app
then writes to a property that does not exist. SharePoint accepts the POST,
discards the unknown field, and returns success.

**Why it took so long:** every symptom looked like a different bug. Duplicate
columns looked like a race in setup; the blank form looked like a permissions or
content-type issue; the 201-with-no-data looked like caching. One flag caused all
of it.

**Fix:** `Options: 28`, then verify internal names after creating, then repair the
default view idempotently. [sp-columns.js](../connectors/sp-columns.js).

**Cost:** severe enough that the program abandoned lists for a JSON document store
mid-project. That migration was avoidable.

---

## 2. `?v=N` cache-busting silently 404s on SharePoint

**Symptom:** a completely blank white page. No error, no failed assertion, an
empty console.

**Cause:** `<script src="app.js?v=10">`. SharePoint's static file serving
mishandles query strings on script URLs and the request 404s — *silently*. Nothing
downstream ever ran, so nothing logged.

**Why it took so long:** an empty console is the least informative failure there
is. The instinct is to debug the app, but the app never started.

**Fix:** never put version query strings on script or link tags in a file served
from SharePoint. Use a hard refresh or dev-tools "disable cache" during
development. Note this applies to unbundled multi-file dev setups too, not just
shipped packages.

---

## 3. `async` boot blocking first paint

**Symptom:** on government/CAC-gated networks, the app showed a blank page
indefinitely. Fine on a fast connection.

**Cause:** the boot handler was `async` and awaited SharePoint calls before the
first render. One slow or hung request meant nothing ever painted.

**Fix:** render synchronously from whatever local state you have, *then* upgrade to
SharePoint data in the background. Never block first paint on a network call that
can hang. Structure the app to boot in one of two modes — SharePoint, or a local
fallback — so it is testable and usable either way.

**The general lesson:** in this environment, "slow network" is the normal case, not
the edge case. Design the first paint to be free.

---

## 4. A crash in a shared load path, hidden by catch-and-fallback

**Symptom:** "my data disappeared on refresh." Intermittent-looking. Users
reported it inconsistently, which made it look like a caching or timing problem.

**Cause:** the data-loading function called a normalizer that read `db.user.id`,
which was not yet set at that point in the call chain. It threw on **every single
load**. A broad `try/catch` around the load caught it and substituted an empty
database.

**Why it took so long:** a deterministic bug wearing an intermittent costume. The
fallback was doing its job so well that nothing indicated a failure had occurred.

**Fix, two parts:**
- Audit what a shared function assumes is already set before it runs.
- **Never let a broad catch-and-fallback swallow the error silently.** If you fall
  back, log loudly and surface it in a diagnostics view. A fallback that hides its
  own trigger converts a five-minute bug into a five-day bug.

---

## 5. Fire-and-forget saves made `await` a lie

**Symptom:** child records written with a null or zero foreign key, sometimes.

**Cause:** `Repo.save()` returned as soon as the debounce timer was set, not when
the write landed. So `await Repo.save(parent)` resolved immediately, and the very
next line — creating children referencing `parent.Id` — ran before the server had
assigned an Id.

**Why it took so long:** the code looked correct. There *was* an `await`. The
sequencing was explicit. The bug was in what the promise meant.

**Fix:** a debounced save must return a promise that resolves only when the real
write completes, and writes for the same object must be serialized so a create
always finishes before a follow-up update. `SPList.createSaver()` in
[sp-list-crud.js](../connectors/sp-list-crud.js) is that pattern.

**The general lesson:** if a function returns a promise, be precise about what
resolution *means*. "Queued" and "durable" are different events, and callers will
assume the second.

---

## 6. A debug aid that wiped the live app

**Symptom:** the app occasionally reverted to its boot/startup screen mid-session,
losing whatever the user was doing.

**Cause:** a boot-logging helper rebuilt a "startup console" by clearing the app
root element. Harmless during boot. It had no guard against being called after
boot finished — so any post-boot log (a manual "reload data", a background
refresh) wiped the rendered app back to the boot screen.

**Fix:** any diagnostic or logging utility that touches the DOM needs an explicit
"has boot finished?" guard, built in from the start rather than added after the
first incident.

**The general lesson:** debug instrumentation runs in production. Give it the same
scrutiny as a feature — it has the same power to break things and less of your
attention.

---

## 7. A slow, blocking user-directory sync

**Symptom:** the app felt slow to open, for everyone, always.

**Cause:** every boot synced every SharePoint site user into the app-roles list —
a full site-users fetch plus a create or update per changed user — *before* the
user could see any of their own data. Almost nobody needed that done
synchronously.

**Fix:** run it in the background after the real data has rendered, and apply its
result to the live UI only if it actually changes something for the current user
(their own role). Better still, make it an explicit admin action.

**The general lesson:** ask what the *current user* needs before their first
paint. Anything that serves the system rather than the person in front of you
belongs after the render.

---

## 8. The wrong packaging target

**Symptom:** the app loaded in the web part but behaved noticeably worse than the
reference shipped file — layout and script-execution oddities.

**Cause:** the build used an offline/portable packaging path — a heavy parent shell
that base64-embeds the whole child app and injects it into an `<iframe srcdoc>` —
for what was actually a SharePoint/Firepit ship. It produced a loadable file, so
it looked like it worked.

**Fix:** ship the flat single-file format with everything inlined (`notes/03`).
That is what the platform's own SharePoint path produces and what behaves
predictably in a web part. If your tooling offers a distinctly-named
SharePoint/Flank-Speed export, use that one rather than the first export you find.

Worth knowing: the host **sanitizes nested `srcdoc` attributes** (`notes/02` §5),
so the wrapper approach is fighting the platform, not merely heavier than it.

---

## 9. Two that will bite you and were not our bugs

Not our defects, but the same class of problem — a symptom that points at your
code when the cause is the platform.

**`Identifier 'X' has already been declared`, for every top-level `const`/`let`.**
Microsoft Defender for Cloud Apps session-proxies the page, wraps DOM APIs, and
executes your scripts twice. Firepit works around it by loading your app from a
`blob:` URL and never touching `srcdoc` or `innerHTML` for the iframe. If you see
this, you are looking at MCAS, not a duplicate script tag. (`notes/02` §3 — the
workaround is in Firepit's own source comments.)

**A `fetch` that fails with an opaque `TypeError` and no status.** Almost always
the host's injected `connect-src` CSP blocking the origin, not a bad key or a bad
URL. Check the console for a CSP violation before debugging anything else.
(`notes/02` §4.)

---

## The pattern across all of these

Six of the nine were **silent failures**: a 201 that discarded data, a 404 with no
log, a catch that substituted empty state, a promise that resolved early, a
truncation with only a `console.warn`, a query string that failed before any code
ran.

The habit that would have caught most of them: **make the failure loud at the
boundary.** Verify after you create. Log when you fall back. Resolve promises only
on durable success. Throw on the external reference at build time rather than
discovering it as a CSP violation in production. Every one of those is a few lines
of code that trades a quiet wrong answer for a noisy correct one.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
