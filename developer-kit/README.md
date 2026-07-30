<!-- UNCLASSIFIED -->

# Developer Kit — building your own apps for Firepit + SharePoint Online

**UNCLASSIFIED**

Everything in this folder exists to answer one question: *if I wanted to build my
own app that runs inside a Firepit web part on SharePoint Online, with SharePoint
as both the login and the database, what would I need to know?*

It is written for a developer who has never seen this stack. Nothing here is
theory — every claim traces to a shipped artifact in this repo, and where
something is inferred rather than observed, it says so.

---

## Start here

| Read this | If you want to |
|---|---|
| [notes/01-platform-model.md](notes/01-platform-model.md) | Understand the whole stack in ten minutes |
| [notes/02-firepit-webpart-internals.md](notes/02-firepit-webpart-internals.md) | Know what the host actually does to your HTML |
| [notes/03-wfc-package-format.md](notes/03-wfc-package-format.md) | Ship a single-file app |
| [notes/04-decompiling-a-shipped-app.md](notes/04-decompiling-a-shipped-app.md) | Recover source from a compiled `.html` |
| [notes/05-sharepoint-backend-recipes.md](notes/05-sharepoint-backend-recipes.md) | Read and write data, files, mail, roles |
| [notes/06-postmortems.md](notes/06-postmortems.md) | Avoid the bugs that already cost us days |
| [connectors/](connectors/) | Copy working code instead of reading prose |

If you only read one thing, read `notes/02`. The host's injected
Content-Security-Policy is the single constraint that shapes every other
decision in this stack, and it is not documented anywhere else.

---

## What's in here

```
developer-kit/
├── forge/                      Source artifacts (provided as-is)
│   ├── Forge.html                     The compiled Forge IDE — 8.2 MB, one file
│   ├── firepit-src-7.1.0.zip          Firepit SPFx web part source
│   ├── firepit-prod.sppkg             Firepit deployment package (7.1.0)
│   ├── firepit_7.0.7.0.sppkg          Previous Firepit release
│   └── Firepit_and_Forge_Architecture_Guide.docx
│
├── forge-ide-decompiled/       ★ Forge.html unpacked back into source
│                                 34/34 manifest entries hash-verified.
│                                 25 of Forge's own modules, readable.
│
├── firepit-webpart/            The host, unpacked
│   ├── src/                           SPFx TypeScript (1,750-line web part)
│   └── sppkg-manifests/               Package manifests from the .sppkg
│
├── notes/                      Developer notes (six documents)
│
├── connectors/                 Runnable, dependency-free integration examples
│   ├── sp-rest-core.js                Session auth, digest, request plumbing
│   ├── sp-list-crud.js                Lists, items, paging
│   ├── sp-columns.js                  Column creation (the part that breaks everyone)
│   ├── sp-json-store.js               Document-library-as-database
│   ├── sp-files.js                    Binary upload, folders, ETags
│   ├── sp-email.js                    Mail with no server and no Graph
│   ├── sp-identity-roles.js           Who is this user, what may they do
│   ├── ai-openai-compatible.js        LLM calls from a browser-only app
│   ├── firepit-bridge.js              Talking to the host web part
│   └── demo.html                      Exercise all of the above in a browser
│
├── examples/
│   └── pulse-code-ide/         A working single-file IDE, split into 16 readable
│                                 modules. Byte-verified against its original.
│
└── tools/
    ├── wfc-decompile.js        Unpack any shipped package back into source
    └── prepare-forge-tree.js   Regenerate forge-ide-decompiled/ as committed
```

★ = the thing most people came here for.

---

## The five-minute version

1. **A Firepit web part renders one blob of HTML inside a sandboxed iframe.**
   You give it a single self-contained `.html` file. Inline scripts run, inline
   styles apply.

2. **Ship one file with everything inlined.** No external `<script src>`, no
   CDN, no fonts from the network. The host injects `default-src 'none'` into
   your frame, so anything you did not inline simply will not load.

3. **SharePoint is your backend, and you are already logged in.** The page is
   served from the SharePoint origin, so the browser carries the user's session.
   `fetch('/_api/web/currentuser', {credentials: 'same-origin'})` works with no
   Azure AD app registration, no Graph, no tokens to manage.

4. **Writes need a request digest.** `POST /_api/contextinfo` first, then pass
   `X-RequestDigest` on every POST/MERGE/DELETE. Reads need nothing.

5. **Render before you fetch.** Paint synchronously from local state, then
   upgrade to SharePoint data in the background. A blocking `await` on boot is
   the difference between a slow app and a blank page — see `notes/06`.

Working code for all of it: [connectors/](connectors/).

---

## Reproducing the decompiled tree

The interesting artifact in this kit was produced by a tool that is also in this
kit, from an input that is also in this kit. Nothing is hand-massaged:

```bash
node developer-kit/tools/wfc-decompile.js developer-kit/forge/Forge.html /tmp/forge-full
```

That writes 43 files and reports how each one was named — hash-verified, or
inferred, and from what evidence. The committed
[forge-ide-decompiled/](forge-ide-decompiled/) is the same output with the public
CDN bundles and the redundant copies of the input removed; see its README.

---

## Ground rules for anything you read here

- **Verified** means it was observed in a shipped artifact in this repo, and the
  document says which one.
- **Inferred** is labelled inferred. Forge and Firepit are internal platforms;
  their source of truth is their maintainers, not this folder.
- Version-pinned claims name the version. Firepit is read from **7.1.0**; the
  Forge build decompiled here was compiled **2026-06-29**.

---

## Contact

> **Developer:** Eli Scalere (Contractor)
> **Program:** AEWTTR IPT · DB46200 · NAVAIR
> **Program contact:** elijah.t.scalere.ctr@us.navy.mil
> **Out-of-service:** eli.scalere@scaleredesign.com

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
