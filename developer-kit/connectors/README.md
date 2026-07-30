<!-- UNCLASSIFIED -->

# Connectors

**UNCLASSIFIED**

Working, dependency-free integration code. Read these instead of the prose if you
learn better from source — every one of them is commented with *why*, not just
what, and every warning in the comments is something that actually went wrong in
this program.

None of them need a build step, a package manager, or a framework. Paste a file
into your app, or `<script src>` it during development and let your packaging step
inline it.

## Load order

`sp-rest-core.js` first — everything else builds on it.

```html
<script src="connectors/sp-rest-core.js"></script>
<script src="connectors/sp-columns.js"></script>
<script src="connectors/sp-list-crud.js"></script>
<script src="connectors/sp-json-store.js"></script>
<script src="connectors/sp-files.js"></script>
<script src="connectors/sp-email.js"></script>
<script src="connectors/sp-identity-roles.js"></script>
<script src="connectors/ai-openai-compatible.js"></script>
<script src="connectors/firepit-bridge.js"></script>
```

## What each one is for

| File | Use it for | The trap it documents |
|---|---|---|
| [sp-rest-core.js](sp-rest-core.js) | Site detection, request digest, the single fetch wrapper, paging | Digest caching and the 403-means-stale-digest retry |
| [sp-columns.js](sp-columns.js) | Creating lists and columns | **`Options: 28`.** Read this before creating one column |
| [sp-list-crud.js](sp-list-crud.js) | Items: read, create, update, delete | Field types that need `{results:[…]}` or an `…Id` property; debounced serialized saves |
| [sp-json-store.js](sp-json-store.js) | A document library as a database | What you give up versus lists, stated plainly |
| [sp-files.js](sp-files.js) | Folders, binary upload, ETags, parsing a pasted SP URL | Binary through a text path corrupts silently and returns 200 |
| [sp-email.js](sp-email.js) | Mail with no server and no Graph | Requires `odata=verbose`; failure must be non-fatal |
| [sp-identity-roles.js](sp-identity-roles.js) | Who the user is, what they may do | `IsSiteAdmin` is not your Admin role |
| [ai-openai-compatible.js](ai-openai-compatible.js) | LLM calls from a browser-only app | There is nowhere to hide an API key; and `connect-src` will block you |
| [firepit-bridge.js](firepit-bridge.js) | Talking to the host web part | Dataverse proxy, CSP nonce, destructive-op throttle |

## Try them

[demo.html](demo.html) is a single self-contained page that exercises every
connector against a live site and reports what worked. Upload it to a document
library and open it from SharePoint — **not** from your desktop, since session
auth needs the SharePoint origin.

It is read-only by default. Write tests are behind an explicit checkbox and
clearly labelled, because a demo page that quietly creates lists on someone's
site is not a demo, it is an incident.

## The five things worth knowing before you write any of this yourself

1. **`credentials: "same-origin"` is the whole auth story.** No app registration,
   no MSAL, no tokens.
2. **`POST /_api/contextinfo` for a digest before any write.** Cache it; it lives
   ~30 minutes.
3. **`Accept: application/json;odata=nometadata`** everywhere — except
   `createfieldasxml` and `SendEmail`, which demand `odata=verbose`.
4. **Follow `odata.nextLink`.** Lists page at ~100 items and the failure mode is
   silently partial data.
5. **`Options: 28` on every column you create.** See `sp-columns.js`. This one
   changed an architecture here.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
