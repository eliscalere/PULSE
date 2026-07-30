<!-- UNCLASSIFIED -->

# Firepit web part — the host, unpacked

**UNCLASSIFIED**

The SPFx web part that renders your app. Source extracted from
[`../forge/firepit-src-7.1.0.zip`](../forge/); manifests extracted from
[`../forge/firepit-prod.sppkg`](../forge/).

**Read [../notes/02-firepit-webpart-internals.md](../notes/02-firepit-webpart-internals.md)
first.** It is the analysis of this code — the CSP it injects, the rendering path,
the postMessage protocol, and the eight things to check before blaming your own
app. This README is just the map.

## Layout

```
firepit-webpart/
├── src/
│   ├── index.ts
│   └── webparts/htmlFileRenderer/
│       ├── HtmlFileRendererWebPart.ts             1,750 lines — everything
│       ├── HtmlFileRendererWebPart.manifest.json  web part id, hosts, toolbox entry
│       ├── HtmlFileRendererWebPart.module.scss    container and iframe styling
│       ├── HtmlFileRendererWebPart.module.scss.ts generated style typings
│       ├── assets/                                welcome-light.png, welcome-dark.png
│       └── loc/                                   en-us strings
└── sppkg-manifests/
    ├── AppManifest.xml          solution identity + API permission requests
    ├── ClientSideAssets.xml     bundle inventory
    └── [Content_Types].xml
```

Not extracted: the 60 pre-built JS bundles under `ClientSideAssets/` in the
`.sppkg`. They are webpack output of the source in `src/` and add nothing readable.

## Identity

| | |
|---|---|
| Solution | `firepit-client-side-solution` **7.1.0.0** |
| Product ID | `6dd7e353-52da-4020-8eaf-4ff0de1ec35a` |
| Web part | `HtmlFileRendererWebPart` · `4b2ad7ba-a39b-49e8-90e3-10e08433578b` |
| Shows up as | **Firepit**, in the *Advanced* toolbox group |
| `requiresCustomScript` | `false` — installs where Custom Script is disabled |
| `SkipFeatureDeployment` | `true` — tenant-wide deployable |
| Hosts | SharePointWebPart · SharePointFullPage · TeamsTab · TeamsPersonalApp |
| API permission requested | **Microsoft Dataverse**, `user_impersonation` |

That last row is the reason the Dataverse proxy exists: the tenant grants the
scope to the *solution*, the host holds the token, and your sandboxed frame asks
the host to make calls on its behalf. Your app never sees a token.

## Where to look in the 1,750 lines

| Lines | What |
|---|---|
| 14–24 | `IHtmlFileRendererWebPartProps` — the full property schema |
| 62–71 | Chunking limits and the four postMessage type constants |
| 79–152 | Message handlers: edit-panel hotkey, Dataverse request, destructive-command approval |
| 155–260 | `render()` — error state, then the blob-URL iframe construction |
| 227–240 | **The MCAS comment.** Why it is a blob URL and not `srcdoc`. Read this one |
| 349–412 | `_loadHtmlFile()` — fetching your HTML from SharePoint |
| 414–490 | Chunked inline storage: detect, reassemble, save, clear |
| 845–950 | Security wrapper construction: CSP nonce, injected metas, nonce application |
| 892–920 | **The CSP.** `connect-src`, `frame-src`, and the full directive list |
| ~1040–1150 | Injected child-side scripts: destructive-command approval flow |
| ~1330–1455 | Injected DOM guards: attribute sanitizer, `srcdoc` sanitizer, allow-lists |
| 1681–1690 | `_getSandboxAttribute()` — the iframe sandbox token list |

## The three facts that will change how you build

1. **Your app loads from a `blob:` URL inside a sandboxed iframe.** Not from the
   SharePoint origin, and not via `srcdoc`. Both choices are deliberate
   workarounds for MCAS session-proxy double-execution — the symptom being
   `Identifier 'X' has already been declared` on every top-level declaration.

2. **The host injects `default-src 'none'` into your document.** Every "why can't
   I load a CDN / a font / submit a form / go fullscreen" question in this stack
   is answered by the directive list at line 902. Inline everything.

3. **`lockDown` sets `connect-src 'none'`.** A locked-down web part cannot reach
   SharePoint at all. If your app boots to "cannot reach SharePoint" and the code
   looks right, check this toggle before debugging anything else.

## Provenance and scope

Read from Firepit **7.1.0** on 2026-07-29. `firepit_7.0.7.0.sppkg` is also in
[`../forge/`](../forge/) if you need to diff against the previous release.

This is a **read-only reference**, checked in so app developers can see what the
host does to their code. It is not a fork and not something to build from — Firepit
belongs to its maintainers, and anything you need changed in it is a request to
them, not a patch here.

**UNCLASSIFIED**

<!-- UNCLASSIFIED -->
