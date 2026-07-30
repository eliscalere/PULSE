<!-- UNCLASSIFIED -->
# _vendor-cdn — extracted, then deliberately not committed

Forge fetches these at compile time and inlines them into the shipped file,
recording them in the manifest with `external: true`. `wfc-decompile.js` writes
them here, but they are unmodified public library releases, so the bytes are not
kept in git. What Forge shipped on 2026-06-29T17:03:54.348Z:

| kind | URL |
|---|---|
| css | https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css |
| js | https://code.jquery.com/jquery-3.7.1.min.js |
| js | https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js |
| js | https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.0/FileSaver.min.js |
| js | https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js |
| js | https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.8.0/mammoth.browser.min.js |
| js | https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js |

Two things worth noticing:

1. Forge inlines CDN dependencies rather than rejecting them. The PULSE packagers
   throw on any `https://` reference instead. Both end up dependency-free at
   runtime; Forge just does the fetching for you at compile time.
2. It has to. The Firepit host injects `default-src 'none'` into the child frame,
   so a runtime CDN fetch would be blocked outright. See
   `../../notes/02-firepit-webpart-internals.md`.

Regenerate the full tree, vendor bundles included:

```bash
node developer-kit/tools/wfc-decompile.js developer-kit/forge/Forge.html /tmp/forge-full
```

<!-- UNCLASSIFIED -->
