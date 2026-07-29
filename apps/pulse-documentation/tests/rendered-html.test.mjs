import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the PULSE documentation loader", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>PULSE Documentation<\/title>/i);
  assert.match(html, /class="pulse-doc-loader"/);
  assert.match(html, /role="status"/);
  assert.match(html, /aria-label="Loading PULSE Documentation"/);
  assert.match(html, /pulse-svgl-dot pulse-svgl-dot--upper/);
  assert.match(html, /pulse-svgl-dot pulse-svgl-dot--lower/);
  assert.match(html, /Loading the PULSE documentation library/);
});

test("keeps the loader purposeful and motion-accessible", async () => {
  const [page, css] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);

  assert.match(page, /Promise\.all\(documents\.map/);
  assert.match(page, /setBootPhase\("exiting"\)/);
  assert.match(page, /setBootPhase\("ready"\)/);

  assert.match(css, /\.pulse-doc-loader\s*\{/);
  assert.match(css, /@keyframes pulse-svgl-dot-upper/);
  assert.match(css, /@keyframes pulse-svgl-dot-lower/);
  assert.match(css, /@keyframes pulse-svgl-reveal/);
  assert.match(css, /@media \(prefers-reduced-motion:reduce\)/);
  assert.doesNotMatch(css, /bounce|elastic/i);
});
