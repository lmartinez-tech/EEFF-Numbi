import assert from "node:assert/strict";
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

test("renders the Numbi EEFF application", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<html lang="es">/i);
  assert.match(html, /<title>Numbi EEFF \| Estados financieros desde tu balance<\/title>/i);
  assert.match(html, /Tus n(?:ú|&#xFA;|&uacute;)meros/i);
  assert.match(html, /Sube tus balances de prueba por tercero/i);
  assert.match(html, /Seleccionar archivos/i);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/i);
});
