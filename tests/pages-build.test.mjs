import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("builds the GitHub Pages application under the repository base path", async () => {
  const html = await readFile(path.join(root, "dist-pages", "index.html"), "utf8");
  assert.match(html, /<title>Numbi EEFF \| Estados financieros desde tu balance<\/title>/i);
  assert.match(html, /\/EEFF-Numbi\/assets\/index-[^"']+\.js/i);
  assert.match(html, /\/EEFF-Numbi\/numbi-assistant\.png/i);
  assert.doesNotMatch(html, /<h1[^>]*>EEFF-Numbi<\/h1>/i);
  await access(path.join(root, "dist-pages", "numbi-upload.png"));
});
