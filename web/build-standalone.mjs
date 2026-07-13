/*
 * build-standalone.mjs — bundle the site into one self-contained HTML file.
 *
 * Inlines styles.css and concatenates calcs.js + app.js (stripping ES-module
 * syntax) into a single classic <script>, so the result opens by double-click
 * over file:// with no server. Output: standalone/div296-reset-calculator.html
 *
 *   node web/build-standalone.mjs
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, f), "utf8");

const css = read("styles.css");
const calcs = read("calcs.js").replace(/^export\s+/gm, "");
const app = read("app.js").replace(
  /import\s*\{[\s\S]*?\}\s*from\s*["']\.\/calcs\.js["'];\s*/,
  ""
);

const bundle = `// ==== calcs.js (engine) ====\n${calcs}\n\n// ==== app.js (ui) ====\n${app}`;

// NB: function replacements, not string replacements. A string replacement
// runs JS's $-substitution ($$ -> $, $& -> match, ...), which silently ate a
// dollar sign from `$${t1...}` in app.js and shipped the standalone with
// "exceeds 3,000,000" (no $). A function replacement inserts the text verbatim.
let html = read("index.html")
  .replace(
    /<link rel="stylesheet" href="styles\.css" \/>/,
    () => `<style>\n${css}\n</style>`
  )
  .replace(
    /<script type="module" src="app\.js"><\/script>/,
    () => `<script>\n${bundle}\n</script>`
  );

if (/\bexport\s/.test(bundle) || /\bimport\s/.test(bundle)) {
  throw new Error("module syntax leaked into the bundle");
}

mkdirSync(join(here, "standalone"), { recursive: true });
const out = join(here, "standalone", "div296-reset-calculator.html");
writeFileSync(out, html);
console.log(`wrote ${out} (${html.length} bytes)`);
