// Bundle index.html, styles.css and the scripts into one self-contained HTML file.
// Usage: node build-single-file.mjs <output.html> [--fragment]
// --fragment drops the <!doctype>/<html>/<head>/<body> wrapper for hosts that add their own.
import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const sourceDirectory = dirname(fileURLToPath(import.meta.url));
const [outputPath, mode] = process.argv.slice(2);
if (!outputPath) {
  console.error("Usage: node build-single-file.mjs <output.html> [--fragment]");
  process.exit(1);
}

const read = (name) => readFileSync(join(sourceDirectory, name), "utf8");
let html = read("index.html")
  .replace('<link rel="stylesheet" href="styles.css" />', () => `<style>\n${read("styles.css")}</style>`)
  .replace('<script src="cleaner.js"></script>', () => `<script>\n${read("cleaner.js")}</script>`)
  .replace('<script src="app.js"></script>', () => `<script>\n${read("app.js")}</script>`);

if (mode === "--fragment") {
  html = html
    .replace(/<!doctype html>\s*/i, "")
    .replace(/<\/?(html|head|body)[^>]*>\s*/gi, "")
    .replace(/<meta charset="utf-8" \/>\s*/, "")
    .replace(/<meta name="viewport"[^>]*\/>\s*/, "");
}

writeFileSync(outputPath, html);
console.log(`Wrote ${outputPath}`);
