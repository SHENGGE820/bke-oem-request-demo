import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = resolve(import.meta.dirname);
const pageFiles = ["preview.html", "wizard.html", "case.html", "tasks.html", "reports.html"];
const pages = {};

for (const file of pageFiles) {
  pages[`/${file}`] = await readFile(resolve(root, file), "utf8");
}
pages["/"] = pages["/preview.html"];

const worker = `const pages=${JSON.stringify(pages)};
export default {
  async fetch(request) {
    const url = new URL(request.url);
    const html = pages[url.pathname];
    if (!html) return new Response("Not found", { status: 404 });
    return new Response(html, {
      headers: {
        "content-type": "text/html; charset=utf-8",
        "cache-control": "no-store",
        "x-content-type-options": "nosniff"
      }
    });
  }
};
`;

await rm(resolve(root, "dist"), { recursive: true, force: true });
await mkdir(resolve(root, "dist/server"), { recursive: true });
await writeFile(resolve(root, "dist/server/index.js"), worker);
