import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = resolve(import.meta.dirname);
const pageFiles = ["preview.html", "role-portal.html", "sales.html", "form-versions.html", "wizard.html", "wizard-simple.html", "product-data-form.html", "marketing-request-form.html", "nutritionist.html", "rd.html", "case.html", "tasks.html", "reports.html", "shared-case-store.js"];
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
        "content-type": url.pathname.endsWith(".js") ? "text/javascript; charset=utf-8" : "text/html; charset=utf-8",
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
