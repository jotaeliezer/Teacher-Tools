import { cpSync, existsSync, mkdirSync } from "node:fs";
import { resolve } from "node:path";

const root = resolve(process.cwd());
const docsDir = resolve(root, "docs");

const mappings = [
  ["teacher_tools.html", "docs/teacher_tools.html"],
  ["styles/teacher_tools.css", "docs/styles/teacher_tools.css"],
  ["scripts", "docs/scripts"],
  ["icon", "docs/icon"]
];

if (!existsSync(docsDir)) {
  mkdirSync(docsDir, { recursive: true });
}

for (const [src, dest] of mappings) {
  cpSync(resolve(root, src), resolve(root, dest), {
    recursive: true,
    force: true
  });
}

console.log("Synced docs assets for GitHub Pages.");
