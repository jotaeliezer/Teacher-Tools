import { cpSync, existsSync, mkdirSync } from "node:fs";
import { resolve } from "node:path";

const root = resolve(process.cwd());
const docsDir = resolve(root, "docs");

const publicDir = resolve(root, "public");
const docsMappings = [
  ["teacher_tools.html", "docs/teacher_tools.html"],
  ["styles/teacher_tools.css", "docs/styles/teacher_tools.css"],
  ["scripts", "docs/scripts"],
  ["icon", "docs/icon"]
];

// Step 1: Sync root → public (keeps public in sync with source)
if (existsSync(publicDir)) {
  for (const [src, dest] of docsMappings) {
    const destPath = dest.replace("docs/", "public/");
    const destFull = resolve(root, destPath);
    const destParent = resolve(destFull, "..");
    if (!existsSync(destParent)) mkdirSync(destParent, { recursive: true });
    cpSync(resolve(root, src), destFull, { recursive: true, force: true });
  }
}

// Step 2: Sync public → docs (or root → docs if no public)
if (!existsSync(docsDir)) {
  mkdirSync(docsDir, { recursive: true });
}

const sourceDir = existsSync(publicDir) ? publicDir : root;
const sourceMappings = [
  ["teacher_tools.html", "docs/teacher_tools.html"],
  ["styles/teacher_tools.css", "docs/styles/teacher_tools.css"],
  ["scripts", "docs/scripts"],
  ["icon", "docs/icon"]
];

for (const [src, dest] of sourceMappings) {
  const srcPath = resolve(sourceDir, src);
  const destPath = resolve(root, dest);
  if (existsSync(srcPath)) {
    cpSync(srcPath, destPath, { recursive: true, force: true });
  }
}

console.log("Synced docs assets for GitHub Pages.");
