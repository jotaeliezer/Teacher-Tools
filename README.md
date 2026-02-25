# teacher_tools_file
tools site for teachers

## Build Pipeline

1. Install dependencies (first time only):

```bash
npm install
```

2. Build the app:

```bash
npm run build
```

3. Sync the deployable `/docs` copy for GitHub Pages:

```bash
npm run sync:docs
```

This copies `teacher_tools.html`, `styles/`, `scripts/`, and `icon/` from the project root into `docs/`. It also syncs root → `public/` first (if present), so `public/` stays in sync.

4. Commit and push (`dist/` is local build output, `docs/` is what Pages serves).

## GitHub Pages (`/docs`)

- Set GitHub Pages source to `/docs`.
- Entry point is `docs/index.html`, which redirects to `docs/teacher_tools.html`.
- Root files remain the editable source; `docs/` is the publish copy.

After editing root teacher tools files, run the full pipeline:

```bash
npm run build
npm run sync:docs
```
