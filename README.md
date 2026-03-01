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

## How Teacher Tools Works

### Core navigation

- **Files drawer (left):** upload multiple CSV/XLSX class files, switch active class, rename, or remove loaded files.
- **Data:** clean and review class data (show/hide columns, show/hide rows, transpose, filter, inline edits, export).
- **Print View:** pick one template at a time (attendance / marking / drill / report card), preview WYSIWYG pages, then print.
- **Phone Logs:** track family contact entries per student and print a phone log sheet.
- **Report Comments:** build deterministic report card comments from performance level, assignment marks, and comment bank entries; optional AI revise/create actions run through the API endpoint.
- **Seating Plan:** define room size, edit student traits, generate AI seating suggestions, then manually drag/swap and lock seats.

### Typical teacher workflow

1. Load one or more class files in the Files drawer.
2. Select a class and quickly verify columns/rows in **Data**.
3. Use **Report Comments** or **Seating Plan** for planning/writing tasks.
4. Open **Print View** to print attendance/marking/drill/report sheets.
5. Use **Phone Logs** for parent-contact records and printing.

### Persistence model

- UI preferences (dark mode, zoom, selected print template, API URL) are saved in browser localStorage.
- Saved comments are stored locally in the same browser.
- Seating plans are stored per loaded class context in localStorage.
- No server database is used by default.

## GitHub Pages (`/docs`)

- Set GitHub Pages source to `/docs`.
- Entry point is `docs/index.html`, which redirects to `docs/teacher_tools.html`.
- Root files remain the editable source; `docs/` is the publish copy.

After editing root teacher tools files, run the full pipeline:

```bash
npm run build
npm run sync:docs
```
