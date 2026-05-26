# Job hunt Visualized

This repository builds a **single static HTML file** that visualizes job-application progress from an Excel macro-enabled workbook (**`JobHunt.xlsm`**).

There is no server: the HTML loads Plotly from a CDN and reads data that was **embedded at build time**.

Sample [report can be seen here](https://arun-ks.github.io/JobHuntVisualized/dist/index.html) which provides multiple views & filtering options.

You just need to maintain the DB in the excel & let this took take care of the reporting.

## Repository layout

| Path | Purpose |
|------|---------|
| **`JobHunt.xlsm`** (repo root) | Your tracker workbook. Sheet name must be **`JobHunt`**. |
| **`JobHunt/`** (Node project folder) | Build script and HTML template. |
| **`JobHunt/build.mjs`** | Reads the workbook, normalizes rows to JSON, writes **`dist/jobhunt-progress.html`**. |
| **`JobHunt/report-template.html`** | UI, filters, and Plotly charts. Placeholder `__APP_B64__` is replaced during the build. |
| **`dist/jobhunt-progress.html`** | **Generated output** — run the build to create or refresh it. |

## Prerequisites

- **Node.js** 18+ (includes `npm`).

## How the build works

1. **`build.mjs`** resolves the input file: path passed on the command line, else **`../JobHunt.xlsm`** (repo root, next to the `JobHunt/` folder).
2. The workbook is read **as a buffer** and parsed with **SheetJS (`xlsx`)** (more reliable on Windows than reading the path directly).
3. The worksheet named **`JobHunt`** is used if it exists; otherwise the **first sheet** in the file is used.
4. Row **1** is treated as **headers**; each following row becomes an object keyed by the normalized header string.
5. Each row is mapped to an **application** object (company, position, role type, status, dates, ordered `steps`, etc.). Rows without **`Date- Applied`** are dropped.
6. A JSON payload `{ generatedAt, source, sheet, applications }` is **Base64-encoded** and injected into the template (avoids broken HTML if text contained `</script>`-like sequences).
7. The result is written to **`dist/jobhunt-progress.html`**.

## Commands

```bash
cd JobHunt
npm install
npm run build
```

- **`npm run build:open`** — builds then tries to open the HTML on Windows (`start ..\dist\jobhunt-progress.html`).

Custom workbook path:

```bash
node build.mjs "C:\path\JobHunt.xlsm"
```

## Excel Workbook columns

Headers are matched **by name** (not by column letter), so you can reorder columns as long as names stay consistent. Important names include:

- **Company Name**, **Position Name**, **Role Type**, **Status**, **Log**, **Portal**, **URL of Job Posting** (and variants with trailing spaces are tolerated).
- **Date-** milestone columns: **Job Posted**, **Applied**, **Short Listed**, **Interview/Test**, **Offer**, **Rejected**, **Declined by Me**.
- **Flag to see if the job posting is still accepting candidates** — read into data; charts do not depend on it.

**Status** should be stored as a **cached value** in Excel (save after formulas calculate) so the export shows the intended text.


## Viewing the report

Open **`dist/jobhunt-progress.html`** in a modern browser. **Internet access** is required for the Plotly CDN unless you vendor the library locally later.

Peers can receive **only the HTML file** after you run a build; no Node is needed on their side.

## Excel VBA 

The Workbook **`JobHunt.xlsm`** has easy to use buttons to search jobs (based on Company Name) and to add new row.
Copy the URL to Clipboard, before you press the "add Record" button, as it would copy the URL from clipboard directly into the new record.

**Buttons:** **Add Record** (to add new record in DB), **Search Company** and **Clear Filter**.

**Hotkey:** **Ctrl+Shift+S** opens dialog for "Search Company", it defaults to value in 1st column of current row. **Ctrl+Shift+I** would add the new row.

## Maintenance notes

- The HTML report uses a **fixed light** theme 
- Chart logic and filter behaviour live in **`report-template.html`** (inline script). Build logic and column mapping live in **`build.mjs`** — update both if you add columns that should appear in charts or filters.

## Co-written by AI

The initial nodejs code was co-written with [CursorAI](https://cursor.com/), which also added detailed documentation in the code.

The excel template was created by [Arun-KS](https://github.com/arun-ks/) who also had created the VBA Macros, as Cursor struggled & complicated it.

The MJS code was maintained & enhanced by Arun-KS, to handle some bugs in Cursor's code.

