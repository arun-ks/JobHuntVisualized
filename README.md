# Job hunt report

This repository builds a **single static HTML file** that visualizes job-application progress from an Excel macro-enabled workbook (**`JobHunt.xlsm`**). There is no server: the HTML loads Plotly from a CDN and reads data that was **embedded at build time**.

## Repository layout

| Path | Purpose |
|------|---------|
| **`JobHunt.xlsm`** (repo root) | Your tracker workbook. Sheet name must be **`JobHunt`**. |
| **`JobHunt/`** (Node project folder) | Build script and HTML template. |
| **`JobHunt/build.mjs`** | Reads the workbook, normalizes rows to JSON, writes **`dist/jobhunt-progress.html`**. |
| **`JobHunt/report-template.html`** | UI, filters, and Plotly charts. Placeholder `__APP_B64__` is replaced during the build. |
| **`excel-vba/`** | Optional VBA: **`Mod_AddApplication.bas`**, **`clsJobHuntFormButtons.cls`**, form paste text. Import into `JobHunt.xlsm`; see **Excel VBA (quick-add)** below. |
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
node build.mjs "D:\path\JobHunt.xlsm"
```

## Expected workbook columns

Headers are matched **by name** (not by column letter), so you can reorder columns as long as names stay consistent. Important names include:

- **Company Name**, **Position Name**, **Role Type**, **Status**, **Log**, **Portal**, **URL of Job Posting** (and variants with trailing spaces are tolerated).
- **Date-** milestone columns: **Job Posted**, **Applied**, **Short Listed**, **Interview/Test**, **Offer**, **Rejected**, **Declined by Me**.
- **Flag to see if the job posting is still accepting candidates** — read into data; charts do not depend on it.

**Status** should be stored as a **cached value** in Excel (save after formulas calculate) so the export shows the intended text.

## Viewing the report

Open **`dist/jobhunt-progress.html`** in a modern browser. **Internet access** is required for the Plotly CDN unless you vendor the library locally later.

Peers can receive **only the HTML file** after you run a build; no Node is needed on their side.

## Excel VBA (quick-add)

Optional macros live under **`excel-vba/`**:

1. In Excel, open **`JobHunt.xlsm`**, press **Alt+F11**.
2. **File → Import File…** and import **`excel-vba/Mod_AddApplication.bas`** and **`excel-vba/clsJobHuntFormButtons.cls`** (re-import after updates so **`AppendJobHuntApplication`** and the button hook stay in sync).
3. **Insert → UserForm**. Set its **Name** to **`frmAddApplication`** in the Properties window. Leave the form surface **empty** (no toolbox controls). If you still have old **cmdOK** / **cmdCancel** from the toolbox, delete them so the code can create those controls without a name clash.
4. Double-click **`frmAddApplication`**, open **`excel-vba/frmAddApplication_PASTE_INTO_USERFORM.txt`**, copy **all** of it, and paste into the UserForm code window (replace everything there).
5. **Save** the workbook as macro-enabled (`.xlsm`). Close and reopen once so **`Auto_Open`** runs, or in the Immediate window (**Ctrl+G**) run: **`RegisterJobHuntHotkey`**

**Buttons:** **Add row** and **Cancel** are created in **`BuildControls`**. **`clsJobHuntFormButtons`** uses **`WithEvents`** on those `CommandButton` references so clicks are handled even though the buttons are not drawn on the designer (run-time `Controls.Add` does not wire `cmdOK_Click` on the UserForm module).

**Hotkey:** **Ctrl+Shift+N** opens the dialog. Change **`HOTKEY_SEQUENCE`** in `Mod_AddApplication.bas` if that shortcut clashes with another add-in.

**Role type** is a **plain single-select list** (not checkboxes). Excel often returns **error 380** if `ListStyle = 2` (checkbox) is applied to a **run-time** `ListBox`; plain list + `ListIndex` avoids that. Click one row before **Add row**.

Labels, text boxes, and the role list are created in code, so they are **not** real members of the UserForm (e.g. `Me.txtJobPosted` will fail). The pasted code uses **`Tb("…")`** and **`LstRoleCtl()`** helpers that read **`Me.Controls("name")`** instead.

**New rows:** **`AppendJobHuntApplication`** copies **number / font / fill / border** formatting from **`A2:O2`** onto the new row and matches **row height** to row **2**.

If column **C** has no values yet, add one manual row first so the list is not empty.

## Maintenance notes

- The HTML report uses a **fixed light** theme (no dark mode toggle).
- Chart logic and filter behaviour live in **`report-template.html`** (inline script). Build logic and column mapping live in **`build.mjs`** — update both if you add columns that should appear in charts or filters.
