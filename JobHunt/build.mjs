/**
 * Build pipeline: read the Job Hunt workbook and emit a single static HTML report.
 *
 * Flow:
 * 1. Resolve workbook path (CLI arg, or ../JobHunt.xlsm at repo root).
 * 2. Load workbook bytes with SheetJS; pick sheet "JobHunt" or fall back to first sheet.
 * 3. Convert rows to objects keyed by header names; map each row to a normalized "application".
 * 4. Serialize applications + metadata to JSON, Base64-encode (avoids `</script>` issues in HTML).
 * 5. Inject payload into report-template.html → ../dist/index.html.
 *
 * Usage: `node build.mjs`  or  `node build.mjs "D:\path\JobHunt.xlsm"`
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");
const SHEET_NAME = "JobHunt";
const defaultWorkbook = path.join(repoRoot, "JobHunt.xlsm");
const outDir = path.join(repoRoot, "dist");
const outHtml = path.join(outDir, "index.html");

const argPath = process.argv[2];
const workbookPath = argPath ? path.resolve(argPath) : defaultWorkbook;

/**
 * Normalizes a header or string cell for matching: trim and collapse internal whitespace.
 * @param {unknown} s
 * @returns {string}
 */
function norm(s) {
  return String(s ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

/**
 * Converts an Excel cell value to ISO calendar date `YYYY-MM-DD`, or null if missing/invalid.
 * Handles Date objects, Excel serial numbers (1900-based), and parseable strings.
 * @param {unknown} v
 * @returns {string | null} 
 * ArunS : Renamed this Function, as it assumed dates are in UTC. 
 */
function toISODateAsUTC(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return v.toISOString().slice(0, 10);
  }
  if (typeof v === "number" && Number.isFinite(v)) {
    const epoch = Date.UTC(1899, 11, 30);
    const ms = epoch + Math.round(v * 86400000);
    const d = new Date(ms);
    if (Number.isNaN(d.getTime())) return null;
    return d.toISOString().slice(0, 10);
  }
  const t = Date.parse(String(v));
  if (Number.isNaN(t)) return null;
  return new Date(t).toISOString().slice(0, 10);
}

/**
 * Converts an Excel cell value to ISO calendar date `YYYY-MM-DD`, or null if missing/invalid.
 * Handles Date objects, Excel serial numbers (1900-based), and parseable strings. It assumes dates are in LOCAL Time.
 * @param {unknown} v
 * @returns {string | null}
 * ArunS: New function which assumes Local Time, instead of UTC.
 */
function toISODate(dateVal) {
  if (!dateVal) return null;
  
  const d = new Date(dateVal);
  // Check for invalid date
  if (isNaN(d.getTime())) return null; 

  const year = d.getFullYear();
  // Months are 0-indexed, so add 1. Pad with leading zero if needed.
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}



/**
 * Converts a 2D array (first row = headers) into an array of plain objects keyed by header.
 * Skips completely empty rows.
 * @param {string[][]} rows
 * @returns {Record<string, unknown>[]}
 */
function rowsToObjects(rows) {
  if (!rows.length) return [];
  const headers = rows[0].map((h) => norm(h));
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => c == null || c === "")) continue;
    const o = {};
    for (let c = 0; c < headers.length; c++) {
      o[headers[c]] = row[c];
    }
    out.push(o);
  }
  return out;
}

/**
 * Looks up a column value from a row object by trying several possible header spellings
 * (e.g. trailing spaces in "URL of Job Posting ").
 * @param {Record<string, unknown>} raw
 * @param {...string} candidates Header names to try in order
 * @returns {unknown}
 */
function getCol(raw, ...candidates) {
  const keys = Object.keys(raw);
  for (const want of candidates) {
    const w = norm(want);
    const hit = keys.find((k) => norm(k) === w);
    if (hit != null) return raw[hit];
  }
  return null;
}

/** Milestone columns: order defines timeline step sequence in the report. */
const STEP_DEFS = [
  { key: "applied", label: "Applied", candidates: ["Date- Applied"] },
  { key: "shortListed", label: "Short listed", candidates: ["Date- Short Listed"] },
  { key: "interview", label: "Interview / test", candidates: ["Date- Interview/Test"] },
  { key: "offer", label: "Offer", candidates: ["Date- Offer"] },
  { key: "rejected", label: "Rejected", candidates: ["Date- Rejected"] },
  { key: "declined", label: "Declined by me", candidates: ["Date- Declined by Me"] },
];




/**
 * Maps one workbook row to the JSON shape consumed by report-template.html.
 * Requires at least `Date- Applied` for the row to be kept downstream.
 * @param {Record<string, unknown>} raw
 * @param {number} idx Row index for stable(ish) id generation
 */
function buildRecord(raw, idx) {
  const company = norm(getCol(raw, "Company Name")) || "(Unknown company)";
  const position = norm(getCol(raw, "Position Name")) || "(Role)";
  const roleType = norm(getCol(raw, "Role Type")) || "(Unspecified)";
  const status = norm(getCol(raw, "Status")) || "—";
  const portal = norm(getCol(raw, "Portal"));
  const flag = norm(
    getCol(raw, "Flag to see if the job posting is still accepting candidates")
  );
  const url = String(getCol(raw, "URL of Job Posting", "URL of Job Posting ") ?? "").trim();
  const log = String(getCol(raw, "Log") ?? "").trim();

  const dates = {
    jobPosted: toISODate(getCol(raw, "Date- Job Posted")),
    applied: toISODate(getCol(raw, "Date- Applied")),
    shortListed: toISODate(getCol(raw, "Date- Short Listed")),
    interview: toISODate(getCol(raw, "Date- Interview/Test")),
    offer: toISODate(getCol(raw, "Date- Offer")),
    rejected: toISODate(getCol(raw, "Date- Rejected")),
    declined: toISODate(getCol(raw, "Date- Declined by Me")),
  };

  const steps = [];
  for (const { key, label, candidates } of STEP_DEFS) {
    const iso = toISODate(getCol(raw, ...candidates));
    if (iso) steps.push({ key, label, date: iso });
  }
  steps.sort((a, b) => a.date.localeCompare(b.date));

  const id = `${idx}-${company.slice(0, 24)}`.replace(/\s+/g, "-");

  return {
    id,
    company,
    position,
    roleType,
    status,
    portal,
    flag,
    url,
    log: log.length > 400 ? log.slice(0, 400) + "…" : log,
    dates,
    steps,
    label: `${company} — ${position}`.slice(0, 120),
  };
}

/**
 * Reads an .xlsm/.xlsx path via Node fs and parses with SheetJS from a buffer
 * (more reliable on Windows than XLSX.readFile for some paths).
 * @param {string} p Absolute path to workbook
 */
function loadWorkbook(p) {
  const buf = fs.readFileSync(p);
  return XLSX.read(buf, {
    type: "buffer",
    cellDates: true,
    cellFormula: false,
    cellNF: false,
    sheetStubs: false,
  });
}

/**
 * Entry point: load workbook, extract applications, merge into HTML template, write dist file.
 */
function main() {
  let chosen = workbookPath;
  if (!fs.existsSync(chosen)) {
    console.error("Workbook not found:", chosen);
    console.error("Place JobHunt.xlsm in the repo root, or pass a path: node build.mjs \"D:\\path\\JobHunt.xlsm\"");
    process.exit(1);
  }

  let wb;
  try {
    wb = loadWorkbook(chosen);
  } catch (err) {
    console.error(err.message || err);
    process.exit(1);
  }
  const sheetName = wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    console.error("No worksheet found.");
    process.exit(1);
  }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const objects = rowsToObjects(rows);
  const applications = objects.map((o, i) => buildRecord(o, i)).filter((a) => a.dates.applied);

  if (!applications.length) {
    console.error("No rows with an application date were found. Check sheet name and headers.");
    process.exit(1);
  }

  fs.mkdirSync(outDir, { recursive: true });
  const templatePath = path.join(__dirname, "report-template.html");
  let html = fs.readFileSync(templatePath, "utf8");
  const payload = JSON.stringify({
    generatedAt: new Date().toISOString(),
    source: path.basename(chosen),
    sheet: sheetName,
    applications,
  });
  const b64 = Buffer.from(payload, "utf8").toString("base64");
  html = html.replace("__APP_B64__", b64);
  fs.writeFileSync(outHtml, html, "utf8");
  console.log(
    "Wrote",
    outHtml,
    "—",
    applications.length,
    "applications (from",
    path.basename(chosen) + ")"
  );
}

main();
