/**
 * Fill sales_by_rate.xlsx template with All-Accounts totals from a FERC Form 1 HTML filing.
 *
 * - Does NOT modify the template file
 * - Copies template → output, then fills only:
 *     B1  utility_id
 *     B2  report_year
 *     C4:G4  line 41 Total Billed - All Accounts
 *     C5:G5  line 42 Total Unbilled Revenue - All Accounts
 *     C6:G6  line 43 Total - All Accounts
 * - Writes under a company folder (e.g. CLECO/sales_by_rate_C000447_2021.xlsx)
 * - If COMPANY folder already exists, reuses it (no duplicate folders)
 *
 * Usage:
 *   node fill_sales_by_rate_totals.js [html] [template] [outputName]
 *
 * Defaults (this folder):
 *   html       = ClecoPowerLlc2021.html
 *   template   = sales_by_rate.xlsx
 *   outputName = sales_by_rate_<utilityId>_<year>.xlsx  (inside COMPANY/)
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const { execSync } = require("child_process");

// --- cross-platform xlsx pack/unpack ---
function _unpackXlsx(zipCopy, unpacked) {
  const { execSync } = require("child_process");
  const fs = require("fs");
  fs.mkdirSync(unpacked, { recursive: true });
  if (process.platform === "win32") {
    try {
      const z = String(zipCopy).replace(/'/g, "''");
      const u = String(unpacked).replace(/'/g, "''");
      execSync(
        "powershell -NoProfile -Command \"Expand-Archive -LiteralPath '" + z + "' -DestinationPath '" + u + "' -Force\"",
        { stdio: "pipe" }
      );
      return;
    } catch (e) { /* fall through to unzip */ }
  }
  execSync("unzip -q -o \"" + zipCopy + "\" -d \"" + unpacked + "\"", { stdio: "pipe" });
}

function _packXlsx(unpacked, finalPath) {
  const { execSync } = require("child_process");
  const fs = require("fs");
  const path = require("path");
  fs.mkdirSync(path.dirname(finalPath), { recursive: true });
  if (fs.existsSync(finalPath)) fs.unlinkSync(finalPath);
  if (process.platform === "win32") {
    try {
      const ps1 = path.join(require("os").tmpdir(), "zip_ferc_" + Date.now() + ".ps1");
      const fp = String(finalPath).replace(/'/g, "''");
      const up = String(unpacked).replace(/'/g, "''");
      const lines = [
        "Add-Type -AssemblyName System.IO.Compression.FileSystem",
        "if (Test-Path -LiteralPath '" + fp + "') { Remove-Item -LiteralPath '" + fp + "' -Force }",
        "[System.IO.Compression.ZipFile]::CreateFromDirectory('" + up + "', '" + fp + "', [System.IO.Compression.CompressionLevel]::Optimal, $false)",
      ];
      fs.writeFileSync(ps1, lines.join(String.fromCharCode(13, 10)));
      execSync("powershell -NoProfile -ExecutionPolicy Bypass -File \"" + ps1 + "\"", { stdio: "pipe" });
      try { fs.unlinkSync(ps1); } catch (_) {}
      return;
    } catch (e) { /* fall through */ }
  }
  execSync("cd \"" + unpacked + "\" && zip -qr \"" + finalPath + "\" .", { stdio: "pipe" });
}
// --- end cross-platform helpers ---

const TOTAL_PATTERNS = [
  {
    key: "billed",
    match: /TOTAL\s+Billed\s+-\s+All\s+Accounts/i,
    // template row 4 labels stay; we only write metrics
  },
  {
    key: "unbilled",
    match: /TOTAL\s+Unbilled\s+Rev\.\s*\(See\s+Instr\.\s*6\)\s*-\s*All\s+Accounts/i,
  },
  {
    key: "grand",
    // HTML sometimes has two spaces: "TOTAL -  All Accounts"
    match: /TOTAL\s+-\s+All\s+Accounts/i,
  },
];

function cleanText(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stripTags(html) {
  return cleanText(String(html || "").replace(/<[^>]+>/g, " "));
}

/**
 * Pull respondent legal name (and optional HTML title) from a FERC Form 1 filing.
 */
function extractRespondentInfo(html) {
  const grabIx = (name) => {
    const re = new RegExp(
      `name="${name.replace(/:/g, "\\:")}"[^>]*>([\\s\\S]*?)</ix:`,
      "i"
    );
    const m = html.match(re);
    return m ? stripTags(m[1]) : "";
  };

  const respondent =
    grabIx("ferc:RespondentLegalName") ||
    grabIx("ferc:EntityName") ||
    "";

  const titleM = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const title = titleM ? stripTags(titleM[1]) : "";

  // Filename-style cues sometimes appear in title: "ClecoPowerLlc2022"
  return { respondent, title };
}

/**
 * Short clean company folder name, e.g. "Cleco Power LLC" → "CLECO".
 * Prefer brand token; drop legal/generic words.
 */
function companyFolderName(respondent, title = "", htmlPath = "") {
  const LEGAL =
    /\b(LLC|L\.?L\.?C\.?|Inc\.?|Incorporated|Corp\.?|Corporation|Company|Co\.|Ltd\.?|Limited|LP|L\.P\.|PLC)\b/gi;
  const GENERIC =
    /^(Power|Electric|Electrical|Energy|Utility|Utilities|Gas|Transmission|Service|Services|Holding|Holdings|Group|of|and|the|d\/b\/a)$/i;

  // Known short aliases (first match wins)
  const ALIASES = [
    { test: /\bcleco\b/i, name: "CLECO" },
    { test: /\bswepco\b|southwestern\s+electric\s+power/i, name: "SWEPCO" },
    { test: /\bentergy\b/i, name: "ENTERGY" },
  ];

  const sources = [respondent, title, path.basename(htmlPath, path.extname(htmlPath))]
    .map((s) => cleanText(s))
    .filter(Boolean);

  for (const src of sources) {
    for (const a of ALIASES) {
      if (a.test.test(src)) return a.name;
    }
  }

  let raw = sources[0] || "UNKNOWN";
  // If camel/pascal filename like ClecoPowerLlc2022, split and drop year
  if (!/\s/.test(raw) && /[a-z][A-Z]/.test(raw)) {
    raw = raw
      .replace(/([a-z])([A-Z])/g, "$1 $2")
      .replace(/([A-Za-z])(\d+)/g, "$1")
      .trim();
  }

  raw = raw.replace(LEGAL, " ").replace(/[,\.]/g, " ");
  const words = raw.split(/\s+/).filter((w) => w && !GENERIC.test(w));
  const brand = words[0] || sources[0] || "UNKNOWN";
  const safe = brand.replace(/[^A-Za-z0-9_-]/g, "").toUpperCase();
  return safe || "UNKNOWN";
}

/**
 * Resolve company folder under baseDir.
 * If a folder with the same name already exists (case-insensitive), reuse it —
 * never create a second folder for the same company.
 * Only creates the directory when none exists yet.
 *
 * @returns {{ companyDir: string, company: string, created: boolean }}
 */
function resolveCompanyDir(baseDir, companyShortName) {
  const desired = companyShortName || "UNKNOWN";
  fs.mkdirSync(baseDir, { recursive: true });

  const entries = fs.readdirSync(baseDir, { withFileTypes: true });
  const existing = entries.find(
    (e) => e.isDirectory() && e.name.toLowerCase() === desired.toLowerCase()
  );

  if (existing) {
    return {
      companyDir: path.join(baseDir, existing.name),
      company: existing.name, // keep on-disk casing (e.g. CLECO)
      created: false,
    };
  }

  const companyDir = path.join(baseDir, desired);
  fs.mkdirSync(companyDir, { recursive: true });
  return { companyDir, company: desired, created: true };
}

/**
 * Detect utility_id and report_year from FERC HTML when not provided.
 */
function extractFilingMeta(html) {
  const grabIx = (name) => {
    const re = new RegExp(
      `name="${name.replace(/:/g, "\\:")}"[^>]*>([\\s\\S]*?)</ix:`,
      "i"
    );
    const m = html.match(re);
    return m ? stripTags(m[1]) : "";
  };

  const utilityId =
    grabIx("ferc:CompanyIdentifier") ||
    grabIx("ferc:RespondentIdentifier") ||
    grabIx("ferc:FilerIdentifier") ||
    "";

  const yearRaw = grabIx("ferc:ReportYear");
  const reportYear = yearRaw ? Number(yearRaw) : null;

  return { utilityId, reportYear };
}

/**
 * Parse a FERC numeric cell: "(a) 1,234", "(1,234)", "7.0200", blank → number|null
 */
function parseNumber(value) {
  let text = cleanText(value);
  if (!text) return null;
  text = text.replace(/^\([a-z]\)\s*/i, "").trim();
  if (!text) return null;
  const neg = /^\(.*\)$/.test(text);
  const norm = text.replace(/[(),$]/g, "").replace(/,/g, "");
  if (!norm || /[a-zA-Z]/.test(norm)) return null;
  const n = Number(norm);
  if (Number.isNaN(n)) return null;
  return neg ? -n : n;
}

/**
 * Extract the three All-Accounts total lines from FERC Form 1 HTML.
 * @param {string} html
 * @returns {{ billed: object, unbilled: object, grand: object }}
 */
function parseAllAccountsTotals(html) {
  const results = {};

  // Restrict search to Sales of Electricity by Rate Schedules region when possible
  const start = html.indexOf("ScheduleSalesOfElectricityByRateSchedulesAbstract");
  const endMarkers = [
    'id="ScheduleSalesForResale',
    "SALES FOR RESALE (Account 447)",
  ];
  let end = html.length;
  for (const m of endMarkers) {
    const i = html.indexOf(m, start > 0 ? start : 0);
    if (i > 0 && i < end) end = i;
  }
  const region =
    start >= 0 ? html.slice(start, end) : html;

  const trRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let m;
  while ((m = trRe.exec(region))) {
    const cells = [...m[1].matchAll(/<td\b[^>]*>([\s\S]*?)<\/td>/gi)].map((td) =>
      stripTags(td[1])
    );
    if (cells.length < 2) continue;
    const desc = cells[1] || "";
    for (const p of TOTAL_PATTERNS) {
      if (p.match.test(desc) && !results[p.key]) {
        results[p.key] = {
          line_no: Number(cells[0]) || null,
          description: desc,
          mwh_sold: parseNumber(cells[2]),
          revenue: parseNumber(cells[3]),
          avg_customers: parseNumber(cells[4]),
          kwh_per_customer: parseNumber(cells[5]),
          revenue_per_kwh: parseNumber(cells[6]),
        };
      }
    }
  }

  for (const p of TOTAL_PATTERNS) {
    if (!results[p.key]) {
      throw new Error(`Could not find All-Accounts total: ${p.key}`);
    }
  }
  return results;
}

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Set or replace a cell's value in sheet XML, preserving style attribute when present. */
function setCellValue(sheetXml, ref, value, { asString = false } = {}) {
  const cellRe = new RegExp(`<c r="${ref}"([^>/]*)(?:/>|>([\\s\\S]*?)</c>)`);
  const m = sheetXml.match(cellRe);

  let inner;
  let attrs;
  if (value === null || value === undefined || value === "") {
    // leave empty cell (self-closing or empty body), keep style
    if (!m) return sheetXml;
    attrs = m[1] || "";
    // strip t= and keep s=
    attrs = attrs.replace(/\s*t="[^"]*"/g, "");
    const empty = `<c r="${ref}"${attrs}/>`;
    return sheetXml.replace(cellRe, empty);
  }

  if (asString || typeof value === "string") {
    inner = ` t="inlineStr"><is><t>${xmlEscape(String(value))}</t></is></c>`;
  } else {
    inner = `><v>${value}</v></c>`;
  }

  if (m) {
    let attrs = m[1] || "";
    // remove existing type; we'll set via inner for strings, or none for numbers
    attrs = attrs.replace(/\s*t="[^"]*"/g, "");
    if (asString || typeof value === "string") {
      // t="inlineStr" goes in opening tag
      const open = `<c r="${ref}"${attrs} t="inlineStr"><is><t>${xmlEscape(
        String(value)
      )}</t></is></c>`;
      return sheetXml.replace(cellRe, open);
    }
    const open = `<c r="${ref}"${attrs}><v>${value}</v></c>`;
    return sheetXml.replace(cellRe, open);
  }

  // Cell missing: insert before </row> of its row number
  const rowNum = ref.replace(/^[A-Z]+/, "");
  const rowRe = new RegExp(`(<row r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
  const rm = sheetXml.match(rowRe);
  if (!rm) throw new Error(`Row ${rowNum} not found for cell ${ref}`);
  let newCell;
  if (asString || typeof value === "string") {
    newCell = `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(
      String(value)
    )}</t></is></c>`;
  } else {
    newCell = `<c r="${ref}"><v>${value}</v></c>`;
  }
  return sheetXml.replace(rowRe, `$1$2${newCell}$3`);
}

function fillMetrics(sheetXml, row, totals) {
  const cols = [
    ["C", totals.mwh_sold],
    ["D", totals.revenue],
    ["E", totals.avg_customers],
    ["F", totals.kwh_per_customer],
    ["G", totals.revenue_per_kwh],
  ];
  let xml = sheetXml;
  for (const [col, val] of cols) {
    xml = setCellValue(xml, `${col}${row}`, val);
  }
  return xml;
}

/**
 * Copy template → company folder output and fill All-Accounts totals + metadata.
 * Optionally rename sheet to sales_by_rate.
 *
 * @param {object} opts
 * @param {string} opts.htmlPath
 * @param {string} opts.templatePath
 * @param {string} [opts.outputPath] - full path; if omitted, uses baseDir/COMPANY/outputFileName
 * @param {string} [opts.baseDir] - parent of company folder (default: dirname of template)
 * @param {string} [opts.outputFileName] - file name only (default: sales_by_rate_<id>_<year>.xlsx)
 * @param {string} [opts.utilityId]
 * @param {number} [opts.reportYear]
 * @param {string} [opts.sheetName]
 * @param {string} [opts.companyName] - override folder name (else detected from HTML)
 */
function fillSalesByRateFromHtml({
  htmlPath,
  templatePath,
  outputPath,
  baseDir,
  outputFileName,
  utilityId,
  reportYear,
  sheetName = "sales_by_rate",
  companyName,
}) {
  if (!fs.existsSync(htmlPath)) throw new Error(`HTML not found: ${htmlPath}`);
  if (!fs.existsSync(templatePath))
    throw new Error(`Template not found: ${templatePath}`);

  const html = fs.readFileSync(htmlPath, "utf8");
  const { respondent, title } = extractRespondentInfo(html);
  const filingMeta = extractFilingMeta(html);
  const resolvedUtilityId = utilityId || filingMeta.utilityId || "UNKNOWN";
  const resolvedYear =
    reportYear != null
      ? Number(reportYear)
      : filingMeta.reportYear != null
        ? Number(filingMeta.reportYear)
        : null;

  if (resolvedYear == null || Number.isNaN(resolvedYear)) {
    throw new Error("Could not determine report_year from HTML or arguments");
  }

  const companyShort =
    companyName || companyFolderName(respondent, title, htmlPath);
  const totals = parseAllAccountsTotals(html);

  const root = baseDir || path.dirname(path.resolve(templatePath));
  // Reuse existing company folder if present; do not create duplicates
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    (outputPath ? path.basename(outputPath) : null) ||
    `sales_by_rate_${resolvedUtilityId}_${resolvedYear}.xlsx`;

  // Always write into the resolved company folder
  const finalPath = path.join(companyDir, path.basename(fileName));

  // Work in a temp package copied from template
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "sbr-fill-"));
  const zipCopy = path.join(tmp, "book.zip");
  const unpacked = path.join(tmp, "out");
  fs.copyFileSync(templatePath, zipCopy);
  _unpackXlsx(zipCopy, unpacked);

  const sheetPath = path.join(unpacked, "xl", "worksheets", "sheet1.xml");
  let sheet = fs.readFileSync(sheetPath, "utf8");

  // Metadata
  sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
  sheet = setCellValue(sheet, "B2", resolvedYear);

  // Three total lines (template rows 4–6); labels already in template
  sheet = fillMetrics(sheet, 4, totals.billed);
  sheet = fillMetrics(sheet, 5, totals.unbilled);
  sheet = fillMetrics(sheet, 6, totals.grand);

  fs.writeFileSync(sheetPath, sheet, "utf8");

  // Rename sheet if requested (template default is sales_by_rate_schedule)
  if (sheetName) {
    const wbPath = path.join(unpacked, "xl", "workbook.xml");
    let wb = fs.readFileSync(wbPath, "utf8");
    wb = wb.replace(
      /name="sales_by_rate_schedule"/,
      `name="${sheetName.replace(/"/g, "")}"`
    );
    // also if already sales_by_rate, leave alone
    if (!wb.includes(`name="${sheetName}"`)) {
      wb = wb.replace(
        /(<sheet\s+name=")[^"]+(")/,
        `$1${sheetName.replace(/"/g, "")}$2`
      );
    }
    fs.writeFileSync(wbPath, wb, "utf8");
  }

  _packXlsx(unpacked, finalPath);

  fs.rmSync(tmp, { recursive: true, force: true });

  // Remove stale copy left in the base dir (same file name from older runs)
  const stale = path.join(root, path.basename(finalPath));
  if (
    fs.existsSync(stale) &&
    path.resolve(stale).toLowerCase() !== path.resolve(finalPath).toLowerCase()
  ) {
    try {
      fs.unlinkSync(stale);
    } catch (err) {
      // File may be open in Excel; non-fatal
      console.warn(
        `Could not remove stale file ${stale}: ${err.message}`
      );
    }
  }

  return {
    outputPath: finalPath,
    companyDir,
    company,
    companyFolderCreated,
    respondent,
    utilityId: resolvedUtilityId,
    reportYear: resolvedYear,
    totals,
  };
}

function main() {
  const dir = __dirname;
  const htmlPath = path.resolve(dir, process.argv[2] || "ClecoPowerLlc2021.html");
  const templatePath = path.resolve(dir, process.argv[3] || "sales_by_rate.xlsx");
  // argv[4] = bare file name; always placed inside existing/new company folder
  const outputArg = process.argv[4] || null;

  const result = fillSalesByRateFromHtml({
    htmlPath,
    templatePath,
    baseDir: dir,
    outputFileName: outputArg
      ? path.basename(outputArg)
      : undefined, // default: sales_by_rate_<utilityId>_<year>.xlsx
    sheetName: "sales_by_rate",
  });

  console.log(
    "Company:",
    result.company,
    `(${result.respondent || "n/a"})`,
    result.companyFolderCreated ? "[folder created]" : "[existing folder]"
  );
  console.log("Folder:", result.companyDir);
  console.log("Wrote:", result.outputPath);
  console.log("utility_id:", result.utilityId);
  console.log("report_year:", result.reportYear);
  for (const [k, row] of Object.entries(result.totals)) {
    console.log(
      `${k}: line ${row.line_no} | MWh=${row.mwh_sold} | Rev=${row.revenue} | Cust=${row.avg_customers} | kWh/c=${row.kwh_per_customer} | $/kWh=${row.revenue_per_kwh}`
    );
  }
}

if (require.main === module) {
  try {
    main();
  } catch (e) {
    console.error(e.message || e);
    process.exit(1);
  }
}

module.exports = {
  parseAllAccountsTotals,
  parseNumber,
  extractRespondentInfo,
  extractFilingMeta,
  companyFolderName,
  resolveCompanyDir,
  fillSalesByRateFromHtml,
};
