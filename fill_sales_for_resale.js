/**
 * Fill sales_for_resale.xlsx with Sales for Resale (Account 447) from a FERC Form 1 HTML filing.
 *
 * - Does NOT modify the template file
 * - Copies template → output, then fills:
 *     B1  utility_id
 *     B2  report_year
 *     Rows for lines 1–14: counterparty detail (B–L)
 *     Lines 15–17: subtotal_rq / subtotal_non_rq / total metrics (H–L, etc.)
 *       (template labels in column B for 15–17 are preserved)
 * - Writes under a company folder (e.g. CLECO/sales_for_resale_C000447_2022.xlsx)
 * - If COMPANY folder already exists, reuses it (no duplicate folders)
 *
 * Usage:
 *   node fill_sales_for_resale.js [html] [template] [outputName]
 *
 * Defaults (this folder):
 *   html       = ClecoPowerLlc2022.html
 *   template   = sales_for_resale.xlsx
 *   outputName = sales_for_resale_<utilityId>_<year>.xlsx  (inside COMPANY/)
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

// Template: Excel row = line_no + 3  (line 1 → row 4, line 17 → row 20)
const LINE_TO_ROW = (lineNo) => lineNo + 3;
const MAX_DETAIL_LINE = 14;
const SUBTOTAL_LINES = new Set([15, 16, 17]);

// ---------------------------------------------------------------------------
// Shared-style helpers
// ---------------------------------------------------------------------------

function cleanText(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/&#160;/g, " ")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/\s+/g, " ")
    .trim();
}

function stripTags(html) {
  return cleanText(String(html || "").replace(/<[^>]+>/g, " "));
}

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
    grabIx("ferc:RespondentLegalName") || grabIx("ferc:EntityName") || "";
  const titleM = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const title = titleM ? stripTags(titleM[1]) : "";
  return { respondent, title };
}

function companyFolderName(respondent, title = "", htmlPath = "") {
  const LEGAL =
    /\b(LLC|L\.?L\.?C\.?|Inc\.?|Incorporated|Corp\.?|Corporation|Company|Co\.|Ltd\.?|Limited|LP|L\.P\.|PLC)\b/gi;
  const GENERIC =
    /^(Power|Electric|Electrical|Energy|Utility|Utilities|Gas|Transmission|Service|Services|Holding|Holdings|Group|of|and|the|d\/b\/a)$/i;
  const ALIASES = [
    { test: /\bcleco\b/i, name: "CLECO" },
    { test: /\bswepco\b|southwestern\s+electric\s+power/i, name: "SWEPCO" },
    { test: /\bentergy\b/i, name: "ENTERGY" },
  ];

  const sources = [
    respondent,
    title,
    path.basename(htmlPath, path.extname(htmlPath)),
  ]
    .map((s) => cleanText(s))
    .filter(Boolean);

  for (const src of sources) {
    for (const a of ALIASES) {
      if (a.test.test(src)) return a.name;
    }
  }

  let raw = sources[0] || "UNKNOWN";
  if (!/\s/.test(raw) && /[a-z][A-Z]/.test(raw)) {
    raw = raw
      .replace(/([a-z])([A-Z])/g, "$1 $2")
      .replace(/([A-Za-z])(\d+)/g, "$1")
      .trim();
  }
  raw = raw.replace(LEGAL, " ").replace(/[,\.]/g, " ");
  const words = raw.split(/\s+/).filter((w) => w && !GENERIC.test(w));
  const brand = words[0] || sources[0] || "UNKNOWN";
  return brand.replace(/[^A-Za-z0-9_-]/g, "").toUpperCase() || "UNKNOWN";
}

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
      company: existing.name,
      created: false,
    };
  }
  const companyDir = path.join(baseDir, desired);
  fs.mkdirSync(companyDir, { recursive: true });
  return { companyDir, company: desired, created: true };
}

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

// ---------------------------------------------------------------------------
// Sales for Resale extraction
// ---------------------------------------------------------------------------

/**
 * Parse one Sales for Resale data row from table cells.
 * Expected columns (FERC form):
 *   0 Line No.
 *   1 Name of Company (a)
 *   2 Statistical Classification (b)
 *   3 FERC Rate Schedule or Tariff Number (c)
 *   4 Average Monthly Billing Demand (d)
 *   5 Average Monthly NCP Demand (e)
 *   6 Average Monthly CP Demand (f)
 *   7 Megawatt Hours Sold (g)
 *   8 Demand Charges (h)
 *   9 Energy Charges (i)
 *  10 Other Charges (j)
 *  11 Total (k)
 */
function cellsToRow(cells) {
  const lineNo = Number(cells[0]);
  if (!Number.isFinite(lineNo)) return null;

  return {
    line_no: lineNo,
    name_of_company: cleanText(cells[1] || ""),
    statistical_classification: cleanText(cells[2] || ""),
    FERC_rate_schedule: cleanText(cells[3] || ""),
    average_monthly_billing: parseNumber(cells[4]),
    average_monthly_ncp: parseNumber(cells[5]),
    average_monthly_cp: parseNumber(cells[6]),
    megawatt_hours_sold: parseNumber(cells[7]),
    demand_revenue: parseNumber(cells[8]),
    energy_revenue: parseNumber(cells[9]),
    other_revenue: parseNumber(cells[10]),
    total: parseNumber(cells[11]),
  };
}

function isSalesForResaleDataRow(cells) {
  if (cells.length < 2) return false;
  if (!/^\d+$/.test(cells[0])) return false;
  // Detail / subtotal / total rows have at least line + name (and usually more cols)
  // Exclude O&M-style 2–3 column rows
  if (cells.length < 8 && !/subtotal|total/i.test(cells[1] || "")) return false;
  // Subtotal/total may have sparse numeric columns
  if (/subtotal|^\s*total\s*$/i.test(cells[1] || "") && cells.length >= 2)
    return true;
  return cells.length >= 8;
}

/**
 * Extract Sales for Resale table rows from FERC Form 1 HTML.
 * @returns {Array<object>} sorted by line_no
 */
function parseSalesForResale(html) {
  const start = html.search(
    /id="ScheduleSalesForResale|SALES FOR RESALE \(Account 447\)/i
  );
  if (start < 0) {
    throw new Error("Sales for Resale section not found in HTML");
  }

  const endMarkers = [
    'id="ScheduleElectricOperationAndMaintenance',
    "ELECTRIC OPERATION AND MAINTENANCE EXPENSES",
    'id="SchedulePurchasedPower',
    "PURCHASED POWER (Account 555)",
    'id="ScheduleTransmissionOfElectricityForOthers',
  ];
  let end = html.length;
  for (const mk of endMarkers) {
    const i = html.indexOf(mk, start + 80);
    if (i > start && i < end) end = i;
  }

  const region = html.slice(start, end);
  const byLine = new Map();
  const trRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let m;
  while ((m = trRe.exec(region))) {
    const cells = [...m[1].matchAll(/<td\b[^>]*>([\s\S]*?)<\/td>/gi)].map(
      (td) => stripTags(td[1])
    );
    if (!isSalesForResaleDataRow(cells)) continue;
    const row = cellsToRow(cells);
    if (!row) continue;
    // Prefer first occurrence of each line number within the section
    if (!byLine.has(row.line_no)) byLine.set(row.line_no, row);
  }

  const rows = [...byLine.values()].sort((a, b) => a.line_no - b.line_no);
  if (!rows.length) {
    throw new Error("No Sales for Resale data rows parsed from HTML");
  }
  return rows;
}

// ---------------------------------------------------------------------------
// XLSX helpers
// ---------------------------------------------------------------------------

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setCellValue(sheetXml, ref, value, { asString = false } = {}) {
  const cellRe = new RegExp(`<c r="${ref}"([^>/]*)(?:/>|>([\\s\\S]*?)</c>)`);
  const m = sheetXml.match(cellRe);

  if (value === null || value === undefined || value === "") {
    if (!m) return sheetXml;
    let attrs = m[1] || "";
    attrs = attrs.replace(/\s*t="[^"]*"/g, "");
    return sheetXml.replace(cellRe, `<c r="${ref}"${attrs}/>`);
  }

  const writeAsString = asString || typeof value === "string";

  if (m) {
    let attrs = m[1] || "";
    attrs = attrs.replace(/\s*t="[^"]*"/g, "");
    if (writeAsString) {
      return sheetXml.replace(
        cellRe,
        `<c r="${ref}"${attrs} t="inlineStr"><is><t>${xmlEscape(
          String(value)
        )}</t></is></c>`
      );
    }
    return sheetXml.replace(
      cellRe,
      `<c r="${ref}"${attrs}><v>${value}</v></c>`
    );
  }

  const rowNum = ref.replace(/^[A-Z]+/, "");
  const rowRe = new RegExp(`(<row r="${rowNum}"[^>]*>)([\\s\\S]*?)(</row>)`);
  if (!sheetXml.match(rowRe)) {
    throw new Error(`Row ${rowNum} not found for cell ${ref}`);
  }
  const newCell = writeAsString
    ? `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(
        String(value)
      )}</t></is></c>`
    : `<c r="${ref}"><v>${value}</v></c>`;
  return sheetXml.replace(rowRe, `$1$2${newCell}$3`);
}

/**
 * Fill one template data row from a parsed Sales for Resale record.
 * For lines 15–17, do not overwrite B (template has subtotal_rq / etc.).
 */
function fillDataRow(sheetXml, row) {
  const excelRow = LINE_TO_ROW(row.line_no);
  if (excelRow < 4 || excelRow > 20) return sheetXml;

  let xml = sheetXml;
  const isSubtotal = SUBTOTAL_LINES.has(row.line_no);

  const textCols = isSubtotal
    ? [] // keep B label; C–G usually blank on totals
    : [
        ["B", row.name_of_company],
        ["C", row.statistical_classification],
        ["D", row.FERC_rate_schedule],
      ];

  const numCols = isSubtotal
    ? [
        ["E", row.average_monthly_billing],
        ["F", row.average_monthly_ncp],
        ["G", row.average_monthly_cp],
        ["H", row.megawatt_hours_sold],
        ["I", row.demand_revenue],
        ["J", row.energy_revenue],
        ["K", row.other_revenue],
        ["L", row.total],
      ]
    : [
        ["E", row.average_monthly_billing],
        ["F", row.average_monthly_ncp],
        ["G", row.average_monthly_cp],
        ["H", row.megawatt_hours_sold],
        ["I", row.demand_revenue],
        ["J", row.energy_revenue],
        ["K", row.other_revenue],
        ["L", row.total],
      ];

  for (const [col, val] of textCols) {
    xml = setCellValue(xml, `${col}${excelRow}`, val || "", { asString: true });
  }
  for (const [col, val] of numCols) {
    xml = setCellValue(xml, `${col}${excelRow}`, val);
  }
  return xml;
}

function writeXlsxFromUnpacked(unpackedDir, outputPath) {
  _packXlsx(unpackedDir, outputPath);
}

// ---------------------------------------------------------------------------
// Main API
// ---------------------------------------------------------------------------

/**
 * Copy sales_for_resale template, fill from HTML, write into company folder.
 */
function fillSalesForResaleFromHtml({
  htmlPath,
  templatePath,
  baseDir,
  outputFileName,
  companyName,
  utilityId,
  reportYear,
  sheetName = "sales_for_resale",
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

  const rows = parseSalesForResale(html);
  const companyShort =
    companyName || companyFolderName(respondent, title, htmlPath);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `sales_for_resale_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "sfr-fill-"));
  const zipCopy = path.join(tmp, "book.zip");
  const unpacked = path.join(tmp, "out");
  fs.copyFileSync(templatePath, zipCopy);
  _unpackXlsx(zipCopy, unpacked);

  const sheetPath = path.join(unpacked, "xl", "worksheets", "sheet1.xml");
  let sheet = fs.readFileSync(sheetPath, "utf8");

  sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
  sheet = setCellValue(sheet, "B2", resolvedYear);

  let filledDetail = 0;
  let filledTotals = 0;
  const skipped = [];

  for (const row of rows) {
    if (row.line_no >= 1 && row.line_no <= MAX_DETAIL_LINE) {
      sheet = fillDataRow(sheet, row);
      filledDetail++;
    } else if (SUBTOTAL_LINES.has(row.line_no)) {
      sheet = fillDataRow(sheet, row);
      filledTotals++;
    } else {
      skipped.push(row.line_no);
    }
  }

  fs.writeFileSync(sheetPath, sheet, "utf8");

  if (sheetName) {
    const wbPath = path.join(unpacked, "xl", "workbook.xml");
    let wb = fs.readFileSync(wbPath, "utf8");
    if (!wb.includes(`name="${sheetName}"`)) {
      wb = wb.replace(
        /(<sheet\s+name=")[^"]+(")/,
        `$1${sheetName.replace(/"/g, "")}$2`
      );
      fs.writeFileSync(wbPath, wb, "utf8");
    }
  }

  writeXlsxFromUnpacked(unpacked, finalPath);
  fs.rmSync(tmp, { recursive: true, force: true });

  const stale = path.join(root, path.basename(finalPath));
  if (
    fs.existsSync(stale) &&
    path.resolve(stale).toLowerCase() !== path.resolve(finalPath).toLowerCase()
  ) {
    try {
      fs.unlinkSync(stale);
    } catch (err) {
      console.warn(`Could not remove stale file ${stale}: ${err.message}`);
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
    rows,
    filledDetail,
    filledTotals,
    skipped,
  };
}

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------

function main() {
  const dir = __dirname;
  const htmlPath = path.resolve(
    dir,
    process.argv[2] || "ClecoPowerLlc2022.html"
  );
  const templatePath = path.resolve(
    dir,
    process.argv[3] || "sales_for_resale.xlsx"
  );
  const outputArg = process.argv[4] || null;

  const result = fillSalesForResaleFromHtml({
    htmlPath,
    templatePath,
    baseDir: dir,
    outputFileName: outputArg ? path.basename(outputArg) : undefined,
    sheetName: "sales_for_resale",
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
  console.log(
    `Filled ${result.filledDetail} detail row(s), ${result.filledTotals} total/subtotal row(s)` +
      (result.skipped.length ? `; skipped lines: ${result.skipped.join(",")}` : "")
  );

  for (const row of result.rows) {
    const name =
      row.line_no >= 15
        ? row.name_of_company || `(line ${row.line_no})`
        : row.name_of_company;
    console.log(
      `  L${row.line_no}: ${name} | ${row.statistical_classification || ""} | MWh=${row.megawatt_hours_sold} | Total=${row.total}`
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
  parseSalesForResale,
  parseNumber,
  fillSalesForResaleFromHtml,
  companyFolderName,
  resolveCompanyDir,
  extractFilingMeta,
};
