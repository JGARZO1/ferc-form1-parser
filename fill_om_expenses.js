/**
 * Fill om_expenses.xlsx with Electric Operation and Maintenance Expenses
 * from a FERC Form 1 HTML filing.
 *
 * - Does NOT modify the template file
 * - Copies template → output, then fills:
 *     B1  utility_id
 *     B2  report_year
 *     For each template data row (by line_no in col A):
 *       C  amount_for_current_year
 *       D  amount_for_previous_year
 *     Column B account labels from the template are preserved
 * - Writes under a company folder (e.g. CLECO/om_expenses_C000447_2022.xlsx)
 * - If COMPANY folder already exists, reuses it (no duplicate folders)
 *
 * Usage:
 *   node fill_om_expenses.js [html] [template] [outputName]
 *
 * Defaults (this folder):
 *   html       = ClecoPowerLlc2022.html
 *   template   = om_expenses.xlsx
 *   outputName = om_expenses_<utilityId>_<year>.xlsx  (inside COMPANY/)
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

/**
 * Parse a FERC numeric cell: "(a) 1,234", "(1,234)", blank → number|null
 */
function parseNumber(value) {
  let text = cleanText(value);
  if (!text) return null;
  // drop footnote markers like (a), (b) before the number
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
 * Normalize line numbers for map keys: "4", "64.1", "93.1"
 * Handles Excel float artifacts like 64.099999999999994 → 64.1
 */
function lineKey(raw) {
  const t = cleanText(String(raw));
  if (!t) return "";
  if (!/^\d+(\.\d+)?$/.test(t)) return t;
  const n = Number(t);
  if (!Number.isFinite(n)) return t;
  if (Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));
  // snap common one-decimal line numbers (x.1)
  const one = Math.round(n * 10) / 10;
  if (Math.abs(n - one) < 1e-6) {
    const s = one.toFixed(1);
    return s.endsWith(".0") ? String(Math.round(one)) : s;
  }
  return String(n);
}

// ---------------------------------------------------------------------------
// O&M extraction
// ---------------------------------------------------------------------------

/**
 * FERC Electric Operation and Maintenance Expenses columns:
 *   0 Line No.
 *   1 Account / description (a)
 *   2 Amount for Current Year (b)
 *   3 Amount for Previous Year (c)
 */
function cellsToOmRow(cells) {
  const key = lineKey(cells[0]);
  if (!key || !/^\d+(\.\d+)?$/.test(key)) return null;
  if (cells.length < 3) return null;

  return {
    line_no: key,
    account: cleanText(cells[1] || ""),
    amount_for_current_year: parseNumber(cells[2]),
    amount_for_previous_year: parseNumber(cells[3]),
  };
}

/**
 * Extract Electric O&M expense rows from FERC Form 1 HTML.
 * @returns {Map<string, object>} line_no → row
 */
function parseOmExpenses(html) {
  // Note: filing uses "Operations" (plural)
  const idMatch =
    html.match(
      /id="(ScheduleElectricOperationsAndMaintenanceExpensesAbstract[^"]*)"/i
    ) ||
    html.match(
      /id="(ScheduleElectricOperationAndMaintenanceExpensesAbstract[^"]*)"/i
    );

  let start = -1;
  if (idMatch) start = html.indexOf(idMatch[0]);
  if (start < 0) {
    start = html.search(
      /ELECTRIC OPERATION[S]? AND MAINTENANCE EXPENSES/i
    );
  }
  if (start < 0) {
    throw new Error(
      "Electric Operation and Maintenance Expenses section not found in HTML"
    );
  }

  // Do NOT use "Purchased Power" text — it appears as line 76 inside O&M.
  // Prefer the next schedule's id after this section.
  const endMarkers = [
    'id="SchedulePurchasedPowerAbstract',
    'id="ScheduleTransmissionOfElectricityForOthers',
    'id="ScheduleSalesForResaleAbstract',
    'id="ScheduleDepreciationDepletionAndAmortizationAbstract',
  ];
  let end = html.length;
  for (const mk of endMarkers) {
    const i = html.indexOf(mk, start + 1000);
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
    if (cells.length < 3) continue;
    if (!/^\d+(\.\d+)?$/.test(cleanText(cells[0]))) continue;

    const row = cellsToOmRow(cells);
    if (!row) continue;

    // Skip pure section headers (no amounts, abstract/section labels)
    const hasAmount =
      row.amount_for_current_year != null ||
      row.amount_for_previous_year != null;
    const looksLikeHeader =
      /Abstract\b/i.test(row.account) ||
      /^(POWER PRODUCTION|STEAM POWER|NUCLEAR|HYDRAULIC|OTHER POWER|TRANSMISSION EXPENSES|REGIONAL MARKET|DISTRIBUTION EXPENSES|CUSTOMER ACCOUNTS|CUSTOMER SERVICE|SALES EXPENSES|ADMINISTRATIVE)/i.test(
        row.account
      ) ||
      /^[A-E]\.\s/i.test(row.account) ||
      /^\d+\.\s+(POWER|TRANSMISSION|REGIONAL|DISTRIBUTION|CUSTOMER|SALES|ADMINISTRATIVE)/i.test(
        row.account
      ) ||
      /^(Operation|Maintenance)\s*$/i.test(row.account);

    // Keep header-only template lines that map by number even if empty
    // (they still get empty metrics). Store all numeric line rows.
    if (!byLine.has(row.line_no)) {
      byLine.set(row.line_no, row);
    } else if (hasAmount && !looksLikeHeader) {
      // Prefer a later/more complete data row over empty TOC-like hits
      const prev = byLine.get(row.line_no);
      const prevHas =
        prev.amount_for_current_year != null ||
        prev.amount_for_previous_year != null;
      if (!prevHas) byLine.set(row.line_no, row);
    }
  }

  if (!byLine.size) {
    throw new Error("No Electric O&M expense data rows parsed from HTML");
  }
  return byLine;
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

function readLineNoFromRowXml(rowInner) {
  const a = rowInner.match(/<c r="A\d+"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/);
  if (!a) return null;
  const body = a[2] || "";
  const inline = body.match(/<t[^>]*>([\s\S]*?)<\/t>/);
  if (inline) return lineKey(inline[1]);
  const v = body.match(/<v>([^<]*)<\/v>/);
  if (v) return lineKey(v[1]);
  return null;
}

function fillMetricsRow(sheetXml, excelRow, data) {
  let xml = sheetXml;
  xml = setCellValue(xml, `C${excelRow}`, data.amount_for_current_year);
  xml = setCellValue(xml, `D${excelRow}`, data.amount_for_previous_year);
  return xml;
}

function writeXlsxFromUnpacked(unpackedDir, outputPath) {
  _packXlsx(unpackedDir, outputPath);
}

// ---------------------------------------------------------------------------
// Main API
// ---------------------------------------------------------------------------

/**
 * Copy om_expenses template, fill from HTML, write into company folder.
 */
function fillOmExpensesFromHtml({
  htmlPath,
  templatePath,
  baseDir,
  outputFileName,
  companyName,
  utilityId,
  reportYear,
  sheetName = "om_expenses",
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

  const byLine = parseOmExpenses(html);

  const companyShort =
    companyName || companyFolderName(respondent, title, htmlPath);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `om_expenses_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "om-fill-"));
  const zipCopy = path.join(tmp, "book.zip");
  const unpacked = path.join(tmp, "out");
  fs.copyFileSync(templatePath, zipCopy);
  _unpackXlsx(zipCopy, unpacked);

  const sheetPath = path.join(unpacked, "xl", "worksheets", "sheet1.xml");
  let sheet = fs.readFileSync(sheetPath, "utf8");

  sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
  sheet = setCellValue(sheet, "B2", resolvedYear);

  const rowMatches = [
    ...sheet.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g),
  ];
  let filled = 0;
  let filledWithAmounts = 0;
  const missing = [];
  const samples = [];

  for (const rm of rowMatches) {
    const excelRow = Number(rm[1]);
    if (excelRow < 4) continue; // skip metadata + header

    const ln = readLineNoFromRowXml(rm[2]);
    if (!ln) continue;

    const data = byLine.get(ln);
    if (!data) {
      missing.push(ln);
      continue;
    }

    sheet = fillMetricsRow(sheet, excelRow, data);
    filled++;
    if (
      data.amount_for_current_year != null ||
      data.amount_for_previous_year != null
    ) {
      filledWithAmounts++;
    }
    if (samples.length < 8 && data.amount_for_current_year != null) {
      samples.push(data);
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
    filled,
    filledWithAmounts,
    missing,
    htmlRowCount: byLine.size,
    byLine,
    samples,
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
  const templatePath = path.resolve(dir, process.argv[3] || "om_expenses.xlsx");
  const outputArg = process.argv[4] || null;

  const result = fillOmExpensesFromHtml({
    htmlPath,
    templatePath,
    baseDir: dir,
    outputFileName: outputArg ? path.basename(outputArg) : undefined,
    sheetName: "om_expenses",
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
    `Filled ${result.filled} template row(s) (${result.filledWithAmounts} with amounts) from ${result.htmlRowCount} HTML line(s)`
  );
  if (result.missing.length) {
    console.log(
      "Template lines with no HTML match:",
      result.missing.slice(0, 30).join(", ") +
        (result.missing.length > 30
          ? ` ... (+${result.missing.length - 30} more)`
          : "")
    );
  }

  const grand = result.byLine.get("198");
  if (grand) {
    console.log(
      `Grand total (L198): current=${grand.amount_for_current_year} previous=${grand.amount_for_previous_year}`
    );
  }

  for (const s of result.samples) {
    console.log(
      `  L${s.line_no}: curr=${s.amount_for_current_year} prev=${s.amount_for_previous_year} | ${s.account.slice(0, 50)}`
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
  parseOmExpenses,
  parseNumber,
  fillOmExpensesFromHtml,
  companyFolderName,
  resolveCompanyDir,
  extractFilingMeta,
  lineKey,
};
