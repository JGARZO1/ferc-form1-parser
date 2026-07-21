/**
 * Fill plant_in_service.xlsx with Electric Plant In Service (Accounts 101, 102, 103, 106)
 * from a FERC Form 1 HTML filing.
 *
 * - Does NOT modify the template file
 * - Copies template → output, then fills:
 *     B1  utility_id
 *     B2  report_year
 *     For each template data row (by line_no in col A): C–H metrics
 *       balance_BoY, additions, retirements, adjustments, transfers, balance_EoY
 *     Column B account labels from the template are preserved
 * - Writes under a company folder (e.g. CLECO/plant_in_service_C000447_2022.xlsx)
 * - If COMPANY folder already exists, reuses it (no duplicate folders)
 * - Validates EoY ≈ BoY + Additions − Retirements + Adjustments + Transfers
 *
 * Usage:
 *   node fill_plant_in_service.js [html] [template] [outputName]
 *
 * Defaults (this folder):
 *   html       = ClecoPowerLlc2022.html
 *   template   = plant_in_service.xlsx
 *   outputName = plant_in_service_<utilityId>_<year>.xlsx  (inside COMPANY/)
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

const VALIDATION_TOLERANCE = 1; // dollar rounding

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
  text = text.replace(/^\([a-z]\)\s*/i, "").trim();
  if (!text) return null;
  const neg = /^\(.*\)$/.test(text);
  const norm = text.replace(/[(),$]/g, "").replace(/,/g, "");
  if (!norm || /[a-zA-Z]/.test(norm)) return null;
  const n = Number(norm);
  if (Number.isNaN(n)) return null;
  return neg ? -n : n;
}

/** Normalize line numbers for map keys: "2", "44.1", "48.1" */
function lineKey(raw) {
  const t = cleanText(String(raw));
  if (!t) return "";
  // Excel may store 44.1 as 44.099999… — normalize
  if (/^\d+(\.\d+)?$/.test(t)) {
    const n = Number(t);
    if (!Number.isFinite(n)) return t;
    // Prefer one-decimal form for .1 lines when close
    if (Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));
    const one = Math.round(n * 10) / 10;
    if (Math.abs(n - one) < 1e-6) return one.toFixed(1).replace(/\.0$/, "");
    return String(n);
  }
  return t;
}

// ---------------------------------------------------------------------------
// Electric Plant In Service extraction
// ---------------------------------------------------------------------------

/**
 * FERC columns:
 *   0 Line No.
 *   1 Account / description (a)
 *   2 Balance Beginning of Year (b)
 *   3 Additions (c)
 *   4 Retirements (d)
 *   5 Adjustments (e)
 *   6 Transfers (f)
 *   7 Balance End of Year (g)
 */
function cellsToPlantRow(cells) {
  const key = lineKey(cells[0]);
  if (!key || !/^\d+(\.\d+)?$/.test(key)) return null;

  // Section headers like "1. INTANGIBLE PLANT" have no metrics and are
  // not present as fillable template rows; still skip junk with too few cols
  if (cells.length < 3) return null;

  return {
    line_no: key,
    account: cleanText(cells[1] || ""),
    balance_BoY: parseNumber(cells[2]),
    additions: parseNumber(cells[3]),
    retirements: parseNumber(cells[4]),
    adjustments: parseNumber(cells[5]),
    transfers: parseNumber(cells[6]),
    balance_EoY: parseNumber(cells[7]),
  };
}

/**
 * Extract Electric Plant In Service rows from FERC Form 1 HTML.
 * @returns {Map<string, object>} line_no → row
 */
function parsePlantInService(html) {
  const idMatch = html.match(
    /id="(ScheduleElectricPlantInServiceAbstract[^"]*)"/i
  );
  let start = -1;
  if (idMatch) {
    start = html.indexOf(idMatch[0]);
  }
  if (start < 0) {
    start = html.search(/ELECTRIC PLANT IN SERVICE \(Account/i);
  }
  if (start < 0) {
    start = html.search(/ELECTRIC PLANT IN SERVICE/i);
  }
  if (start < 0) {
    throw new Error("Electric Plant In Service section not found in HTML");
  }

  const endMarkers = [
    "ACCUMULATED PROVISION FOR DEPRECIATION OF ELECTRIC UTILITY PLANT",
    "ACCUMULATED PROVISION FOR DEPRECIATION",
    'id="ScheduleAccumulatedProvisionForDepreciation',
    "ELECTRIC PLANT LEASED TO OTHERS",
    'id="ScheduleElectricPlantLeasedToOthers',
    "CONSTRUCTION WORK IN PROGRESS",
    'id="ScheduleConstructionWorkInProgress',
  ];
  let end = html.length;
  for (const mk of endMarkers) {
    const i = html.indexOf(mk, start + 200);
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

    // Plant-in-service data rows typically have 7–8 columns
    // (line, account, 6 money cols). Allow sparse rows.
    if (cells.length < 4 && !/total|subtotal/i.test(cells[1] || "")) continue;

    const row = cellsToPlantRow(cells);
    if (!row) continue;

    // Skip pure section headers with no numbers and no account code
    const hasMetric =
      row.balance_BoY != null ||
      row.additions != null ||
      row.retirements != null ||
      row.adjustments != null ||
      row.transfers != null ||
      row.balance_EoY != null;
    const looksLikeHeader =
      /^(INTANGIBLE|PRODUCTION|STEAM|NUCLEAR|HYDRAULIC|OTHER PRODUCTION|TRANSMISSION|DISTRIBUTION|REGIONAL|GENERAL)\b/i.test(
        row.account
      ) || /^[A-D]\.\s/i.test(row.account) || /^\d+\.\s/i.test(row.account);
    if (!hasMetric && looksLikeHeader) continue;

    // Prefer first occurrence of each line within the section
    if (!byLine.has(row.line_no)) byLine.set(row.line_no, row);
  }

  if (!byLine.size) {
    throw new Error("No Electric Plant In Service data rows parsed from HTML");
  }
  return byLine;
}

/**
 * EoY should equal BoY + Additions - Retirements + Adjustments + Transfers.
 * Missing metrics treated as 0 for the check.
 */
function validatePlantRow(row) {
  const boy = row.balance_BoY ?? 0;
  const add = row.additions ?? 0;
  const ret = row.retirements ?? 0;
  const adj = row.adjustments ?? 0;
  const xfer = row.transfers ?? 0;
  const eoy = row.balance_EoY;
  if (eoy == null && boy === 0 && add === 0 && ret === 0 && adj === 0 && xfer === 0) {
    return { ok: true, expected: null, actual: null, diff: 0, skipped: true };
  }
  if (eoy == null) {
    return { ok: false, expected: boy + add - ret + adj + xfer, actual: null, diff: null, skipped: false };
  }
  const expected = boy + add - ret + adj + xfer;
  const diff = eoy - expected;
  return {
    ok: Math.abs(diff) <= VALIDATION_TOLERANCE,
    expected,
    actual: eoy,
    diff,
    skipped: false,
  };
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
 * Read line_no from column A of a sheet row (numeric or inlineStr / shared string index not resolved).
 * Template uses plain numeric <v> for line numbers.
 */
function readLineNoFromRowXml(rowInner) {
  const a = rowInner.match(
    /<c r="A\d+"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/
  );
  if (!a) return null;
  const body = a[2] || "";
  const inline = body.match(/<t[^>]*>([\s\S]*?)<\/t>/);
  if (inline) return lineKey(inline[1]);
  const v = body.match(/<v>([^<]*)<\/v>/);
  if (v) return lineKey(v[1]);
  return null;
}

function fillMetricsRow(sheetXml, excelRow, data) {
  const cols = [
    ["C", data.balance_BoY],
    ["D", data.additions],
    ["E", data.retirements],
    ["F", data.adjustments],
    ["G", data.transfers],
    ["H", data.balance_EoY],
  ];
  let xml = sheetXml;
  for (const [col, val] of cols) {
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
 * Copy plant_in_service template, fill from HTML, write into company folder.
 */
function fillPlantInServiceFromHtml({
  htmlPath,
  templatePath,
  baseDir,
  outputFileName,
  companyName,
  utilityId,
  reportYear,
  sheetName = "plant_in_service",
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

  const byLine = parsePlantInService(html);

  const companyShort =
    companyName || companyFolderName(respondent, title, htmlPath);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `plant_in_service_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "pis-fill-"));
  const zipCopy = path.join(tmp, "book.zip");
  const unpacked = path.join(tmp, "out");
  fs.copyFileSync(templatePath, zipCopy);
  _unpackXlsx(zipCopy, unpacked);

  const sheetPath = path.join(unpacked, "xl", "worksheets", "sheet1.xml");
  let sheet = fs.readFileSync(sheetPath, "utf8");

  sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
  sheet = setCellValue(sheet, "B2", resolvedYear);

  // Walk template data rows and fill metrics by line_no
  const rowMatches = [...sheet.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)];
  let filled = 0;
  let missing = [];
  const validationFailures = [];
  const validationOk = [];

  for (const rm of rowMatches) {
    const excelRow = Number(rm[1]);
    if (excelRow < 4) continue; // skip metadata + header

    const ln = readLineNoFromRowXml(rm[2]);
    if (!ln) continue;

    const data = byLine.get(ln);
    if (!data) {
      // Also try alternate key forms (44.1 vs 44.10)
      const alt = byLine.get(lineKey(Number(ln)));
      if (!alt) {
        missing.push(ln);
        continue;
      }
      sheet = fillMetricsRow(sheet, excelRow, alt);
      filled++;
      const v = validatePlantRow(alt);
      if (!v.skipped) {
        if (v.ok) validationOk.push(ln);
        else validationFailures.push({ line: ln, ...v, account: alt.account });
      }
      continue;
    }

    sheet = fillMetricsRow(sheet, excelRow, data);
    filled++;
    const v = validatePlantRow(data);
    if (!v.skipped) {
      if (v.ok) validationOk.push(ln);
      else validationFailures.push({ line: ln, ...v, account: data.account });
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
    missing,
    htmlRowCount: byLine.size,
    byLine,
    validationOk,
    validationFailures,
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
    process.argv[3] || "plant_in_service.xlsx"
  );
  const outputArg = process.argv[4] || null;

  const result = fillPlantInServiceFromHtml({
    htmlPath,
    templatePath,
    baseDir: dir,
    outputFileName: outputArg ? path.basename(outputArg) : undefined,
    sheetName: "plant_in_service",
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
    `Filled ${result.filled} template row(s) from ${result.htmlRowCount} HTML line(s)`
  );
  if (result.missing.length) {
    console.log(
      "Template lines with no HTML match:",
      result.missing.join(", ")
    );
  }

  // Highlight grand total
  const grand =
    result.byLine.get("104") || result.byLine.get("100");
  if (grand) {
    console.log(
      `Grand (line ${grand.line_no}): BoY=${grand.balance_BoY} Add=${grand.additions} Ret=${grand.retirements} Adj=${grand.adjustments} Xfer=${grand.transfers} EoY=${grand.balance_EoY}`
    );
  }

  console.log(
    `Validation: ${result.validationOk.length} ok, ${result.validationFailures.length} mismatch (tol ±${VALIDATION_TOLERANCE})`
  );
  for (const f of result.validationFailures.slice(0, 15)) {
    console.log(
      `  FAIL L${f.line}: EoY=${f.actual} expected=${f.expected} diff=${f.diff} | ${f.account}`
    );
  }
  if (result.validationFailures.length > 15) {
    console.log(
      `  ... and ${result.validationFailures.length - 15} more failures`
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
  parsePlantInService,
  parseNumber,
  validatePlantRow,
  fillPlantInServiceFromHtml,
  companyFolderName,
  resolveCompanyDir,
  extractFilingMeta,
  lineKey,
};
