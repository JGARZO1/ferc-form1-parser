/**
 * Fill about.xlsx template with cover-page / contact metadata from a FERC Form 1 HTML filing.
 *
 * - Does NOT modify the template file
 * - Copies template → output, then fills column B for:
 *     utility_id, report_year, address, name_of_contact, title_of_contact,
 *     address_of_contact, telephone_of_contact, date_of_report
 * - Writes under a company folder (e.g. CLECO/about_C000447_2021.xlsx)
 * - If COMPANY folder already exists, reuses it (no duplicate folders)
 *
 * Usage:
 *   node fill_about_info.js [html] [template] [outputName]
 *
 * Defaults (this folder):
 *   html       = ClecoPowerLlc2021.html
 *   template   = about.xlsx
 *   outputName = about_<utilityId>_<year>.xlsx  (inside COMPANY/)
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
// Text helpers
// ---------------------------------------------------------------------------

function cleanText(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&nbsp;/g, " ")
    .replace(/&#160;/g, " ")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ")
    .trim();
}

function stripTags(html) {
  return cleanText(String(html || "").replace(/<[^>]+>/g, " "));
}

/**
 * First non-empty iXBRL fact value for ferc:<concept>.
 */
function grabIxFact(html, concept) {
  const name = concept.startsWith("ferc:") ? concept : `ferc:${concept}`;
  const re = new RegExp(
    `name="${name.replace(/:/g, "\\:")}"[^>]*>([\\s\\S]*?)</ix:`,
    "gi"
  );
  let m;
  while ((m = re.exec(html))) {
    const v = stripTags(m[1]);
    if (v) return v;
  }
  return "";
}

/**
 * All non-empty values for a concept (order of appearance).
 */
function grabIxFacts(html, concept) {
  const name = concept.startsWith("ferc:") ? concept : `ferc:${concept}`;
  const re = new RegExp(
    `name="${name.replace(/:/g, "\\:")}"[^>]*>([\\s\\S]*?)</ix:`,
    "gi"
  );
  const out = [];
  let m;
  while ((m = re.exec(html))) {
    const v = stripTags(m[1]);
    if (v) out.push(v);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Company folder (same behavior as fill_sales_by_rate_totals.js)
// ---------------------------------------------------------------------------

function extractRespondentInfo(html) {
  const respondent =
    grabIxFact(html, "RespondentLegalName") ||
    grabIxFact(html, "EntityName") ||
    "";
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
  const safe = brand.replace(/[^A-Za-z0-9_-]/g, "").toUpperCase();
  return safe || "UNKNOWN";
}

/**
 * Reuse existing company folder (case-insensitive); create only if missing.
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
      company: existing.name,
      created: false,
    };
  }

  const companyDir = path.join(baseDir, desired);
  fs.mkdirSync(companyDir, { recursive: true });
  return { companyDir, company: desired, created: true };
}

// ---------------------------------------------------------------------------
// About metadata extraction
// ---------------------------------------------------------------------------

/**
 * Extract about-sheet fields from FERC Form 1 HTML.
 * @returns {{
 *   utility_id: string,
 *   report_year: number|null,
 *   address: string,
 *   name_of_contact: string,
 *   title_of_contact: string,
 *   address_of_contact: string,
 *   telephone_of_contact: string,
 *   date_of_report: string,
 * }}
 */
function parseAboutInfo(html) {
  const utility_id =
    grabIxFact(html, "CompanyIdentifier") ||
    grabIxFact(html, "RespondentIdentifier") ||
    grabIxFact(html, "FilerIdentifier") ||
    "";

  const yearRaw = grabIxFact(html, "ReportYear");
  const report_year = yearRaw ? Number(yearRaw) : null;

  // Principal office preferred; fall back to principal business address
  const address =
    grabIxFact(html, "AddressOfPrincipalOfficeAtEndOfPeriod") ||
    grabIxFact(html, "PrincipalBusinessAddress") ||
    "";

  const name_of_contact = grabIxFact(html, "NameOfContactPerson") || "";
  const title_of_contact = grabIxFact(html, "TitleOfContactPerson") || "";
  const address_of_contact =
    grabIxFact(html, "AddressOfContactPerson") || address || "";
  const telephone_of_contact =
    grabIxFact(html, "TelephoneOfContactPerson") ||
    grabIxFact(html, "TelephoneNumber") ||
    "";

  // Report date on cover; fall back to attestation date
  const date_of_report =
    grabIxFact(html, "ReportDate") ||
    grabIxFact(html, "AttestationDate") ||
    grabIxFact(html, "DocumentDate") ||
    "";

  return {
    utility_id,
    report_year: Number.isFinite(report_year) ? report_year : null,
    address,
    name_of_contact,
    title_of_contact,
    address_of_contact,
    telephone_of_contact,
    date_of_report,
  };
}

// ---------------------------------------------------------------------------
// XLSX cell helpers (preserve template styles)
// ---------------------------------------------------------------------------

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Set or replace a cell value in sheet XML, preserving style when present. */
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

  // Insert missing cell into its row
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

function writeXlsxFromUnpacked(unpackedDir, outputPath) {
  _packXlsx(unpackedDir, outputPath);
}

// ---------------------------------------------------------------------------
// Main fill API
// ---------------------------------------------------------------------------

/**
 * Copy about.xlsx template, fill metadata from HTML, write into company folder.
 */
function fillAboutFromHtml({
  htmlPath,
  templatePath,
  baseDir,
  outputFileName,
  companyName,
  sheetName = "about",
}) {
  if (!fs.existsSync(htmlPath)) throw new Error(`HTML not found: ${htmlPath}`);
  if (!fs.existsSync(templatePath))
    throw new Error(`Template not found: ${templatePath}`);

  const html = fs.readFileSync(htmlPath, "utf8");
  const { respondent, title } = extractRespondentInfo(html);
  const about = parseAboutInfo(html);

  if (!about.utility_id) {
    console.warn("Warning: utility_id not found in HTML");
  }
  if (about.report_year == null) {
    throw new Error("Could not determine report_year from HTML");
  }

  const companyShort =
    companyName || companyFolderName(respondent, title, htmlPath);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `about_${about.utility_id || "UNKNOWN"}_${about.report_year}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  // Unpack template copy
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "about-fill-"));
  const zipCopy = path.join(tmp, "book.zip");
  const unpacked = path.join(tmp, "out");
  fs.copyFileSync(templatePath, zipCopy);
  _unpackXlsx(zipCopy, unpacked);

  const sheetPath = path.join(unpacked, "xl", "worksheets", "sheet1.xml");
  let sheet = fs.readFileSync(sheetPath, "utf8");

  // Template layout: A=label, B=value (rows 1–8)
  const fills = [
    { ref: "B1", value: about.utility_id, asString: true },
    { ref: "B2", value: about.report_year, asString: false },
    { ref: "B3", value: about.address, asString: true },
    { ref: "B4", value: about.name_of_contact, asString: true },
    { ref: "B5", value: about.title_of_contact, asString: true },
    { ref: "B6", value: about.address_of_contact, asString: true },
    { ref: "B7", value: about.telephone_of_contact, asString: true },
    { ref: "B8", value: about.date_of_report, asString: true },
  ];

  for (const f of fills) {
    sheet = setCellValue(sheet, f.ref, f.value, { asString: f.asString });
  }
  fs.writeFileSync(sheetPath, sheet, "utf8");

  // Ensure sheet name
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

  // Soft-remove stale root copy from older runs
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
    about,
  };
}

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------

function main() {
  const dir = __dirname;
  const htmlPath = path.resolve(dir, process.argv[2] || "ClecoPowerLlc2021.html");
  const templatePath = path.resolve(dir, process.argv[3] || "about.xlsx");
  const outputArg = process.argv[4] || null;

  const result = fillAboutFromHtml({
    htmlPath,
    templatePath,
    baseDir: dir,
    outputFileName: outputArg ? path.basename(outputArg) : undefined,
    sheetName: "about",
  });

  console.log(
    "Company:",
    result.company,
    `(${result.respondent || "n/a"})`,
    result.companyFolderCreated ? "[folder created]" : "[existing folder]"
  );
  console.log("Folder:", result.companyDir);
  console.log("Wrote:", result.outputPath);
  for (const [k, v] of Object.entries(result.about)) {
    console.log(`  ${k}: ${v}`);
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
  parseAboutInfo,
  fillAboutFromHtml,
  companyFolderName,
  resolveCompanyDir,
  grabIxFact,
};
