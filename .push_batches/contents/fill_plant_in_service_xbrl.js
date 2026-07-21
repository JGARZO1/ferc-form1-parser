/**
 * Fill plant_in_service.xlsx from a FERC Form 1 XBRL filing.
 *
 * Maps template line_no → concept via lib/plant_xbrl_map.js.
 * BoY/EoY use instant contexts; flow columns use primary duration.
 *
 * Usage:
 *   node fill_plant_in_service_xbrl.js [xbrl] [template] [outputName]
 */

const fs = require("fs");
const path = require("path");
const {
  loadXbrl,
  getNumericFact,
  detectXbrlMeta,
  detectCompanyFromXbrl,
} = require("./xbrl_common");
const { resolveCompanyDir, resolveTemplatePath } = require("./lib/company");
const { openTemplateCopy, setCellValue, setSheetName } = require("./lib/xlsx_utils");
const { PLANT_LINE_CONCEPTS } = require("./lib/plant_xbrl_map");

const VALIDATION_TOLERANCE = 1;

function lineKey(raw) {
  const t = String(raw || "").trim();
  if (!t) return "";
  if (/^\d+(\.\d+)?$/.test(t)) {
    const n = Number(t);
    if (!Number.isFinite(n)) return t;
    if (Math.abs(n - Math.round(n)) < 1e-9) return String(Math.round(n));
    const one = Math.round(n * 10) / 10;
    if (Math.abs(n - one) < 1e-6) {
      const s = one.toFixed(1);
      return s.endsWith(".0") ? String(Math.round(one)) : s;
    }
    return String(n);
  }
  return t;
}

function plantMetricsForConcept(doc, concept) {
  const boyCtx = doc.boyInstant ? doc.boyInstant.id : null;
  const eoyCtx = doc.eoyInstant ? doc.eoyInstant.id : null;
  const flowCtx = doc.primaryDuration ? doc.primaryDuration.id : null;

  const balance_BoY = getNumericFact(doc, concept, boyCtx);
  const balance_EoY = getNumericFact(doc, concept, eoyCtx);
  const additions = getNumericFact(doc, `${concept}Additions`, flowCtx);
  const retirements = getNumericFact(doc, `${concept}Retirements`, flowCtx);
  const adjustments = getNumericFact(doc, `${concept}Adjustments`, flowCtx);
  const transfers = getNumericFact(doc, `${concept}Transfers`, flowCtx);

  return {
    balance_BoY,
    additions,
    retirements,
    adjustments,
    transfers,
    balance_EoY,
  };
}

function parsePlantInServiceFromXbrl(doc) {
  const byLine = new Map();
  for (const [line, concept] of Object.entries(PLANT_LINE_CONCEPTS)) {
    const m = plantMetricsForConcept(doc, concept);
    // Keep row if any metric present
    if (
      m.balance_BoY != null ||
      m.balance_EoY != null ||
      m.additions != null ||
      m.retirements != null ||
      m.adjustments != null ||
      m.transfers != null
    ) {
      byLine.set(line, { line_no: line, account: concept, ...m });
    }
  }
  return byLine;
}

/**
 * Read template line numbers from column A (rows 4+).
 */
function readTemplateLineRows(sheetXml) {
  const rows = [];
  const rowRe = /<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let m;
  while ((m = rowRe.exec(sheetXml))) {
    const r = Number(m[1]);
    if (r < 4) continue;
    const cellA = m[2].match(
      /<c r="A\d+"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/
    );
    if (!cellA) continue;
    let val = null;
    const body = cellA[2] || "";
    const v = body.match(/<v>([^<]*)<\/v>/);
    const t = body.match(/<t[^>]*>([^<]*)<\/t>/);
    if (t) val = t[1];
    else if (v) val = v[1];
    if (val == null || val === "") continue;
    const key = lineKey(val);
    if (key) rows.push({ excelRow: r, line_no: key });
  }
  return rows;
}

function fillPlantInServiceFromXbrl({
  xbrlPath,
  templatePath,
  baseDir,
  outputFileName,
  utilityId,
  reportYear,
  companyName,
  sheetName = "plant_in_service",
}) {
  if (!fs.existsSync(xbrlPath)) throw new Error(`XBRL not found: ${xbrlPath}`);
  if (!fs.existsSync(templatePath))
    throw new Error(`Template not found: ${templatePath}`);

  const doc = loadXbrl(xbrlPath);
  const meta = detectXbrlMeta(doc, xbrlPath);
  const resolvedUtilityId = utilityId || meta.utilityId || "UNKNOWN";
  const resolvedYear =
    reportYear != null ? Number(reportYear) : meta.reportYear;
  if (resolvedYear == null || Number.isNaN(resolvedYear)) {
    throw new Error("Could not determine report_year from XBRL");
  }

  const companyShort =
    companyName || detectCompanyFromXbrl(doc, xbrlPath);
  const byLine = parsePlantInServiceFromXbrl(doc);
  if (!byLine.size) {
    throw new Error("No Electric Plant In Service facts matched in XBRL");
  }

  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `plant_in_service_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const pack = openTemplateCopy(templatePath);
  let filled = 0;
  const validationWarnings = [];
  try {
    const sheetPath = path.join(pack.unpacked, "xl", "worksheets", "sheet1.xml");
    let sheet = fs.readFileSync(sheetPath, "utf8");
    sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
    sheet = setCellValue(sheet, "B2", resolvedYear);

    const templateRows = readTemplateLineRows(sheet);
    for (const tr of templateRows) {
      const data = byLine.get(tr.line_no);
      if (!data) continue;
      const cols = [
        ["C", data.balance_BoY],
        ["D", data.additions],
        ["E", data.retirements],
        ["F", data.adjustments],
        ["G", data.transfers],
        ["H", data.balance_EoY],
      ];
      for (const [col, val] of cols) {
        sheet = setCellValue(sheet, `${col}${tr.excelRow}`, val);
      }
      filled++;

      if (
        data.balance_BoY != null &&
        data.balance_EoY != null &&
        (data.additions != null ||
          data.retirements != null ||
          data.adjustments != null ||
          data.transfers != null)
      ) {
        const expected =
          data.balance_BoY +
          (data.additions || 0) -
          (data.retirements || 0) +
          (data.adjustments || 0) +
          (data.transfers || 0);
        if (Math.abs(expected - data.balance_EoY) > VALIDATION_TOLERANCE) {
          validationWarnings.push({
            line_no: tr.line_no,
            expected,
            balance_EoY: data.balance_EoY,
          });
        }
      }
    }

    fs.writeFileSync(sheetPath, sheet, "utf8");
    setSheetName(pack.unpacked, sheetName);
    pack.finish(finalPath);
  } finally {
    pack.cleanup();
  }

  return {
    outputPath: finalPath,
    companyDir,
    company,
    companyFolderCreated,
    respondent: meta.respondent,
    utilityId: resolvedUtilityId,
    reportYear: resolvedYear,
    filled,
    matchedConcepts: byLine.size,
    validationWarnings,
  };
}

function main() {
  const dir = __dirname;
  const xbrlPath = path.resolve(
    dir,
    process.argv[2] || "filings/ClecoPowerLlc2020.xbrl"
  );
  const templatePath = resolveTemplatePath(
    dir,
    process.argv[3] || "plant_in_service.xlsx"
  );
  const result = fillPlantInServiceFromXbrl({
    xbrlPath,
    templatePath,
    baseDir: dir,
    outputFileName: process.argv[4]
      ? path.basename(process.argv[4])
      : undefined,
  });
  console.log("Company:", result.company);
  console.log("Wrote:", result.outputPath);
  console.log(
    `Filled ${result.filled} template rows from ${result.matchedConcepts} concepts`
  );
  if (result.validationWarnings.length) {
    console.log("Validation warnings:", result.validationWarnings.length);
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
  parsePlantInServiceFromXbrl,
  fillPlantInServiceFromXbrl,
};
