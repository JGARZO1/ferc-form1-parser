/**
 * Fill om_expenses.xlsx from a FERC Form 1 XBRL filing.
 *
 * Maps template line_no → concept via lib/om_xbrl_map.js.
 * Current year = primary duration; previous year = prior duration.
 *
 * Usage:
 *   node fill_om_expenses_xbrl.js [xbrl] [template] [outputName]
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
const { OM_LINE_CONCEPTS } = require("./lib/om_xbrl_map");

function lineKey(raw) {
  const t = String(raw || "").trim();
  if (!t) return "";
  if (!/^\d+(\.\d+)?$/.test(t)) return t;
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

function parseOmExpensesFromXbrl(doc) {
  const curCtx = doc.primaryDuration ? doc.primaryDuration.id : null;
  const prevCtx = doc.priorDuration ? doc.priorDuration.id : null;
  const byLine = new Map();

  for (const [line, concept] of Object.entries(OM_LINE_CONCEPTS)) {
    const amount_for_current_year = getNumericFact(doc, concept, curCtx);
    const amount_for_previous_year = getNumericFact(doc, concept, prevCtx);
    if (
      amount_for_current_year != null ||
      amount_for_previous_year != null
    ) {
      byLine.set(line, {
        line_no: line,
        account: concept,
        amount_for_current_year,
        amount_for_previous_year,
      });
    }
  }
  return byLine;
}

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

function fillOmExpensesFromXbrl({
  xbrlPath,
  templatePath,
  baseDir,
  outputFileName,
  utilityId,
  reportYear,
  companyName,
  sheetName = "om_expenses",
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
  const byLine = parseOmExpensesFromXbrl(doc);
  if (!byLine.size) {
    throw new Error("No O&M expense facts matched in XBRL");
  }

  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `om_expenses_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const pack = openTemplateCopy(templatePath);
  let filled = 0;
  try {
    const sheetPath = path.join(pack.unpacked, "xl", "worksheets", "sheet1.xml");
    let sheet = fs.readFileSync(sheetPath, "utf8");
    sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
    sheet = setCellValue(sheet, "B2", resolvedYear);

    const templateRows = readTemplateLineRows(sheet);
    for (const tr of templateRows) {
      const data = byLine.get(tr.line_no);
      if (!data) continue;
      sheet = setCellValue(
        sheet,
        `C${tr.excelRow}`,
        data.amount_for_current_year
      );
      sheet = setCellValue(
        sheet,
        `D${tr.excelRow}`,
        data.amount_for_previous_year
      );
      filled++;
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
    process.argv[3] || "om_expenses.xlsx"
  );
  const result = fillOmExpensesFromXbrl({
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
  parseOmExpensesFromXbrl,
  fillOmExpensesFromXbrl,
};
