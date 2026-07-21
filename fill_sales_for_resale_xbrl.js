/**
 * Fill sales_for_resale.xlsx from a FERC Form 1 XBRL filing.
 *
 * Detail rows come from SalesForResale axis segments; lines 15–17 are
 * unsegmented RQ / non-RQ / total aggregates when available.
 *
 * Usage:
 *   node fill_sales_for_resale_xbrl.js [xbrl] [template] [outputName]
 */

const fs = require("fs");
const path = require("path");
const {
  loadXbrl,
  getNumericFact,
  getStringFact,
  detectXbrlMeta,
  detectCompanyFromXbrl,
  groupBySegmentMember,
} = require("./xbrl_common");
const { resolveCompanyDir, resolveTemplatePath, cleanText } = require("./lib/company");
const { openTemplateCopy, setCellValue, setSheetName } = require("./lib/xlsx_utils");

const LINE_TO_ROW = (lineNo) => lineNo + 3;
const MAX_DETAIL_LINE = 14;

function factNum(group, concept) {
  const f = group.facts.get(concept);
  return f && f.numeric != null ? f.numeric : null;
}

function factStr(group, concept) {
  const f = group.facts.get(concept);
  return f && f.value ? f.value : "";
}

function parseSalesForResaleFromXbrl(doc) {
  const groups = groupBySegmentMember(doc, "SalesForResale");
  const detail = [];

  for (const [, g] of groups) {
    const name =
      factStr(g, "NameOfCompanyOrPublicAuthorityPurchasingElectricity") ||
      factStr(g, "NameOfCompanyOrPublicAuthority") ||
      factStr(g, "NameOfPurchaser") ||
      "";
    const mwh = factNum(g, "MegawattHoursSoldSalesForResale");
    const demand = factNum(g, "DemandChargesRevenueSalesForResale");
    const energy = factNum(g, "EnergyChargesRevenueSalesForResale");
    const other = factNum(g, "OtherChargesRevenueSalesForResale");
    const total =
      factNum(g, "SalesForResale") != null
        ? factNum(g, "SalesForResale")
        : demand != null || energy != null || other != null
          ? (demand || 0) + (energy || 0) + (other || 0)
          : null;

    // Skip empty groups
    if (!name && mwh == null && total == null) continue;

    const classif =
      factStr(g, "StatisticalClassification") ||
      factStr(g, "StatisticalClassificationCode") ||
      "";
    const rate =
      factStr(g, "RateScheduleTariffNumber") ||
      factStr(g, "FERCRateScheduleOrTariffNumber") ||
      factStr(g, "RateScheduleNumber") ||
      "";

    detail.push({
      member: g.member,
      name_of_company: name,
      statistical_classification: classif,
      FERC_rate_schedule: rate,
      average_monthly_billing: factNum(g, "AverageMonthlyBillingDemand"),
      average_monthly_ncp: factNum(g, "AverageMonthlyNonCoincidentPeakDemand"),
      average_monthly_cp: factNum(g, "AverageMonthlyCoincidentPeakDemand"),
      megawatt_hours_sold: mwh,
      demand_revenue: demand,
      energy_revenue: energy,
      other_revenue: other,
      total,
    });
  }

  // Stable sort by member key (0-1, 0-2, ...)
  detail.sort((a, b) => {
    const na = Number(String(a.member).replace(/[^\d.]/g, "")) || 0;
    const nb = Number(String(b.member).replace(/[^\d.]/g, "")) || 0;
    return na - nb;
  });

  const rows = detail.slice(0, MAX_DETAIL_LINE).map((r, i) => ({
    line_no: i + 1,
    ...r,
  }));

  const ctxId = doc.primaryDuration ? doc.primaryDuration.id : null;
  const totalRow = {
    line_no: 17,
    name_of_company: "total",
    megawatt_hours_sold: getNumericFact(
      doc,
      "MegawattHoursSoldSalesForResale",
      ctxId
    ),
    demand_revenue: getNumericFact(
      doc,
      "DemandChargesRevenueSalesForResale",
      ctxId
    ),
    energy_revenue: getNumericFact(
      doc,
      "EnergyChargesRevenueSalesForResale",
      ctxId
    ),
    other_revenue: getNumericFact(
      doc,
      "OtherChargesRevenueSalesForResale",
      ctxId
    ),
    total: getNumericFact(doc, "SalesForResale", ctxId),
  };

  // RQ / non-RQ subtotals if present as separate concepts
  const subtotalRq = {
    line_no: 15,
    name_of_company: "subtotal_rq",
    megawatt_hours_sold: getNumericFact(
      doc,
      "MegawattHoursSoldRequiredSalesForResale",
      ctxId
    ),
    total: getNumericFact(doc, "RequiredSalesForResale", ctxId),
  };
  const subtotalNonRq = {
    line_no: 16,
    name_of_company: "subtotal_non_rq",
    megawatt_hours_sold: getNumericFact(
      doc,
      "MegawattHoursSoldNonRequiredSalesForResale",
      ctxId
    ),
    total:
      getNumericFact(doc, "NonRequiredSalesForResale", ctxId) ??
      getNumericFact(doc, "NonRequiredSalesForResaleEnergy", ctxId),
  };

  if (!rows.length && totalRow.total == null) {
    throw new Error("No Sales for Resale data found in XBRL");
  }

  return { rows, subtotalRq, subtotalNonRq, totalRow };
}

function writeDetailRow(sheet, excelRow, row) {
  let s = sheet;
  // B–L: name, class, rate, billing, ncp, cp, mwh, demand, energy, other, total
  const pairs = [
    ["B", row.name_of_company, true],
    ["C", row.statistical_classification, true],
    ["D", row.FERC_rate_schedule, true],
    ["E", row.average_monthly_billing, false],
    ["F", row.average_monthly_ncp, false],
    ["G", row.average_monthly_cp, false],
    ["H", row.megawatt_hours_sold, false],
    ["I", row.demand_revenue, false],
    ["J", row.energy_revenue, false],
    ["K", row.other_revenue, false],
    ["L", row.total, false],
  ];
  for (const [col, val, asString] of pairs) {
    if (val === undefined) continue;
    s = setCellValue(s, `${col}${excelRow}`, val, { asString });
  }
  return s;
}

function writeSubtotalRow(sheet, excelRow, row) {
  let s = sheet;
  // Preserve label in B; write metrics H–L
  const pairs = [
    ["H", row.megawatt_hours_sold],
    ["I", row.demand_revenue],
    ["J", row.energy_revenue],
    ["K", row.other_revenue],
    ["L", row.total],
  ];
  for (const [col, val] of pairs) {
    if (val === undefined) continue;
    s = setCellValue(s, `${col}${excelRow}`, val);
  }
  return s;
}

function fillSalesForResaleFromXbrl({
  xbrlPath,
  templatePath,
  baseDir,
  outputFileName,
  utilityId,
  reportYear,
  companyName,
  sheetName = "sales_for_resale",
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
  const parsed = parseSalesForResaleFromXbrl(doc);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `sales_for_resale_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const pack = openTemplateCopy(templatePath);
  try {
    const sheetPath = path.join(pack.unpacked, "xl", "worksheets", "sheet1.xml");
    let sheet = fs.readFileSync(sheetPath, "utf8");
    sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
    sheet = setCellValue(sheet, "B2", resolvedYear);

    for (const row of parsed.rows) {
      sheet = writeDetailRow(sheet, LINE_TO_ROW(row.line_no), row);
    }
    sheet = writeSubtotalRow(sheet, LINE_TO_ROW(15), parsed.subtotalRq);
    sheet = writeSubtotalRow(sheet, LINE_TO_ROW(16), parsed.subtotalNonRq);
    sheet = writeSubtotalRow(sheet, LINE_TO_ROW(17), parsed.totalRow);

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
    rowCount: parsed.rows.length,
    parsed,
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
    process.argv[3] || "sales_for_resale.xlsx"
  );
  const result = fillSalesForResaleFromXbrl({
    xbrlPath,
    templatePath,
    baseDir: dir,
    outputFileName: process.argv[4]
      ? path.basename(process.argv[4])
      : undefined,
  });
  console.log("Company:", result.company);
  console.log("Wrote:", result.outputPath);
  console.log("Detail rows:", result.rowCount);
  console.log("Total revenue:", result.parsed.totalRow.total);
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
  parseSalesForResaleFromXbrl,
  fillSalesForResaleFromXbrl,
};
