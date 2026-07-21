/**
 * Fill sales_by_rate.xlsx All-Accounts totals from a FERC Form 1 XBRL filing.
 *
 * Lines:
 *   41 billed     – base sales-by-rate schedule totals (excl. unbilled)
 *   42 unbilled   – unbilled revenue concepts
 *   43 grand      – including unbilled revenue
 *
 * Usage:
 *   node fill_sales_by_rate_totals_xbrl.js [xbrl] [template] [outputName]
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

function metricsFromConcepts(doc, ctxId, {
  mwhConcept,
  revConcept,
  custConcept,
  kwhPerCustConcept,
  revPerKwhConcept,
}) {
  return {
    mwh_sold: getNumericFact(doc, mwhConcept, ctxId),
    revenue: getNumericFact(doc, revConcept, ctxId),
    avg_customers: getNumericFact(doc, custConcept, ctxId),
    kwh_per_customer: getNumericFact(doc, kwhPerCustConcept, ctxId),
    revenue_per_kwh: getNumericFact(doc, revPerKwhConcept, ctxId),
  };
}

/**
 * Parse All-Accounts totals from XBRL.
 */
function parseAllAccountsTotalsFromXbrl(doc) {
  const ctxId = doc.primaryDuration ? doc.primaryDuration.id : null;

  // Grand total (including unbilled) – often on unsegmented context
  const grand = metricsFromConcepts(doc, ctxId, {
    mwhConcept: "MegawattHoursOfElectricitySoldByRateSchedulesIncludingUnbilledRevenue",
    revConcept: "RevenueFromSalesOfElectricityByRateSchedulesIncludingUnbilledRevenue",
    custConcept:
      "AverageNumberOfCustomersPerMonthSalesOfElectricityByRateSchedulesIncludingUnbilledRevenue",
    kwhPerCustConcept:
      "AverageKilowattHoursOfSalesPerCustomerSalesOfElectricityByRateSchedulesIncludingUnbilledRevenue",
    revPerKwhConcept:
      "AverageRevenuePerKilowattHourSoldSalesOfElectricityByRateSchedulesIncludingUnbilledRevenue",
  });

  // Fallback grand: unsegmented RevenueFromSalesOfElectricityByRateSchedules
  if (grand.revenue == null) {
    grand.revenue = getNumericFact(
      doc,
      "RevenueFromSalesOfElectricityByRateSchedules",
      ctxId
    );
  }
  if (grand.mwh_sold == null) {
    grand.mwh_sold = getNumericFact(
      doc,
      "MegawattHoursSoldSalesOfElectricityByRateSchedules",
      ctxId
    );
  }
  if (grand.avg_customers == null) {
    grand.avg_customers = getNumericFact(
      doc,
      "AverageNumberOfCustomersPerMonthSalesOfElectricityByRateSchedules",
      ctxId
    );
  }
  if (grand.kwh_per_customer == null) {
    grand.kwh_per_customer = getNumericFact(
      doc,
      "AverageKilowattHoursOfSalesPerCustomerSalesOfElectricityByRateSchedules",
      ctxId
    );
  }
  if (grand.revenue_per_kwh == null) {
    grand.revenue_per_kwh = getNumericFact(
      doc,
      "AverageRevenuePerKilowattHourSoldSalesOfElectricityByRateSchedules",
      ctxId
    );
  }

  const unbilled = metricsFromConcepts(doc, ctxId, {
    mwhConcept: "MegawattHoursOfElectricitySoldByRateSchedulesUnbilled",
    revConcept: "RevenueFromSalesOfElectricityByRateSchedulesUnbilled",
    custConcept: null,
    kwhPerCustConcept: null,
    revPerKwhConcept:
      "AverageRevenuePerKilowattHourSoldSalesOfElectricityByRateSchedulesUnbilled",
  });
  // Alternate unbilled concept names
  if (unbilled.mwh_sold == null) {
    unbilled.mwh_sold = getNumericFact(
      doc,
      "MegawattHoursOfElectricitySoldUnbilled",
      ctxId
    );
  }
  if (unbilled.revenue == null) {
    unbilled.revenue = getNumericFact(
      doc,
      "RevenueFromSalesOfElectricityUnbilled",
      ctxId
    );
  }

  // Billed = grand − unbilled when both present; else use base (non-including) concepts
  const billedBase = metricsFromConcepts(doc, ctxId, {
    mwhConcept: "MegawattHoursSoldSalesOfElectricityByRateSchedules",
    revConcept: "RevenueFromSalesOfElectricityByRateSchedules",
    custConcept:
      "AverageNumberOfCustomersPerMonthSalesOfElectricityByRateSchedules",
    kwhPerCustConcept:
      "AverageKilowattHoursOfSalesPerCustomerSalesOfElectricityByRateSchedules",
    revPerKwhConcept:
      "AverageRevenuePerKilowattHourSoldSalesOfElectricityByRateSchedules",
  });

  const sub = (a, b) =>
    a != null && b != null ? a - b : a != null ? a : null;

  const billed = {
    mwh_sold:
      grand.mwh_sold != null && unbilled.mwh_sold != null
        ? grand.mwh_sold - unbilled.mwh_sold
        : billedBase.mwh_sold,
    revenue:
      grand.revenue != null && unbilled.revenue != null
        ? grand.revenue - unbilled.revenue
        : billedBase.revenue,
    avg_customers: billedBase.avg_customers ?? grand.avg_customers,
    kwh_per_customer: billedBase.kwh_per_customer ?? grand.kwh_per_customer,
    revenue_per_kwh: billedBase.revenue_per_kwh ?? grand.revenue_per_kwh,
  };

  // If grand empty but billed+unbilled exist, synthesize grand
  if (grand.revenue == null && billed.revenue != null) {
    grand.revenue = sub(
      billed.revenue + (unbilled.revenue || 0),
      0
    );
    if (billed.mwh_sold != null) {
      grand.mwh_sold = billed.mwh_sold + (unbilled.mwh_sold || 0);
    }
  }

  if (billed.revenue == null && grand.revenue == null) {
    throw new Error(
      "Could not find Sales by Rate Schedules All-Accounts totals in XBRL"
    );
  }

  return {
    billed: { line_no: 41, description: "TOTAL Billed - All Accounts", ...billed },
    unbilled: {
      line_no: 42,
      description: "TOTAL Unbilled Rev. - All Accounts",
      ...unbilled,
    },
    grand: { line_no: 43, description: "TOTAL - All Accounts", ...grand },
  };
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

function fillSalesByRateFromXbrl({
  xbrlPath,
  templatePath,
  baseDir,
  outputFileName,
  utilityId,
  reportYear,
  companyName,
  sheetName = "sales_by_rate",
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
  const totals = parseAllAccountsTotalsFromXbrl(doc);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `sales_by_rate_${resolvedUtilityId}_${resolvedYear}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const pack = openTemplateCopy(templatePath);
  try {
    const sheetPath = path.join(pack.unpacked, "xl", "worksheets", "sheet1.xml");
    let sheet = fs.readFileSync(sheetPath, "utf8");
    sheet = setCellValue(sheet, "B1", resolvedUtilityId, { asString: true });
    sheet = setCellValue(sheet, "B2", resolvedYear);
    sheet = fillMetrics(sheet, 4, totals.billed);
    sheet = fillMetrics(sheet, 5, totals.unbilled);
    sheet = fillMetrics(sheet, 6, totals.grand);
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
    totals,
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
    process.argv[3] || "sales_by_rate.xlsx"
  );
  const result = fillSalesByRateFromXbrl({
    xbrlPath,
    templatePath,
    baseDir: dir,
    outputFileName: process.argv[4]
      ? path.basename(process.argv[4])
      : undefined,
  });
  console.log("Company:", result.company);
  console.log("Wrote:", result.outputPath);
  for (const [k, row] of Object.entries(result.totals)) {
    console.log(
      `${k}: MWh=${row.mwh_sold} | Rev=${row.revenue} | Cust=${row.avg_customers}`
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
  parseAllAccountsTotalsFromXbrl,
  fillSalesByRateFromXbrl,
};
