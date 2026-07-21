/**
 * Master runner: process any FERC Form 1 HTML or XBRL filing through all extractors.
 *
 * Usage:
 *   node run_all_ferc.js <path-to-filing>
 *   node run_all_ferc.js ClecoPowerLlc2023.html
 *   node run_all_ferc.js filings/ClecoPowerLlc2020.xbrl
 *
 * Auto-detects format from extension (.html / .htm → HTML parsers;
 * .xbrl / .xml → XBRL parsers). Detects utility_id, report_year, and
 * company short name; reuses/creates the company folder; runs five schedules.
 */

const fs = require("fs");
const path = require("path");

const fillAbout = require("./fill_about_info.js");
const fillSalesByRate = require("./fill_sales_by_rate_totals.js");
const fillSalesForResale = require("./fill_sales_for_resale.js");
const fillPlantInService = require("./fill_plant_in_service.js");
const fillOmExpenses = require("./fill_om_expenses.js");

const fillAboutXbrl = require("./fill_about_info_xbrl.js");
const fillSalesByRateXbrl = require("./fill_sales_by_rate_totals_xbrl.js");
const fillSalesForResaleXbrl = require("./fill_sales_for_resale_xbrl.js");
const fillPlantInServiceXbrl = require("./fill_plant_in_service_xbrl.js");
const fillOmExpensesXbrl = require("./fill_om_expenses_xbrl.js");

const {
  COMPANY_ALIASES,
  companyFolderName,
  resolveCompanyDir,
  resolveTemplatePath,
  cleanText,
} = require("./lib/company");

const BASE_DIR = __dirname;

const PIPELINE_HTML = [
  {
    key: "about",
    label: "About / contact info",
    template: "about.xlsx",
    filePrefix: "about",
    run: (opts) => fillAbout.fillAboutFromHtml(opts),
  },
  {
    key: "sales_by_rate",
    label: "Sales by rate schedules",
    template: "sales_by_rate.xlsx",
    filePrefix: "sales_by_rate",
    run: (opts) => fillSalesByRate.fillSalesByRateFromHtml(opts),
  },
  {
    key: "sales_for_resale",
    label: "Sales for resale",
    template: "sales_for_resale.xlsx",
    filePrefix: "sales_for_resale",
    run: (opts) => fillSalesForResale.fillSalesForResaleFromHtml(opts),
  },
  {
    key: "plant_in_service",
    label: "Electric plant in service",
    template: "plant_in_service.xlsx",
    filePrefix: "plant_in_service",
    run: (opts) => fillPlantInService.fillPlantInServiceFromHtml(opts),
  },
  {
    key: "om_expenses",
    label: "O&M expenses",
    template: "om_expenses.xlsx",
    filePrefix: "om_expenses",
    run: (opts) => fillOmExpenses.fillOmExpensesFromHtml(opts),
  },
];

const PIPELINE_XBRL = [
  {
    key: "about",
    label: "About / contact info",
    template: "about.xlsx",
    filePrefix: "about",
    run: (opts) =>
      fillAboutXbrl.fillAboutFromXbrl({
        ...opts,
        xbrlPath: opts.htmlPath || opts.xbrlPath,
      }),
  },
  {
    key: "sales_by_rate",
    label: "Sales by rate schedules",
    template: "sales_by_rate.xlsx",
    filePrefix: "sales_by_rate",
    run: (opts) =>
      fillSalesByRateXbrl.fillSalesByRateFromXbrl({
        ...opts,
        xbrlPath: opts.htmlPath || opts.xbrlPath,
      }),
  },
  {
    key: "sales_for_resale",
    label: "Sales for resale",
    template: "sales_for_resale.xlsx",
    filePrefix: "sales_for_resale",
    run: (opts) =>
      fillSalesForResaleXbrl.fillSalesForResaleFromXbrl({
        ...opts,
        xbrlPath: opts.htmlPath || opts.xbrlPath,
      }),
  },
  {
    key: "plant_in_service",
    label: "Electric plant in service",
    template: "plant_in_service.xlsx",
    filePrefix: "plant_in_service",
    run: (opts) =>
      fillPlantInServiceXbrl.fillPlantInServiceFromXbrl({
        ...opts,
        xbrlPath: opts.htmlPath || opts.xbrlPath,
      }),
  },
  {
    key: "om_expenses",
    label: "O&M expenses",
    template: "om_expenses.xlsx",
    filePrefix: "om_expenses",
    run: (opts) =>
      fillOmExpensesXbrl.fillOmExpensesFromXbrl({
        ...opts,
        xbrlPath: opts.htmlPath || opts.xbrlPath,
      }),
  },
];

// ---------------------------------------------------------------------------
// Detection helpers (HTML iXBRL)
// ---------------------------------------------------------------------------

function stripTags(html) {
  return cleanText(String(html || "").replace(/<[^>]+>/g, " "));
}

function grabIxFact(html, concept) {
  const name = concept.startsWith("ferc:") ? concept : `ferc:${concept}`;
  const re = new RegExp(
    `name="${name.replace(/:/g, "\\:")}"[^>]*>([\\s\\S]*?)</ix:`,
    "i"
  );
  const m = html.match(re);
  return m ? stripTags(m[1]) : "";
}

function detectFilingMetaHtml(html, htmlPath) {
  const utilityId =
    grabIxFact(html, "CompanyIdentifier") ||
    grabIxFact(html, "RespondentIdentifier") ||
    grabIxFact(html, "FilerIdentifier") ||
    "";

  const yearRaw = grabIxFact(html, "ReportYear");
  let reportYear = yearRaw ? Number(yearRaw) : null;

  if (reportYear == null || Number.isNaN(reportYear)) {
    const base = path.basename(htmlPath, path.extname(htmlPath));
    const ym = base.match(/(19|20)\d{2}/);
    if (ym) reportYear = Number(ym[0]);
  }

  const respondent =
    grabIxFact(html, "RespondentLegalName") ||
    grabIxFact(html, "EntityName") ||
    "";

  const titleM = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const title = titleM ? stripTags(titleM[1]) : "";

  return { utilityId, reportYear, respondent, title };
}

function detectFormat(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".html" || ext === ".htm") return "html";
  if (ext === ".xbrl" || ext === ".xml") return "xbrl";
  // Peek at content
  try {
    const head = fs.readFileSync(filePath, "utf8").slice(0, 2000);
    if (/<html[\s>]/i.test(head) || /ix:nonFraction|ix:nonNumeric/i.test(head)) {
      return "html";
    }
    if (/<xbrl[\s>]/i.test(head) || /xmlns:ferc=/i.test(head)) {
      return "xbrl";
    }
  } catch {
    /* ignore */
  }
  throw new Error(
    `Unrecognized filing format for ${filePath} (expected .html, .xbrl, or .xml)`
  );
}

// ---------------------------------------------------------------------------
// Pipeline
// ---------------------------------------------------------------------------

/**
 * Process one FERC Form 1 filing through all extractors.
 */
function runAllFerc(filingPath, { baseDir = BASE_DIR } = {}) {
  const resolved = path.resolve(filingPath);
  if (!fs.existsSync(resolved)) {
    throw new Error(`Filing not found: ${resolved}`);
  }

  const format = detectFormat(resolved);
  const pipeline = format === "xbrl" ? PIPELINE_XBRL : PIPELINE_HTML;

  let meta;
  if (format === "html") {
    const html = fs.readFileSync(resolved, "utf8");
    meta = detectFilingMetaHtml(html, resolved);
  } else {
    const { loadXbrl, detectXbrlMeta } = require("./xbrl_common");
    const doc = loadXbrl(resolved);
    const m = detectXbrlMeta(doc, resolved);
    meta = {
      utilityId: m.utilityId,
      reportYear: m.reportYear,
      respondent: m.respondent,
      title: "",
    };
  }

  if (!meta.utilityId) {
    console.warn("Warning: utility_id not found; using UNKNOWN");
  }
  if (meta.reportYear == null || Number.isNaN(meta.reportYear)) {
    throw new Error(
      "Could not determine report_year from filing or filename"
    );
  }

  const companyShort = companyFolderName(
    meta.respondent,
    meta.title,
    resolved
  );
  const { companyDir, company, created } = resolveCompanyDir(
    baseDir,
    companyShort
  );

  const utilityId = meta.utilityId || "UNKNOWN";
  const reportYear = meta.reportYear;

  console.log("══════════════════════════════════════════════════");
  console.log(" FERC Form 1 batch extract");
  console.log("══════════════════════════════════════════════════");
  console.log(` Filing:      ${resolved}`);
  console.log(` Format:      ${format.toUpperCase()}`);
  console.log(` Respondent:  ${meta.respondent || "(n/a)"}`);
  console.log(` Utility ID:  ${utilityId}`);
  console.log(` Report year: ${reportYear}`);
  console.log(
    ` Company:     ${company}${created ? " [folder created]" : " [existing folder]"}`
  );
  console.log(` Output dir:  ${companyDir}`);
  console.log("──────────────────────────────────────────────────");

  const results = [];
  const createdFiles = [];
  const errors = [];

  for (const step of pipeline) {
    const templatePath = resolveTemplatePath(baseDir, step.template);
    const outputFileName = `${step.filePrefix}_${utilityId}_${reportYear}.xlsx`;
    const expectedPath = path.join(companyDir, outputFileName);

    process.stdout.write(` → ${step.label} ... `);

    if (!fs.existsSync(templatePath)) {
      const msg = `template missing: ${step.template}`;
      console.log(`SKIP (${msg})`);
      errors.push({ step: step.key, error: msg });
      results.push({
        key: step.key,
        label: step.label,
        ok: false,
        skipped: true,
        error: msg,
      });
      continue;
    }

    try {
      const out = step.run({
        htmlPath: resolved,
        xbrlPath: resolved,
        templatePath,
        baseDir,
        outputFileName,
        companyName: company,
        utilityId,
        reportYear,
      });

      const outputPath = out.outputPath || expectedPath;
      const exists = fs.existsSync(outputPath);
      console.log(exists ? "OK" : "OK (path?)");
      if (exists) createdFiles.push(outputPath);

      results.push({
        key: step.key,
        label: step.label,
        ok: true,
        outputPath,
        result: out,
      });
    } catch (err) {
      const message = err && err.message ? err.message : String(err);
      console.log(`FAIL`);
      console.log(`   ${message}`);
      errors.push({ step: step.key, error: message });
      results.push({
        key: step.key,
        label: step.label,
        ok: false,
        error: message,
      });
    }
  }

  console.log("──────────────────────────────────────────────────");
  console.log(" SUMMARY");
  console.log("──────────────────────────────────────────────────");
  console.log(` Company folder: ${companyDir}`);
  console.log(` Year:           ${reportYear}`);
  console.log(` Utility ID:     ${utilityId}`);
  console.log(` Format:         ${format}`);
  console.log(
    ` Steps ok:       ${results.filter((r) => r.ok).length}/${pipeline.length}`
  );
  if (createdFiles.length) {
    console.log(" Files written:");
    for (const f of createdFiles) {
      console.log(`   • ${path.basename(f)}`);
    }
  }
  if (errors.length) {
    console.log(" Errors:");
    for (const e of errors) {
      console.log(`   • ${e.step}: ${e.error}`);
    }
  }
  console.log("══════════════════════════════════════════════════");

  return {
    filingPath: resolved,
    format,
    companyDir,
    company,
    companyFolderCreated: created,
    utilityId,
    reportYear,
    respondent: meta.respondent,
    results,
    createdFiles,
    errors,
  };
}

function main() {
  const arg = process.argv[2];
  if (!arg) {
    console.error("Usage: node run_all_ferc.js <path-to-html-or-xbrl>");
    console.error("Example: node run_all_ferc.js ClecoPowerLlc2023.html");
    console.error("Example: node run_all_ferc.js filings/ClecoPowerLlc2020.xbrl");
    process.exit(1);
  }

  try {
    const summary = runAllFerc(arg, { baseDir: BASE_DIR });
    if (summary.errors.length) process.exitCode = 1;
  } catch (err) {
    console.error(err.message || err);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  runAllFerc,
  detectFormat,
  detectFilingMetaHtml,
  companyFolderName,
  resolveCompanyDir,
  COMPANY_ALIASES,
  PIPELINE_HTML,
  PIPELINE_XBRL,
};
