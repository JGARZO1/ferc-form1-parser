/**
 * Fill about.xlsx from a FERC Form 1 XBRL (.xbrl / .xml) filing.
 *
 * Usage:
 *   node fill_about_info_xbrl.js [xbrl] [template] [outputName]
 */

const fs = require("fs");
const path = require("path");
const {
  loadXbrl,
  getStringFact,
  detectXbrlMeta,
  detectCompanyFromXbrl,
} = require("./xbrl_common");
const {
  companyFolderName,
  resolveCompanyDir,
  resolveTemplatePath,
} = require("./lib/company");
const { openTemplateCopy, setCellValue, setSheetName } = require("./lib/xlsx_utils");

function parseAboutFromXbrl(doc) {
  const pref = [];
  if (doc.primaryDuration) pref.push(doc.primaryDuration.id);

  const meta = detectXbrlMeta(doc);
  const address =
    getStringFact(doc, "AddressOfPrincipalOfficeAtEndOfPeriod", pref) ||
    getStringFact(doc, "PrincipalBusinessAddress", pref) ||
    "";
  const name_of_contact = getStringFact(doc, "NameOfContactPerson", pref) || "";
  const title_of_contact = getStringFact(doc, "TitleOfContactPerson", pref) || "";
  const address_of_contact =
    getStringFact(doc, "AddressOfContactPerson", pref) || address || "";
  const telephone_of_contact =
    getStringFact(doc, "TelephoneOfContactPerson", pref) ||
    getStringFact(doc, "TelephoneNumber", pref) ||
    "";
  const date_of_report =
    getStringFact(doc, "ReportDate", pref) ||
    getStringFact(doc, "AttestationDate", pref) ||
    getStringFact(doc, "DocumentDate", pref) ||
    "";

  return {
    utility_id: meta.utilityId || "",
    report_year: meta.reportYear,
    address,
    name_of_contact,
    title_of_contact,
    address_of_contact,
    telephone_of_contact,
    date_of_report,
  };
}

function fillAboutFromXbrl({
  xbrlPath,
  templatePath,
  baseDir,
  outputFileName,
  companyName,
  sheetName = "about",
}) {
  if (!fs.existsSync(xbrlPath)) throw new Error(`XBRL not found: ${xbrlPath}`);
  if (!fs.existsSync(templatePath))
    throw new Error(`Template not found: ${templatePath}`);

  const doc = loadXbrl(xbrlPath);
  const about = parseAboutFromXbrl(doc);
  if (about.report_year == null) {
    throw new Error("Could not determine report_year from XBRL");
  }

  const meta = detectXbrlMeta(doc, xbrlPath);
  const companyShort =
    companyName ||
    detectCompanyFromXbrl(doc, xbrlPath) ||
    companyFolderName(meta.respondent, "", xbrlPath);
  const root = baseDir || path.dirname(path.resolve(templatePath));
  const { companyDir, company, created: companyFolderCreated } =
    resolveCompanyDir(root, companyShort);

  const fileName =
    outputFileName ||
    `about_${about.utility_id || "UNKNOWN"}_${about.report_year}.xlsx`;
  const finalPath = path.join(companyDir, path.basename(fileName));

  const pack = openTemplateCopy(templatePath);
  try {
    const sheetPath = path.join(pack.unpacked, "xl", "worksheets", "sheet1.xml");
    let sheet = fs.readFileSync(sheetPath, "utf8");
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
    about,
  };
}

function main() {
  const dir = __dirname;
  const xbrlPath = path.resolve(
    dir,
    process.argv[2] || "filings/ClecoPowerLlc2020.xbrl"
  );
  const templatePath = resolveTemplatePath(dir, process.argv[3] || "about.xlsx");
  const outputArg = process.argv[4] || null;

  const result = fillAboutFromXbrl({
    xbrlPath,
    templatePath,
    baseDir: dir,
    outputFileName: outputArg ? path.basename(outputArg) : undefined,
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

module.exports = { parseAboutFromXbrl, fillAboutFromXbrl };
