/**
 * Company folder detection and resolution for FERC Form 1 pipeline.
 */

const fs = require("fs");
const path = require("path");

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

/**
 * Ordered company aliases (first match wins).
 * More specific Entergy entities before generic "Entergy".
 */
const COMPANY_ALIASES = [
  { test: /entergy\s+new\s+orleans/i, name: "ENTERGY NOLA" },
  { test: /entergy\s+louisiana/i, name: "ENTERGY LA" },
  { test: /\bcleco\b/i, name: "CLECO" },
  { test: /\bswepco\b|southwestern\s+electric\s+power/i, name: "SWEPCO" },
  { test: /\bentergy\b/i, name: "ENTERGY" },
];

/**
 * Map respondent / title / filename → short company folder name.
 */
function companyFolderName(respondent, title = "", filePath = "") {
  const LEGAL =
    /\b(LLC|L\.?L\.?C\.?|Inc\.?|Incorporated|Corp\.?|Corporation|Company|Co\.|Ltd\.?|Limited|LP|L\.P\.|PLC)\b/gi;
  const GENERIC =
    /^(Power|Electric|Electrical|Energy|Utility|Utilities|Gas|Transmission|Service|Services|Holding|Holdings|Group|of|and|the|d\/b\/a|New|Orleans|Louisiana)$/i;

  const sources = [
    respondent,
    title,
    path.basename(filePath, path.extname(filePath)),
  ]
    .map((s) => cleanText(s))
    .filter(Boolean);

  for (const src of sources) {
    for (const a of COMPANY_ALIASES) {
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
  return brand.replace(/[^A-Za-z0-9 _-]/g, "").trim().toUpperCase() || "UNKNOWN";
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

/**
 * Resolve a template path: prefer templates/<name>, then root/<name>.
 */
function resolveTemplatePath(baseDir, templateName) {
  const candidates = [
    path.join(baseDir, "templates", templateName),
    path.join(baseDir, templateName),
  ];
  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  return candidates[0];
}

module.exports = {
  cleanText,
  COMPANY_ALIASES,
  companyFolderName,
  resolveCompanyDir,
  resolveTemplatePath,
};
