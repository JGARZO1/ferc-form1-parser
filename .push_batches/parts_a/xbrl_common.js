/**
 * Shared utilities for parsing FERC Form 1 XBRL instance documents.
 *
 * Supports classic XBRL fact form:
 *   <ferc:ConceptName contextRef="c-01" unitRef="u-02" decimals="0">123</ferc:ConceptName>
 *
 * Also tolerates alternate prefixes and simple context graphs (duration/instant + segments).
 */

const fs = require("fs");
const path = require("path");
const { cleanText, companyFolderName } = require("./lib/company");

function decodeXmlEntities(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) =>
      String.fromCharCode(parseInt(h, 16))
    );
}

function parseNumber(value) {
  let text = cleanText(decodeXmlEntities(value));
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

/**
 * Parse contexts from an XBRL instance.
 * @returns {Map<string, object>}
 */
function parseContexts(xml) {
  const contexts = new Map();
  const re =
    /<context\b[^>]*\bid="([^"]+)"[^>]*>([\s\S]*?)<\/context>/gi;
  let m;
  while ((m = re.exec(xml))) {
    const id = m[1];
    const body = m[2];

    const identM = body.match(
      /<identifier\b[^>]*>([\s\S]*?)<\/identifier>/i
    );
    const identifier = identM ? cleanText(decodeXmlEntities(identM[1])) : "";

    let periodType = "unknown";
    let startDate = null;
    let endDate = null;
    let instant = null;
    const instantM = body.match(/<instant>([^<]+)<\/instant>/i);
    if (instantM) {
      periodType = "instant";
      instant = cleanText(instantM[1]);
    } else {
      const startM = body.match(/<startDate>([^<]+)<\/startDate>/i);
      const endM = body.match(/<endDate>([^<]+)<\/endDate>/i);
      if (startM || endM) {
        periodType = "duration";
        startDate = startM ? cleanText(startM[1]) : null;
        endDate = endM ? cleanText(endM[1]) : null;
      }
    }

    const segments = [];
    const segRe =
      /<(?:xbrldi:)?(?:explicitMember|typedMember)\b([^>]*)>([\s\S]*?)<\/(?:xbrldi:)?(?:explicitMember|typedMember)>/gi;
    let sm;
    while ((sm = segRe.exec(body))) {
      const attrs = sm[1] || "";
      const inner = sm[2] || "";
      const dimM = attrs.match(/dimension="([^"]+)"/i);
      const dimension = dimM ? dimM[1].replace(/^ferc:/, "") : "";
      const member = cleanText(
        decodeXmlEntities(inner.replace(/<[^>]+>/g, " "))
      );
      // typedMember often has nested element text like 0-1
      const nested = inner.match(/>([^<]+)</);
      segments.push({
        dimension,
        member: nested ? cleanText(decodeXmlEntities(nested[1])) : member,
        raw: cleanText(decodeXmlEntities(inner)),
      });
    }

    contexts.set(id, {
      id,
      identifier,
      periodType,
      startDate,
      endDate,
      instant,
      segments,
      hasSegment: segments.length > 0,
    });
  }
  return contexts;
}

/**
 * Parse all ferc:* facts from an XBRL instance.
 * @returns {Array<{concept: string, contextRef: string, unitRef: string|null, value: string, numeric: number|null}>}
 */
function parseFacts(xml) {
  const facts = [];
  // Require ferc: prefix so we do not scan the entire non-FERC XML tree.
  // Form: <ferc:ConceptName contextRef="c-01" ...>value</ferc:ConceptName>
  const re =
    /<ferc:([A-Za-z][A-Za-z0-9_]*)\b([^>]*)>([\s\S]*?)<\/ferc:\1>/g;
  let m;
  while ((m = re.exec(xml))) {
    const concept = m[1];
    const attrs = m[2] || "";
    // Facts always carry contextRef; skip domain/member placeholders without one
    if (!/\bcontextRef=/.test(attrs)) continue;

    const ctxM = attrs.match(/\bcontextRef="([^"]+)"/i);
    const unitM = attrs.match(/\bunitRef="([^"]+)"/i);
    const raw = cleanText(decodeXmlEntities(m[3].replace(/<[^>]+>/g, " ")));
    facts.push({
      concept,
      contextRef: ctxM ? ctxM[1] : "",
      unitRef: unitM ? unitM[1] : null,
      value: raw,
      numeric: parseNumber(raw),
    });
  }
  return facts;
}

/**
 * Load and index an XBRL instance document.
 */
function loadXbrl(filePath) {
  const xml = fs.readFileSync(filePath, "utf8");
  const contexts = parseContexts(xml);
  const facts = parseFacts(xml);

  // Index: concept → facts[]
  const byConcept = new Map();
  for (const f of facts) {
    if (!byConcept.has(f.concept)) byConcept.set(f.concept, []);
    byConcept.get(f.concept).push(f);
  }

  // Identify primary duration context (report year, no segment)
  const durationNoSeg = [...contexts.values()].filter(
    (c) => c.periodType === "duration" && !c.hasSegment
  );
  // Prefer longest calendar-year span ending in report year
  durationNoSeg.sort((a, b) => {
    const ay = a.endDate || "";
    const by = b.endDate || "";
    return by.localeCompare(ay);
  });

  const primaryDuration =
    durationNoSeg.find((c) => {
      if (!c.startDate || !c.endDate) return false;
      // full year roughly
      return c.startDate.slice(5) === "01-01" && c.endDate.slice(5) === "12-31";
    }) || durationNoSeg[0] || null;

  const reportYearFromCtx = primaryDuration?.endDate
    ? Number(primaryDuration.endDate.slice(0, 4))
    : null;

  // Prior year duration (for O&M previous year)
  const priorDuration =
    durationNoSeg.find((c) => {
      if (!primaryDuration || !c.endDate || !primaryDuration.endDate) return false;
      const y = Number(c.endDate.slice(0, 4));
      const py = Number(primaryDuration.endDate.slice(0, 4));
      return (
        y === py - 1 &&
        c.startDate &&
        c.startDate.slice(5) === "01-01" &&
        c.endDate.slice(5) === "12-31"
      );
    }) || null;

  // Instant contexts: end of report year / beginning (prior year-end)
  const instants = [...contexts.values()].filter(
    (c) => c.periodType === "instant" && !c.hasSegment
  );
  const eoyInstant =
    instants.find(
      (c) =>
        reportYearFromCtx &&
        c.instant === `${reportYearFromCtx}-12-31`
    ) ||
    instants
      .filter((c) => c.instant && c.instant.endsWith("-12-31"))
      .sort((a, b) => (b.instant || "").localeCompare(a.instant || ""))[0] ||
    null;

  const boyInstant =
    (reportYearFromCtx &&
      instants.find((c) => c.instant === `${reportYearFromCtx - 1}-12-31`)) ||
    null;

  return {
    filePath: path.resolve(filePath),
    xml,
    contexts,
    facts,
    byConcept,
    primaryDuration,
    priorDuration,
    eoyInstant,
    boyInstant,
    reportYearFromCtx,
  };
}

/**
 * First non-empty string fact for concept (any context, prefer unsegmented).
 */
function getStringFact(doc, concept, preferredContextIds = []) {
  const list = doc.byConcept.get(concept) || [];
  if (!list.length) return "";

  for (const id of preferredContextIds) {
    const hit = list.find((f) => f.contextRef === id && f.value);
    if (hit) return hit.value;
  }

  // Prefer unsegmented contexts
  const unseg = list.find((f) => {
    const c = doc.contexts.get(f.contextRef);
    return f.value && c && !c.hasSegment;
  });
  if (unseg) return unseg.value;

  const any = list.find((f) => f.value);
  return any ? any.value : "";
}

/**
 * Numeric fact for concept in a specific context (or first unsegmented).
 */
function getNumericFact(doc, concept, contextId = null) {
  const list = doc.byConcept.get(concept) || [];
  if (!list.length) return null;

  if (contextId) {
    const hit = list.find((f) => f.contextRef === contextId);
    if (hit) return hit.numeric;
  }

  const unseg = list.find((f) => {
    const c = doc.contexts.get(f.contextRef);
    return c && !c.hasSegment && f.numeric != null;
  });
  if (unseg) return unseg.numeric;

  const any = list.find((f) => f.numeric != null);
  return any ? any.numeric : null;
}

/**
 * All facts for concept filtered by optional segment dimension.
 */
function getFactsForConcept(doc, concept, { requireSegment = null } = {}) {
  const list = doc.byConcept.get(concept) || [];
  return list.filter((f) => {
    const c = doc.contexts.get(f.contextRef);
    if (!c) return false;
    if (requireSegment === false && c.hasSegment) return false;
    if (requireSegment === true && !c.hasSegment) return false;
    return true;
  });
}

/**
 * Detect utility_id, report_year, respondent from XBRL.
 */
function detectXbrlMeta(doc, filePath = "") {
  const pref = [];
  if (doc.primaryDuration) pref.push(doc.primaryDuration.id);

  const utilityId =
    getStringFact(doc, "CompanyIdentifier", pref) ||
    getStringFact(doc, "RespondentIdentifier", pref) ||
    getStringFact(doc, "FilerIdentifier", pref) ||
    "";

  let reportYear = null;
  const yearRaw = getStringFact(doc, "ReportYear", pref);
  if (yearRaw) reportYear = Number(yearRaw);
  if (reportYear == null || Number.isNaN(reportYear)) {
    reportYear = doc.reportYearFromCtx;
  }
  if (reportYear == null || Number.isNaN(reportYear)) {
    const base = path.basename(filePath || doc.filePath, path.extname(filePath || doc.filePath));
    const ym = base.match(/(19|20)\d{2}/);
    if (ym) reportYear = Number(ym[0]);
  }

  const respondent =
    getStringFact(doc, "RespondentLegalName", pref) ||
    getStringFact(doc, "EntityName", pref) ||
    "";

  return {
    utilityId,
    reportYear: Number.isFinite(reportYear) ? reportYear : null,
    respondent,
  };
}

/**
 * Company folder name from XBRL meta + filename.
 */
function detectCompanyFromXbrl(doc, filePath) {
  const meta = detectXbrlMeta(doc, filePath);
  return companyFolderName(meta.respondent, "", filePath);
}

/**
 * Group dimensional facts by segment member key for a given axis-like dimension name.
 * Returns Map<memberKey, { contextId, facts: Map<concept, fact> }>
 */
function groupBySegmentMember(doc, dimensionNameHint) {
  const hint = (dimensionNameHint || "").toLowerCase();
  const groups = new Map();

  for (const [id, ctx] of doc.contexts) {
    if (!ctx.hasSegment) continue;
    const seg =
      ctx.segments.find((s) =>
        s.dimension.toLowerCase().includes(hint)
      ) || ctx.segments[0];
    if (!seg) continue;
    const key = seg.member || seg.raw;
    if (!groups.has(key)) {
      groups.set(key, { contextId: id, member: key, dimension: seg.dimension, facts: new Map() });
    }
  }

  // Attach facts
  for (const f of doc.facts) {
    const ctx = doc.contexts.get(f.contextRef);
    if (!ctx || !ctx.hasSegment) continue;
    const seg =
      ctx.segments.find((s) =>
        s.dimension.toLowerCase().includes(hint)
      ) || ctx.segments[0];
    if (!seg) continue;
    const key = seg.member || seg.raw;
    if (!groups.has(key)) {
      groups.set(key, {
        contextId: f.contextRef,
        member: key,
        dimension: seg.dimension,
        facts: new Map(),
      });
    }
    const g = groups.get(key);
    // Prefer matching context
    if (g.contextId === f.contextRef || !g.facts.has(f.concept)) {
      g.facts.set(f.concept, f);
    }
  }

  return groups;
}

module.exports = {
  decodeXmlEntities,
  parseNumber,
  parseContexts,
  parseFacts,
  loadXbrl,
  getStringFact,
  getNumericFact,
  getFactsForConcept,
  detectXbrlMeta,
  detectCompanyFromXbrl,
  groupBySegmentMember,
};
