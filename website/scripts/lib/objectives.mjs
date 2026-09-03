/**
 * Every objective of the invoice verification harness, as data, plus the pure
 * functions that decide them. No I/O happens here — the runner reads the
 * measure files and passes them in; this module only judges. That split is
 * what lets the runner hash this file (invariant I2): a threshold can only
 * change by changing bytes the history records.
 *
 * Shapes (see VERIFY-CONTRACT.md "Probe outputs"; keys beyond the contract
 * samples that objectives require are listed here so the probes know what to
 * emit — the contract allows the shapes to grow but not to shift):
 *
 *   m = { a, b, c, v, build }
 *
 *   a  (measure-a.json, OOXML facts) — contract sample, plus:
 *        footerParts[part].drawingWithTextParas   count of w:p with both a
 *                                                 w:drawing and text (F4.A1)
 *        footerParts[part].drawingWithTextSample  text of that paragraph
 *        bodyTables[].rows[].cells[].noteSpacing  {line} of the muted second
 *                                                 paragraph in a desc cell (F7.A1)
 *        cardTable       {borderColor, tcMarTwips, widthTwips, heightTwips} (F13)
 *        pageBreaks.maxConsecutiveEmptyParas      longest run of empty w:p (F12.A2)
 *        borderFixture   {emitsTcBorders} from the minimal borderSides:[] fixture (F9.A2)
 *        variants        {publish?, lines3?, lines30?, unresolvedImage?} — each a
 *                        measure-a-like object for the flagged builds (F8.A2/A3, F13.A3)
 *   b  (measure-b.json, px @96dpi, page-relative) —
 *        regions         {name: {x, y, w, h, section}}
 *        sections        [{w, h, headerHeight, footerHeight, pageNumberText,
 *                          pageNumberRightPx, trailingEmptyPara}]
 *        chargeRowHeights[]
 *        fonts           {amountCell, descriptionCell}  computed font-family
 *        fontProbe       {aptosWidth, monoWidth}
 *        variants        {unresolvedImage?: {regions}} (F13.A3 no-reflow)
 *   c  (measure-c.json, points, page-relative) —
 *        pages, firstPageHeaderText, primaryHeaderText
 *        regions         {name: {page, x, y, xEnd?}}
 *        suppressedHeadingFinds  {"Invoice details": bool, ...} (F1.C1)
 *        footerTable     {xPt, widthPt, width} (F3.C1)
 *        footer          {heightPt, lineCount} (F4.C1)
 *        pageNumberParagraph  {alignment, spaceAfter} (F5.C1)
 *        fieldsTextByPage[]   footer text per page after Fields.Update() (F5.C2/C6.C1)
 *        fonts           {amountCell: {name}, heading1: {spacingPt}} (F6/F10)
 *        amountRightEdgesPt[] (F6.C2)
 *        charges         {topPt, bottomPt} (F7.C1)
 *        chargeRowLineCounts[] (F7.C2)
 *        chargesCol1Shading[8] Word RGB ints, rows 1..8 col 1 (F8.C1)
 *        chargesRow3BottomLineStyle (F9.C1)
 *        vAlign          {bandCells: [], footerCells: []} (F11.C1)
 *        summary         {rightIndentPt, maxLineWidthPt} (F14.C1)
 *        scanCard        {widthPt, heightPt, inlineShapeCount} (F13.C1, D14.C1)
 *        shapesCount, inlineShapesCount (F15.C2)
 *   v  (measure-v.json) —
 *        regions         [{region, pctDiff, pixels}] (G4)
 *        cornerTests     {preview: [{region, cornersOk|corners[], centerOk}],
 *                         word: [...]} (F15.B1/C1)
 *   build — {publish?: {ok, error?, documentJson?}, fullChecks?: {testsOk,
 *        documentsCheckOk}, invariants?: {I1..I4}, aptosRegistered?}
 *
 * Every evaluate returns {status: 'PASS'|'FAIL'|'BLOCKED', measured, expected,
 * note?} and never throws — a missing probe file is FAIL "probe missing", a
 * missing key is FAIL "not found"/"not measured". BLOCKED is reserved for the
 * declared dependencies (publish-build, P0), matching the contract.
 */
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

/** Geometry ground truth (contract: "The region table"). Loaded once. */
const DESIGN = require("../../design/invoice-v2/fixtures/design-regions.json");

// ---------------------------------------------------------------------------
// Thresholds. Every number an objective compares against lives here, named,
// with its source. Changing one changes this file's hash (I2) — deliberately.
// ---------------------------------------------------------------------------

/** pt = px * 0.75 — contract "The region table" conversions. */
export const PX_TO_PT = 0.75;
/** px per mm at 96dpi — design-regions.json meta.pxPerMm. */
export const PX_PER_MM = DESIGN.meta.pxPerMm; // 3.77953
/** pt per mm (72 / 25.4) — contract tolerance derivations. */
export const PT_PER_MM = 2.8346;
/** ±1mm in preview px — contract region table ("±1mm = ±3.78px"). */
const TOL_MM_PX = 3.78;
/** ±1mm in Word points — contract region table ("±1mm = ±2.83pt"). */
const TOL_MM_PT = 2.83;
/** Chrome font probe: Aptos vs mono fallback must differ by more than this
 * many px (task/contract P0; measured real gap is 32.7px). */
const FONT_PROBE_MIN_DELTA_PX = 5;
/** Word reports exactly this many pages — plan gate G1. */
const PAGES_EXPECTED = 2;
/** Preview section 1 height range, px — contract G2. */
const G2_SECTION1_PX = [1000, 1124];
/** Per-visual-region pixel diff ceiling, percent — contract G4. */
const G4_MAX_PCT = 3;
/** Regions exempt from the pixel diff — contract G4 / region table
 * (scan-card is geometry only; the canvas QR is random modules). */
const G4_EXEMPT = new Set(["scan-card", "qr-canvas"]);

/** Headings the design suppresses / keeps — plan F1. */
const SUPPRESSED_HEADINGS = ["Invoice details", "Parties", "Charges"];
const KEPT_HEADINGS = ["Engagement summary", "Pay by bank transfer", "Terms & notes", "Scan to pay"];
/** Sections the built model must keep — contract measure-a sample
 * (sectionsInModel: 7; verified against the current built model.json). */
const SECTIONS_IN_MODEL = 7;
/** Band starts at most this far below the rule, px — plan F1.B. */
const BAND_GAP_MAX_PX = 3;
/** Rounding slack for that gap (an overlap this small is measurement noise). */
const BAND_GAP_MIN_PX = -1;

/** First-page header max height / primary header range, px — plan F2.B. */
const FIRST_HEADER_MAX_PX = 2;
const PRIMARY_HEADER_PX = [40, 80];
/** The invoice reference the running header must carry — plan F2. */
const INVOICE_REFERENCE = "INV-2026-0142";

/** Footer bar bleed: tblInd and full-page width in twips — plan F3.A. */
const FOOTER_TBLIND_TWIPS = -907;
const PAGE_WIDTH_TWIPS = 11906;
/** Footer table position/width in points — plan F3.C. */
const FOOTER_X_PT = 0;
const FOOTER_WIDTH_PT = 595.3;
const FOOTER_TABLE_TOL_PT = 1;
/** Preview footer-table edge tolerance, px — plan F3.B. */
const FOOTER_EDGE_TOL_PX = 1;

/** Footer element height range, px — contract F4.B (design 54.78). */
const FOOTER_HEIGHT_PX = [52, 58];
/** Footer height ceiling, mm, and line count — contract F4.C. */
const FOOTER_MAX_MM = 14.8;
const FOOTER_LINES = 1;
/** The footer text the mark shares a paragraph with — plan F4.A. */
const FOOTER_MARK_TEXT = "Generated with Docxcelerate";

/** Page-number right edge, px page-relative, and tolerance — contract F5.B
 * (design footer text right inset x = 733.2px). */
const PAGE_NUM_RIGHT_PX = 733.2;
const PAGE_NUM_RIGHT_TOL_PX = 2;
/** wdAlignParagraphRight — contract environment facts. */
const WD_ALIGN_RIGHT = 2;
/** Footer field text per page — plan F5.C / C6. */
const PAGE_FIELD_TEXTS = ["1 / 2", "2 / 2"];

/** The tabular face for money columns — plan F6. */
const MONEY_FONT = "Consolas";
/** Amount right edges must agree within this, mm — plan F6.C ("0.35mm"). */
const AMOUNT_EDGE_TOL_MM = 0.35;

/** Body default line value the charge rows must undercut, twentieths — plan
 * F7.A / contract F7.A (1.45 x 240 = 348). */
const BODY_LINE_TWIPS = 348;
/** Charge row height range, px — contract F7.B (design 54.11). */
const CHARGE_ROW_PX = [50, 58];
/** charges-body height range, px — contract F7.B (design 378.77). */
const CHARGES_BODY_PX = [370, 388];
/** Charges head+body ceiling, mm — contract F7.C (design 410.6px = 108.7mm). */
const CHARGES_TOTAL_MAX_MM = 111;
/** Max rendered lines per charge row — plan F7.C ("no third line"). */
const CHARGE_ROW_MAX_LINES = 2;

/** Zebra fill — plan F8 (F7F8FD), and as a Word RGB int (r+g*256+b*65536). */
const ZEBRA_FILL = "F7F8FD";
const ZEBRA_RGB = 247 + 248 * 256 + 253 * 65536; // 16644343
/** Accent fill — contract measure-a sample (2C3D8F), as a Word RGB int. */
const ACCENT_FILL = "2C3D8F";
const ACCENT_RGB = 44 + 61 * 256 + 143 * 65536; // 9387308
/** wdColorAutomatic — an unshaded cell — Word object model constant. */
const WD_COLOR_AUTOMATIC = -16777216;

/** The row hairline colour that must vanish — plan F9 (D9DDEB). */
const RULE_COLOR = "D9DDEB";
/** wdLineStyleNone — contract environment facts. */
const WD_LINE_STYLE_NONE = 0;

/** Heading1 tracking, twentieths of a point — plan F10.A (0.12em x 7.5pt x 20). */
const HEADING1_TRACK_TWENTIETHS = 18;
/** Heading1 Font.Spacing in Word, pt — plan F10.C. */
const HEADING1_SPACING_PT = 0.9;
const HEADING1_SPACING_TOL_PT = 0.05;
/** ENGAGEMENT SUMMARY rendered width tolerance, mm — plan F10.C. */
const LABEL_WIDTH_TOL_MM = 0.5;

/** wdCellAlignVerticalCenter — contract environment facts. */
const WD_VALIGN_CENTER = 1;
/** Pill/value vertical-centre tolerance, px — plan F11.B. */
const PILL_CENTER_TOL_PX = 1;

/** Card border margin, twips — plan F13.A (12pt x 20). */
const CARD_TCMAR_TWIPS = 240;
/** QR box and card width in preview, px — contract F13.B. */
const QR_BOX_PX = 143.6;
const QR_BOX_TOL_PX = 4;
const CARD_WIDTH_PX = 234.3;
const CARD_WIDTH_TOL_PX = 4;
/** Card table in Word, pt — plan F13.C (contract: "card ~= 108pt wide"). */
const CARD_WIDTH_PT = 108;
const CARD_WIDTH_TOL_PT = 1;
const CARD_HEIGHT_TOL_PT = 2;
/** No-reflow: the following region moves by exactly this much — plan F13.A3. */
const NO_REFLOW_TOL_PX = 0;

/** Summary right indent, twips / pt — plan F14.A/C (20mm = 1134 twips). */
const SUMMARY_IND_RIGHT_TWIPS = 1134;
const SUMMARY_IND_RIGHT_PT = 56.7;
const SUMMARY_IND_RIGHT_TOL_PT = 0.5;
/** Summary width, px — contract F14.B (design 597.16). */
const SUMMARY_WIDTH_PX = 597.2;
const SUMMARY_WIDTH_TOL_PX = 2;
/** Summary measure ceiling, mm — plan F14.C ("no line exceeds 158mm"). */
const SUMMARY_MAX_LINE_MM = 158;

/** The rounded regions — contract F15 (six; the mark tiles carry their own
 * radius in the SVG assets). */
const ROUNDED_REGIONS = ["status-pill", "charges-head", "totals-panel", "total-bar", "reference-panel", "scan-card"];

/** The baked grand total that must NOT appear in publish output — contract I4
 * fixture figure, plan D13 ("rather than a baked figure"). */
const BAKED_TOTAL = "22,380.00";
/** Tokens the publish artefact must keep — plan D13/D14. */
// `useDeriver` names a result after the deriver that produced it, so the
// token reads as what made it. The objective is that the figure travels as a
// token at all — the name is how, not what.
const DERIVED_TOTAL_TOKEN = "{{derived.invoiceTotals.total}}";
const DERIVED_QR_TOKEN = "{{derived.paymentQr}}";
const QR_DERIVER_NAME = "paymentQr";

// ---------------------------------------------------------------------------
// The region table (contract) flattened out of the fixture, with pages.
// ---------------------------------------------------------------------------

/** {name: {page, x, y, w, h}} for every named design region. */
export const DESIGN_REGIONS = (() => {
  const out = {};
  for (const [page, key] of [[1, "page1"], [2, "page2"]]) {
    for (const [name, rect] of Object.entries(DESIGN[key])) {
      if (name === "charge-row-heights") continue;
      out[name] = { page, ...rect };
    }
  }
  return out;
})();

/** Contract region-table rows, in order. qr-canvas is fixture-only (F13.B). */
const REGION_NAMES = [
  "letterhead", "rule", "band", "status-pill", "parties", "summary-label",
  "summary", "charges-head", "charges-body", "charges-rule", "totals-panel",
  "total-bar", "closer", "footer",
  "p2-letterhead", "p2-rule", "bank-grid", "reference-panel", "terms",
  "scan-card", "footer-2",
];
/** Located geometrically in B, so probe C cannot Find them — contract. */
const GEOMETRIC_REGIONS = new Set(["rule", "p2-rule", "charges-rule"]);
/** Regions with visual: yes in the contract region table (G4's universe). */
const VISUAL_REGIONS = REGION_NAMES.filter(
  (n) => !GEOMETRIC_REGIONS.has(n) && !G4_EXEMPT.has(n),
);

// ---------------------------------------------------------------------------
// Helpers — conversions, result constructors, measurement accessors.
// ---------------------------------------------------------------------------

/** px at 96dpi -> Word points. */
export const pxToPt = (px) => px * PX_TO_PT;
/** mm -> Word points. */
export const mmToPt = (mm) => mm * PT_PER_MM;
/** mm -> px at 96dpi. */
export const mmToPx = (mm) => mm * PX_PER_MM;

const round2 = (n) => Math.round(n * 100) / 100;
const inRange = (n, [lo, hi]) => typeof n === "number" && n >= lo && n <= hi;
const near = (n, target, tol) => typeof n === "number" && Math.abs(n - target) <= tol;
const clip = (s, len = 140) => {
  const t = String(s ?? "");
  return t.length > len ? `${t.slice(0, len)}…` : t;
};

const pass = (measured, expected, note) => ({ status: "PASS", measured, expected, ...(note ? { note } : {}) });
const fail = (measured, expected, note) => ({ status: "FAIL", measured, expected, ...(note ? { note } : {}) });
const blocked = (measured, expected, note) => ({ status: "BLOCKED", measured, expected, ...(note ? { note } : {}) });
/** The uniform verdict for a probe that never wrote its file. */
const probeMissing = (probe, expected) => fail(`measure-${probe} absent`, expected, "probe missing");

/**
 * Is a preview rect within tolerance of a design rect on all four numbers?
 * Returns the worst axis so a note can say how far out it was.
 */
export function regionsClose(bRect, designRect, tolPx) {
  if (!bRect || !designRect) return { ok: false, worst: Infinity, detail: "rect missing" };
  const deltas = {
    dx: Math.abs((bRect.x ?? NaN) - designRect.x),
    dy: Math.abs((bRect.y ?? NaN) - designRect.y),
    dw: Math.abs((bRect.w ?? NaN) - designRect.w),
    dh: Math.abs((bRect.h ?? NaN) - designRect.h),
  };
  const worst = Math.max(...Object.values(deltas));
  const ok = Number.isFinite(worst) && worst <= tolPx;
  const detail = Object.entries(deltas)
    .map(([k, v]) => `${k}=${Number.isFinite(v) ? round2(v) : "?"}`)
    .join(" ");
  return { ok, worst, detail };
}

/** The body table containing the given run text, or undefined. */
function tableWithText(aLike, text) {
  return (aLike?.bodyTables ?? []).find((t) =>
    (t.rows ?? []).some((r) =>
      (r.cells ?? []).some((c) =>
        (c.runs ?? []).some((run) => typeof run.text === "string" && run.text.includes(text)))));
}

const bodyRows = (table) => (table?.rows ?? []).filter((r) => !r.header);
const headerRows = (table) => (table?.rows ?? []).filter((r) => r.header);

/** All run texts of a table row, joined — for finding anchors in rows. */
const rowText = (row) =>
  (row?.cells ?? []).map((c) => (c.runs ?? []).map((r) => r.text ?? "").join("")).join(" ");

/** The header/footer part a sectPr reference of the given type points at. */
function refPart(aLike, kind, type) {
  const refs = aLike?.sectPr?.[`${kind}Refs`] ?? [];
  const ref = refs.find((r) => r.type === type);
  return ref ? aLike?.[`${kind}Parts`]?.[ref.part] : undefined;
}

/** The default footer part, falling back to any footer part with a table. */
function defaultFooterPart(aLike) {
  return refPart(aLike, "footer", "default")
    ?? Object.values(aLike?.footerParts ?? {}).find((p) => (p.tables ?? []).length > 0)
    ?? Object.values(aLike?.footerParts ?? {})[0];
}

/**
 * F8's alternation, judged on any measure-a-like object so the same rule
 * decides the default, publish and length-variant builds.
 */
function zebraVerdict(aLike) {
  const table = tableWithText(aLike, "Description");
  if (!table) return { ok: false, measured: "charges table not found" };
  const heads = headerRows(table);
  const body = bodyRows(table);
  if (heads.length !== 1) return { ok: false, measured: `${heads.length} header rows` };
  const headOk = (heads[0].cells ?? []).every((c) => c.shd === ACCENT_FILL);
  const wrong = [];
  body.forEach((row, i) => {
    const n = i + 1; // 1-based body row number
    const fills = (row.cells ?? []).map((c) => c.shd ?? null);
    const want = n % 2 === 0 ? ZEBRA_FILL : null;
    if (!fills.every((f) => f === want)) wrong.push(`row${n}=${fills[0] ?? "none"}`);
  });
  const accentBodyRows = body.filter((row) => (row.cells ?? []).some((c) => c.shd === ACCENT_FILL)).length;
  const ok = headOk && wrong.length === 0 && accentBodyRows === 0;
  const measured = `header accent=${headOk}, ${body.length} body rows, off-pattern: ${wrong.length ? wrong.join(", ") : "none"}, accent body rows=${accentBodyRows}`;
  return { ok, measured };
}

/**
 * The shared BLOCKED verdict for objectives behind the publish build; null
 * when the publish build ran and succeeded. Contract: a throwing publish
 * build records BLOCKED with the thrown message as evidence.
 */
function publishBlocked(build, expected) {
  const p = build?.publish;
  if (!p) return blocked("publish build not run", expected, "requires --publish build");
  if (!p.ok) return blocked(`publish threw: ${clip(p.error)}`, expected, "blocked by publish-build");
  return null;
}

/** document.json's text, when the runner passed it through; else null. */
function publishDocumentText(build) {
  const doc = build?.publish?.documentJson;
  if (typeof doc === "string") return doc;
  if (doc && typeof doc === "object") return JSON.stringify(doc);
  return null;
}

/** One corner-test entry judged: all corners ground, centre filled. */
function cornerEntryOk(entry) {
  if (!entry) return false;
  const corners = Array.isArray(entry.corners)
    ? entry.corners.length === 4 && entry.corners.every(Boolean)
    : entry.cornersOk === true;
  return corners && entry.centerOk === true;
}

/** F15's corner test over one engine's entries (preview or word raster). */
function cornerVerdict(entries, engine, expected) {
  if (!Array.isArray(entries)) return fail(`${engine} corner tests absent`, expected, "not measured");
  const pill = entries.find((e) => e.region === "status-pill");
  if (!pill) return fail("status-pill entry absent", expected, "not measured");
  // A region with no fill of its own cannot be judged corner-against-centre.
  // Skipping it is honest; counting it as square would be a finding the
  // measurement did not make.
  const judged = entries.filter((e) => e.applicable !== false);
  if (judged.length === 0) {
    return fail(`no ${engine} region carried a fill to judge`, expected, "not measured");
  }
  const bad = judged.filter((e) => !cornerEntryOk(e)).map((e) => e.region);
  const skipped = entries.length - judged.length;
  const note = skipped > 0 ? `${skipped} region(s) unfilled, not judged` : "";
  return bad.length === 0
    ? pass(`${judged.length} region(s) rounded (${engine})`, expected, note)
    : fail(`square in ${engine}: ${bad.join(", ")}`, expected, note);
}

/**
 * P0 judged from the measurements themselves, so P0-blocked objectives can
 * decide BLOCKED without the runner's help. Chrome: the font probe's Aptos
 * width must differ from the mono fallback (identical widths mean the
 * fallback face was used). Word: probe C having run at all proves Word opened
 * the file with its per-user Aptos install; a fontNames list or the runner's
 * registry check tightens that when present.
 */
function p0Verdict(m) {
  const expected = `chrome |aptos - mono| > ${FONT_PROBE_MIN_DELTA_PX}px; Word resolves Aptos`;
  const fp = m?.b?.fontProbe;
  if (!fp) return probeMissing("b", expected);
  const delta = Math.abs((fp.aptosWidth ?? NaN) - (fp.monoWidth ?? NaN));
  const chromeOk = Number.isFinite(delta) && delta > FONT_PROBE_MIN_DELTA_PX;
  let wordOk = m?.c != null;
  let wordWhy = wordOk ? "probe C ran" : "measure-c absent";
  if (Array.isArray(m?.c?.fontNames)) {
    wordOk = m.c.fontNames.some((n) => /^Aptos$/i.test(String(n)));
    wordWhy = wordOk ? "FontNames lists Aptos" : "FontNames misses Aptos";
  }
  if (m?.build?.aptosRegistered === false) {
    wordOk = false;
    wordWhy = "runner registry check failed";
  }
  const measured = `chromeAptos=${fp.aptosWidth} mono=${fp.monoWidth} (Δ${round2(delta)}); word: ${wordWhy}`;
  return chromeOk && wordOk ? pass(measured, expected) : fail(measured, expected);
}

/** Wraps a P0-gated evaluate: BLOCKED until both engines resolve the face. */
const gatedByP0 = (evaluate) => (m) => {
  const p0 = p0Verdict(m);
  if (p0.status !== "PASS") {
    return blocked(`P0 unmet (${p0.measured})`, p0.expected, "blocked by P0");
  }
  return evaluate(m);
};

// ---------------------------------------------------------------------------
// The objectives. One entry per probe line of plan §3/§4, with the contract's
// corrected numbers. Order follows the plan.
// ---------------------------------------------------------------------------

/** @type {Array<{id: string, gap: string, probe: string, blockedBy?: string, description: string, evaluate: (m: object) => {status: string, measured: any, expected: any, note?: string}}>} */
export const OBJECTIVES = [
  // --- F1 — a section that doesn't print its heading -----------------------
  {
    id: "F1.A1",
    gap: "F1",
    probe: "A",
    description: "No Heading1 for Invoice details/Parties/Charges; one for each of the four kept",
    evaluate(m) {
      const expected = `none of [${SUPPRESSED_HEADINGS.join(", ")}]; all of [${KEPT_HEADINGS.join(", ")}]`;
      if (!m.a) return probeMissing("a", expected);
      const heads = m.a.headings1;
      if (!Array.isArray(heads)) return fail("headings1 absent", expected, "not found");
      const unwanted = SUPPRESSED_HEADINGS.filter((h) => heads.includes(h));
      const lost = KEPT_HEADINGS.filter((h) => !heads.includes(h));
      const measured = `unwanted: [${unwanted.join(", ")}], missing kept: [${lost.join(", ")}]`;
      return unwanted.length === 0 && lost.length === 0 ? pass(measured, expected) : fail(measured, expected);
    },
  },
  {
    id: "F1.A2",
    gap: "F1",
    probe: "A",
    description: "The built model still carries every section with its title (suppressing the heading must not cost the name)",
    evaluate(m) {
      const expected = `sectionsInModel = ${SECTIONS_IN_MODEL}`;
      if (!m.a) return probeMissing("a", expected);
      const n = m.a.sectionsInModel;
      if (typeof n !== "number") return fail("sectionsInModel absent", expected, "not found");
      return n === SECTIONS_IN_MODEL ? pass(n, expected) : fail(n, expected);
    },
  },
  {
    id: "F1.B1",
    gap: "F1",
    probe: "B",
    description: "No text between letterhead and band: band starts <= 3px below the rule",
    evaluate(m) {
      const expected = `band.y - (rule.y + rule.h) in [${BAND_GAP_MIN_PX}, ${BAND_GAP_MAX_PX}]px`;
      if (!m.b) return probeMissing("b", expected);
      const rule = m.b.regions?.rule;
      const band = m.b.regions?.band;
      if (!rule || !band) return fail(`rule ${rule ? "found" : "missing"}, band ${band ? "found" : "missing"}`, expected, "not found");
      const gap = band.y - (rule.y + rule.h);
      return gap >= BAND_GAP_MIN_PX && gap <= BAND_GAP_MAX_PX
        ? pass(`gap ${round2(gap)}px`, expected)
        : fail(`gap ${round2(gap)}px`, expected);
    },
  },
  {
    id: "F1.C1",
    gap: "F1",
    probe: "C",
    description: "Word's Find matches none of the three suppressed heading strings",
    evaluate(m) {
      const expected = "Find misses all three suppressed headings";
      if (!m.c) return probeMissing("c", expected);
      const finds = m.c.suppressedHeadingFinds;
      if (!finds || typeof finds !== "object") return fail("suppressedHeadingFinds absent", expected, "not found");
      const found = SUPPRESSED_HEADINGS.filter((h) => finds[h] === true);
      const unmeasured = SUPPRESSED_HEADINGS.filter((h) => typeof finds[h] !== "boolean");
      if (unmeasured.length) return fail(`unmeasured: [${unmeasured.join(", ")}]`, expected, "not measured");
      return found.length === 0
        ? pass("no matches", expected)
        : fail(`found: [${found.join(", ")}]`, expected);
    },
  },

  // --- F2 — different furniture on the first page --------------------------
  {
    id: "F2.A1",
    gap: "F2",
    probe: "A",
    description: "sectPr has titlePg; the first-page header part is empty; the default header carries the reference",
    evaluate(m) {
      const expected = `titlePg; first header empty; default header contains "${INVOICE_REFERENCE}"`;
      if (!m.a) return probeMissing("a", expected);
      const titlePg = m.a.sectPr?.titlePg === true;
      const first = refPart(m.a, "header", "first");
      const dflt = refPart(m.a, "header", "default");
      const firstEmpty = first != null && !(first.text ?? "").trim();
      const dfltOk = typeof dflt?.text === "string" && dflt.text.includes(INVOICE_REFERENCE);
      const measured = `titlePg=${titlePg}, first=${first ? JSON.stringify(clip(first.text, 40)) : "absent"}, default has ref=${dfltOk}`;
      return titlePg && firstEmpty && dfltOk ? pass(measured, expected) : fail(measured, expected);
    },
  },
  {
    id: "F2.B1",
    gap: "F2",
    probe: "B",
    description: "Page 1's rendered header is <= 2px tall; page 2's is 40-80px",
    evaluate(m) {
      const expected = `header h: page1 <= ${FIRST_HEADER_MAX_PX}px, page2 in [${PRIMARY_HEADER_PX}]px`;
      if (!m.b) return probeMissing("b", expected);
      const s = m.b.sections;
      const h1 = s?.[0]?.headerHeight;
      const h2 = s?.[1]?.headerHeight;
      if (typeof h1 !== "number" || typeof h2 !== "number") {
        return fail(`headerHeight page1=${h1 ?? "?"} page2=${h2 ?? "?"}`, expected, "not found");
      }
      const ok = h1 <= FIRST_HEADER_MAX_PX && inRange(h2, PRIMARY_HEADER_PX);
      return (ok ? pass : fail)(`page1 ${round2(h1)}px, page2 ${round2(h2)}px`, expected);
    },
  },
  {
    id: "F2.C1",
    gap: "F2",
    probe: "C",
    description: "Word's first-page header is empty; the primary header carries the reference",
    evaluate(m) {
      const expected = `Headers(2) empty; Headers(1) contains "${INVOICE_REFERENCE}"`;
      if (!m.c) return probeMissing("c", expected);
      const first = m.c.firstPageHeaderText;
      const primary = m.c.primaryHeaderText;
      if (typeof first !== "string" || typeof primary !== "string") {
        return fail("header texts absent", expected, "not found");
      }
      const firstEmpty = !first.replace(/[\r\n]/g, "").trim();
      const primaryOk = primary.includes(INVOICE_REFERENCE);
      const measured = `first=${JSON.stringify(clip(first, 30))}, primary has ref=${primaryOk}`;
      return firstEmpty && primaryOk ? pass(measured, expected) : fail(measured, expected);
    },
  },

  // --- F3 — a table that bleeds --------------------------------------------
  {
    id: "F3.A1",
    gap: "F3",
    probe: "A",
    description: "The footer table carries tblInd -907 and its columns total 11906 twips",
    evaluate(m) {
      const expected = `tblInd ${FOOTER_TBLIND_TWIPS}, widths sum ${PAGE_WIDTH_TWIPS}`;
      if (!m.a) return probeMissing("a", expected);
      const table = defaultFooterPart(m.a)?.tables?.[0];
      if (!table) return fail("footer table absent", expected, "not found");
      const sum = Array.isArray(table.widthsTwips)
        ? table.widthsTwips.reduce((acc, w) => acc + (w ?? 0), 0)
        : NaN;
      const measured = `tblInd=${table.tblInd ?? "none"}, widths sum ${sum}`;
      return table.tblInd === FOOTER_TBLIND_TWIPS && sum === PAGE_WIDTH_TWIPS
        ? pass(measured, expected)
        : fail(measured, expected);
    },
  },
  {
    id: "F3.B1",
    gap: "F3",
    probe: "B",
    description: "The footer's left edge is at x=0 and its width equals the section width, +/-1px",
    evaluate(m) {
      const expected = `x = 0 ± ${FOOTER_EDGE_TOL_PX}px, w = section w ± ${FOOTER_EDGE_TOL_PX}px`;
      if (!m.b) return probeMissing("b", expected);
      const r = m.b.regions?.footer;
      const section = m.b.sections?.[0];
      const sw = section?.w;
      if (!r || typeof sw !== "number") return fail(`footer ${r ? "found" : "missing"}, section w ${sw ?? "?"}`, expected, "not found");
      // Region x is content-box relative like every other region, and "the
      // paper's left edge" is not the text column's — comparing the two
      // straight would call a bar that starts at the margin edge-to-edge.
      // The offset between them is the page's left padding; the content box's
      // own x is measured from the viewport, which the paper is centred in.
      const pageX = r.x + (section.padding?.left ?? 0);
      const ok = near(pageX, 0, FOOTER_EDGE_TOL_PX) && near(r.w, sw, FOOTER_EDGE_TOL_PX);
      return (ok ? pass : fail)(
        `page x=${round2(pageX)} w=${round2(r.w)} (section ${round2(sw)})`,
        expected,
        ok ? "" : "docx-preview reads w:tblInd's indent from a w:left it never carries, so a bleeding table stays at the margin there; Word honours it — see F3.C1",
      );
    },
  },
  {
    id: "F3.C1",
    gap: "F3",
    probe: "C",
    description: "Word puts the footer table at 0pt from the page edge, 595.3pt wide",
    evaluate(m) {
      const expected = `x ${FOOTER_X_PT} ± ${FOOTER_TABLE_TOL_PT}pt, width ${FOOTER_WIDTH_PT} ± ${FOOTER_TABLE_TOL_PT}pt`;
      if (!m.c) return probeMissing("c", expected);
      const t = m.c.footerTable;
      if (!t) return fail("footerTable absent", expected, "not found");
      // Where the table starts on the paper: the text column's left margin
      // plus the table's own indent out of it. Information(5) is not that —
      // it reports where the first cell's text starts, inside the cell.
      const leftMargin = m.c.pageSetup?.leftMargin;
      const startPt = typeof leftMargin === "number" && typeof t.leftIndent === "number"
        ? leftMargin + t.leftIndent
        : t.x;
      // Across the cells, not through PreferredWidth. Word reports the
      // latter as wdUndefined for a multi-column table laid out fixed — the
      // widths belong to the columns then, not to a preference the table
      // expresses — so asking the table would measure the reading rather than
      // the document. The old value stays as the fallback for a table Word
      // will still answer for.
      const widthPt = typeof t.width === "number" ? t.width : t.preferredWidth;
      const ok = near(startPt, FOOTER_X_PT, FOOTER_TABLE_TOL_PT)
        && near(widthPt, FOOTER_WIDTH_PT, FOOTER_TABLE_TOL_PT);
      return (ok ? pass : fail)(
        `starts ${round2(startPt)}pt w=${widthPt}pt (text at ${t.x}pt)`,
        expected,
      );
    },
  },

  // --- F4 — an image inline with text --------------------------------------
  {
    id: "F4.A1",
    gap: "F4",
    probe: "A",
    description: "Exactly one footer paragraph holds both a drawing and 'Generated with Docxcelerate'",
    evaluate(m) {
      const expected = `1 w:p with drawing + "${FOOTER_MARK_TEXT}"`;
      if (!m.a) return probeMissing("a", expected);
      const fp = defaultFooterPart(m.a);
      if (!fp) return fail("footer part absent", expected, "not found");
      const count = fp.drawingWithTextParas;
      if (typeof count !== "number") return fail("drawingWithTextParas absent", expected, "not measured");
      const textOk = typeof fp.drawingWithTextSample !== "string"
        || fp.drawingWithTextSample.includes(FOOTER_MARK_TEXT);
      const measured = `${count} paragraph(s), sample=${JSON.stringify(clip(fp.drawingWithTextSample, 40))}`;
      return count === 1 && textOk ? pass(measured, expected) : fail(measured, expected);
    },
  },
  {
    id: "F4.B1",
    gap: "F4",
    probe: "B",
    description: "The rendered footer element is 52-58px tall (design 54.78)",
    evaluate(m) {
      const expected = `footer height in [${FOOTER_HEIGHT_PX}]px`;
      if (!m.b) return probeMissing("b", expected);
      const h = m.b.sections?.[0]?.footerHeight ?? m.b.regions?.footer?.h;
      if (typeof h !== "number") return fail("footer height absent", expected, "not found");
      return (inRange(h, FOOTER_HEIGHT_PX) ? pass : fail)(`${round2(h)}px`, expected);
    },
  },
  {
    id: "F4.C1",
    gap: "F4",
    probe: "C",
    description: "Word's footer is <= 14.8mm tall and one line",
    evaluate(m) {
      const maxPt = round2(mmToPt(FOOTER_MAX_MM));
      const expected = `height <= ${maxPt}pt (${FOOTER_MAX_MM}mm), ${FOOTER_LINES} line`;
      if (!m.c) return probeMissing("c", expected);
      const f = m.c.footer;
      if (!f || typeof f.heightPt !== "number" || typeof f.lineCount !== "number") {
        return fail("footer height/lineCount absent", expected, "not measured");
      }
      const ok = f.heightPt <= maxPt && f.lineCount === FOOTER_LINES;
      return (ok ? pass : fail)(`${round2(f.heightPt)}pt, ${f.lineCount} line(s)`, expected);
    },
  },

  // --- F5 — non-paragraph cell children obey the cell ----------------------
  {
    id: "F5.A1",
    gap: "F5",
    probe: "A",
    description: "The PAGE-field paragraph is right-aligned with zero space after",
    evaluate(m) {
      const expected = "jc=right, spacingAfter=0, PAGE + NUMPAGES fields present";
      if (!m.a) return probeMissing("a", expected);
      const p = defaultFooterPart(m.a)?.pageFieldPara;
      if (!p) return fail("pageFieldPara absent", expected, "not found");
      const ok = p.jc === "right" && p.spacingAfter === 0 && p.hasPageField === true && p.hasNumPagesField === true;
      const measured = `jc=${p.jc ?? "none"}, after=${p.spacingAfter ?? "none"}, page=${p.hasPageField}, numpages=${p.hasNumPagesField}`;
      return (ok ? pass : fail)(measured, expected);
    },
  },
  {
    id: "F5.B1",
    gap: "F5",
    probe: "B",
    description: "The page number's right edge sits at the design footer inset, x=733.2 +/- 2px",
    evaluate(m) {
      const expected = `right edge ${PAGE_NUM_RIGHT_PX} ± ${PAGE_NUM_RIGHT_TOL_PX}px`;
      if (!m.b) return probeMissing("b", expected);
      const right = m.b.sections?.[0]?.pageNumberRightPx;
      if (typeof right !== "number") return fail("pageNumberRightPx absent", expected, "not measured");
      return (near(right, PAGE_NUM_RIGHT_PX, PAGE_NUM_RIGHT_TOL_PX) ? pass : fail)(`${round2(right)}px`, expected);
    },
  },
  {
    id: "F5.C1",
    gap: "F5",
    probe: "C",
    description: "Word reads the page-number paragraph as right-aligned with SpaceAfter 0",
    evaluate(m) {
      const expected = `Alignment=${WD_ALIGN_RIGHT}, SpaceAfter=0`;
      if (!m.c) return probeMissing("c", expected);
      const p = m.c.pageNumberParagraph;
      if (!p) return fail("pageNumberParagraph absent", expected, "not measured");
      const ok = p.alignment === WD_ALIGN_RIGHT && p.spaceAfter === 0;
      return (ok ? pass : fail)(`alignment=${p.alignment}, spaceAfter=${p.spaceAfter}`, expected);
    },
  },
  {
    id: "F5.C2",
    gap: "F5",
    probe: "V",
    description: "On the rendered pages the footers read '1 / 2' and '2 / 2'",
    evaluate(m) {
      const expected = PAGE_FIELD_TEXTS.join(" then ");
      // Word evaluates a PAGE field once per story, not once per page, so the
      // per-page text exists nowhere but the rendered PDF. Probe V reads it
      // out of the same export Word produced, which keeps it Word's answer.
      if (!m.v) return probeMissing("v", expected);
      const texts = m.v.footerTextByPage;
      if (!Array.isArray(texts)) return fail("footerTextByPage absent", expected, "not measured");
      const ok = PAGE_FIELD_TEXTS.every((t, i) => typeof texts[i] === "string" && texts[i].includes(t));
      return (ok ? pass : fail)(texts.map((t) => JSON.stringify(clip(t, 20))).join(", "), expected);
    },
  },

  // --- F6 — a block can name a face ----------------------------------------
  {
    id: "F6.A1",
    gap: "F6",
    probe: "A",
    description: "Qty/rate/amount runs name Consolas; description runs do not",
    evaluate(m) {
      const expected = `rFonts ${MONEY_FONT} on cols 2-4, not on col 1`;
      if (!m.a) return probeMissing("a", expected);
      const table = tableWithText(m.a, "Description");
      if (!table) return fail("charges table absent", expected, "not found");
      let moneyRuns = 0;
      let moneyWrong = 0;
      let descWrong = 0;
      for (const row of bodyRows(table)) {
        (row.cells ?? []).forEach((cell, col) => {
          for (const run of cell.runs ?? []) {
            if (!(run.text ?? "").trim()) continue;
            if (col === 0) {
              if (run.font === MONEY_FONT) descWrong += 1;
            } else {
              moneyRuns += 1;
              if (run.font !== MONEY_FONT) moneyWrong += 1;
            }
          }
        });
      }
      const ok = moneyRuns > 0 && moneyWrong === 0 && descWrong === 0;
      return (ok ? pass : fail)(`${moneyRuns} money runs, ${moneyWrong} wrong; ${descWrong} desc runs in ${MONEY_FONT}`, expected);
    },
  },
  {
    id: "F6.B1",
    gap: "F6",
    probe: "B",
    description: "Chrome resolves the amount cell to Consolas and the description cell to the body face",
    evaluate(m) {
      const expected = `amount cell ${MONEY_FONT}; description cell not ${MONEY_FONT}`;
      if (!m.b) return probeMissing("b", expected);
      const f = m.b.fonts;
      if (!f || typeof f.amountCell !== "string" || typeof f.descriptionCell !== "string") {
        return fail("fonts absent", expected, "not found");
      }
      const ok = f.amountCell.includes(MONEY_FONT) && !f.descriptionCell.includes(MONEY_FONT);
      return (ok ? pass : fail)(`amount=${clip(f.amountCell, 40)}, desc=${clip(f.descriptionCell, 40)}`, expected);
    },
  },
  {
    id: "F6.C1",
    gap: "F6",
    probe: "C",
    description: "Word reads the amount cell's font as Consolas",
    evaluate(m) {
      const expected = MONEY_FONT;
      if (!m.c) return probeMissing("c", expected);
      const name = m.c.chargesTable?.amountFontName;
      if (typeof name !== "string") return fail("amount cell font absent", expected, "not measured");
      return (name === MONEY_FONT ? pass : fail)(name, expected);
    },
  },
  {
    id: "F6.C2",
    gap: "F6",
    probe: "C",
    description: "Every amount's right edge agrees within 0.35mm (the objective the font serves)",
    evaluate(m) {
      const tolPt = round2(mmToPt(AMOUNT_EDGE_TOL_MM));
      const expected = `right-edge spread <= ${tolPt}pt (${AMOUNT_EDGE_TOL_MM}mm)`;
      if (!m.c) return probeMissing("c", expected);
      // The amount column alone. The totals panel is a different table set
      // further in, and qty and rate are columns of their own inside this one
      // — three columns of figures line up on three different edges by
      // design, so folding them together would measure the layout, not the
      // column the objective is about.
      const inCharges = (m.c.amounts ?? []).filter((a) =>
        typeof a.tableCell11 === "string"
        && a.tableCell11.toUpperCase().startsWith("DESCRIPTION")
        && typeof a.xEnd === "number");
      const edges = inCharges
        .filter((a) => a.columnIndex === a.columnCount)
        .map((a) => a.xEnd);
      if (edges.length < 2) {
        return fail(`${edges.length} charge amount(s) located`, expected, "not measured");
      }
      const spread = Math.max(...edges) - Math.min(...edges);
      return (spread <= tolPt ? pass : fail)(`spread ${round2(spread)}pt over ${edges.length} amounts`, expected);
    },
  },

  // --- F7 — a block can set its own leading --------------------------------
  {
    id: "F7.A1",
    gap: "F7",
    probe: "A",
    description: "Charge rows carry a line value under the body's 348; the muted note tighter than its sibling",
    evaluate(m) {
      const expected = `every cell line < ${BODY_LINE_TWIPS}; note line < desc line`;
      if (!m.a) return probeMissing("a", expected);
      const table = tableWithText(m.a, "Description");
      if (!table) return fail("charges table absent", expected, "not found");
      const rows = bodyRows(table);
      if (rows.length === 0) return fail("no body rows", expected, "not found");
      let loose = 0;
      let noteChecked = 0;
      let noteWrong = 0;
      for (const row of rows) {
        for (const cell of row.cells ?? []) {
          const line = cell.paraSpacing?.line;
          if (typeof line !== "number" || line >= BODY_LINE_TWIPS) loose += 1;
          if (cell.noteSpacing) {
            noteChecked += 1;
            if (!(typeof cell.noteSpacing.line === "number"
              && typeof line === "number"
              && cell.noteSpacing.line < line)) noteWrong += 1;
          }
        }
      }
      if (noteChecked === 0) return fail(`${loose} loose cell(s); note spacing unmeasured`, expected, "muted-note spacing not measured");
      const ok = loose === 0 && noteWrong === 0;
      return (ok ? pass : fail)(`${loose} loose cell(s); ${noteWrong}/${noteChecked} notes not tighter`, expected);
    },
  },
  {
    id: "F7.B1",
    gap: "F7",
    probe: "B",
    description: "Every charge row is 50-58px tall and charges-body is 370-388px (design 54.11 / 378.77)",
    evaluate(m) {
      const expected = `rows in [${CHARGE_ROW_PX}]px, charges-body h in [${CHARGES_BODY_PX}]px`;
      if (!m.b) return probeMissing("b", expected);
      const rows = m.b.chargeRowHeights;
      const body = m.b.regions?.["charges-body"];
      if (!Array.isArray(rows) || rows.length === 0 || !body) {
        return fail(`rows ${Array.isArray(rows) ? rows.length : "absent"}, charges-body ${body ? "found" : "missing"}`, expected, "not found");
      }
      const badRows = rows.filter((h) => !inRange(h, CHARGE_ROW_PX));
      const ok = badRows.length === 0 && inRange(body.h, CHARGES_BODY_PX);
      return (ok ? pass : fail)(`rows [${rows.map(round2).join(", ")}], body ${round2(body.h)}px`, expected);
    },
  },
  {
    id: "F7.C1",
    gap: "F7",
    probe: "C",
    description: "Charges head + body span <= 111mm in Word",
    evaluate(m) {
      const maxPt = round2(mmToPt(CHARGES_TOTAL_MAX_MM));
      const expected = `bottom - top <= ${maxPt}pt (${CHARGES_TOTAL_MAX_MM}mm)`;
      if (!m.c) return probeMissing("c", expected);
      const head = m.c.regions?.["charges-head"];
      const end = m.c.regions?.["charges-body-end"];
      if (!head || !end) return fail("charges head/end absent", expected, "not found");
      // A table that has spilled onto the next page has no single span to
      // measure — and that is the largest failure of one, not an absence of
      // measurement, so it is reported as the failure it is.
      if (head.page !== end.page) {
        return fail(`spills p${head.page} -> p${end.page}`, expected, "charges cross a page boundary");
      }
      // Top of the head row to the top of the last body row: short by that
      // row's own height, so a pass here has margin in hand rather than debt.
      const span = end.y - head.y;
      return (span <= maxPt ? pass : fail)(`${round2(span)}pt (${round2(span / PT_PER_MM)}mm)`, expected);
    },
  },
  {
    id: "F7.C2",
    gap: "F7",
    probe: "C",
    description: "No charge row wraps to a third line at the fixture's data lengths",
    evaluate(m) {
      const expected = `every row <= ${CHARGE_ROW_MAX_LINES} lines`;
      if (!m.c) return probeMissing("c", expected);
      const counts = m.c.chargesTable?.rowLineCounts;
      if (!Array.isArray(counts) || counts.length === 0) return fail("chargeRowLineCounts absent", expected, "not measured");
      const over = counts.filter((n) => n > CHARGE_ROW_MAX_LINES);
      return (over.length === 0 ? pass : fail)(`[${counts.join(", ")}]`, expected);
    },
  },

  // --- F8 — alternating row fill, from the theme ---------------------------
  {
    id: "F8.A1",
    gap: "F8",
    probe: "A",
    description: "Body rows 2/4/6 carry F7F8FD, odd rows carry no shd, header keeps the accent",
    evaluate(m) {
      const expected = `even body rows ${ZEBRA_FILL}, odd rows unshaded, header ${ACCENT_FILL}`;
      if (!m.a) return probeMissing("a", expected);
      const v = zebraVerdict(m.a);
      return (v.ok ? pass : fail)(v.measured, expected);
    },
  },
  {
    id: "F8.A2",
    gap: "F8",
    probe: "A",
    blockedBy: "publish-build",
    description: "Striping still alternates when the template is built through the publish path",
    evaluate(m) {
      const expected = `publish-built docx alternates (${ZEBRA_FILL} on even body rows)`;
      const gate = publishBlocked(m.build, expected);
      if (gate) return gate;
      const variant = m.a?.variants?.publish;
      if (!variant) return fail("publish variant not measured by probe A", expected, "not measured");
      const v = zebraVerdict(variant);
      return (v.ok ? pass : fail)(v.measured, expected);
    },
  },
  {
    id: "F8.A3",
    gap: "F8",
    probe: "A",
    description: "Striping alternates at 3 lines and at 30, with the header still the only accent row",
    evaluate(m) {
      const expected = "zebra pattern holds in --lines=3 and --lines=30 builds";
      if (!m.a) return probeMissing("a", expected);
      const v3 = m.a.variants?.lines3;
      const v30 = m.a.variants?.lines30;
      if (!v3 || !v30) return fail(`lines3 ${v3 ? "measured" : "missing"}, lines30 ${v30 ? "measured" : "missing"}`, expected, "not measured");
      const r3 = zebraVerdict(v3);
      const r30 = zebraVerdict(v30);
      const measured = `lines3: ${r3.measured}; lines30: ${r30.measured}`;
      return r3.ok && r30.ok ? pass(measured, expected) : fail(measured, expected);
    },
  },
  {
    id: "F8.C1",
    gap: "F8",
    probe: "C",
    description: "Word reads the zebra rows as RGB(247,248,253) and the plain rows as automatic",
    evaluate(m) {
      const expected = `col1 rows 1..8: [${ACCENT_RGB}, auto, ${ZEBRA_RGB}, auto, ${ZEBRA_RGB}, auto, ${ZEBRA_RGB}, auto]`;
      if (!m.c) return probeMissing("c", expected);
      const shading = m.c.chargesTable?.shadingCol1Rows1to8;
      if (!Array.isArray(shading) || shading.length < 8) return fail(`chargesCol1Shading ${Array.isArray(shading) ? shading.length : "absent"}`, expected, "not measured");
      const want = [ACCENT_RGB, WD_COLOR_AUTOMATIC, ZEBRA_RGB, WD_COLOR_AUTOMATIC, ZEBRA_RGB, WD_COLOR_AUTOMATIC, ZEBRA_RGB, WD_COLOR_AUTOMATIC];
      const wrong = want.map((w, i) => (shading[i] === w ? null : `row${i + 1}=${shading[i]}`)).filter(Boolean);
      return (wrong.length === 0 ? pass : fail)(wrong.length ? wrong.join(", ") : "pattern exact", expected);
    },
  },

  // --- F9 — turning the row hairline off -----------------------------------
  {
    id: "F9.A1",
    gap: "F9",
    probe: "A",
    description: "No D9DDEB border inside the charges table except the closing rule under the last row",
    evaluate(m) {
      const expected = `${RULE_COLOR} only as the last body row's bottom border`;
      if (!m.a) return probeMissing("a", expected);
      const table = tableWithText(m.a, "Description");
      if (!table) return fail("charges table absent", expected, "not found");
      const rows = table.rows ?? [];
      const lastIndex = rows.length - 1;
      const offenders = [];
      rows.forEach((row, ri) => {
        (row.cells ?? []).forEach((cell, ci) => {
          for (const [side, border] of Object.entries(cell.borders ?? {})) {
            if (border?.color !== RULE_COLOR) continue;
            if (ri === lastIndex && side === "bottom") continue;
            offenders.push(`r${ri + 1}c${ci + 1}.${side}`);
          }
        });
      });
      return (offenders.length === 0 ? pass : fail)(offenders.length ? `stray rules: ${offenders.join(", ")}` : "only the closing rule", expected);
    },
  },
  {
    id: "F9.A2",
    gap: "F9",
    probe: "A",
    description: "A block with borderSides: [] and no border emits no tcBorders at all (minimal fixture)",
    evaluate(m) {
      const expected = "borderFixture.emitsTcBorders = false";
      if (!m.a) return probeMissing("a", expected);
      const fx = m.a.borderFixture;
      if (!fx || typeof fx.emitsTcBorders !== "boolean") return fail("borderFixture absent", expected, "not measured");
      return (fx.emitsTcBorders === false ? pass : fail)(`emitsTcBorders=${fx.emitsTcBorders}`, expected);
    },
  },
  {
    id: "F9.C1",
    gap: "F9",
    probe: "C",
    description: "Word reads Cell(3,1)'s bottom border as wdLineStyleNone",
    evaluate(m) {
      const expected = `LineStyle ${WD_LINE_STYLE_NONE} (wdLineStyleNone)`;
      if (!m.c) return probeMissing("c", expected);
      const style = m.c.chargesTable?.row3BottomLineStyle;
      if (typeof style !== "number") return fail("chargesRow3BottomLineStyle absent", expected, "not measured");
      return (style === WD_LINE_STYLE_NONE ? pass : fail)(style, expected);
    },
  },

  // --- F10 — tracking on headings ------------------------------------------
  {
    id: "F10.A1",
    gap: "F10",
    probe: "A",
    description: "styles.xml's Heading1 run carries w:spacing 18 (0.12em at 7.5pt)",
    evaluate(m) {
      const expected = `heading1 spacing = ${HEADING1_TRACK_TWENTIETHS}`;
      if (!m.a) return probeMissing("a", expected);
      const h1 = m.a.stylesXml?.heading1;
      if (!h1) return fail("stylesXml.heading1 absent", expected, "not found");
      return (h1.spacing === HEADING1_TRACK_TWENTIETHS ? pass : fail)(`spacing=${h1.spacing ?? "none"}`, expected);
    },
  },
  {
    id: "F10.C1",
    gap: "F10",
    probe: "C",
    description: "Word reads Heading 1's Font.Spacing as 0.9pt",
    evaluate(m) {
      const expected = `${HEADING1_SPACING_PT} ± ${HEADING1_SPACING_TOL_PT}pt`;
      if (!m.c) return probeMissing("c", expected);
      const spacing = m.c.heading1Style?.fontSpacing;
      if (typeof spacing !== "number") return fail("heading1 spacing absent", expected, "not measured");
      return (near(spacing, HEADING1_SPACING_PT, HEADING1_SPACING_TOL_PT) ? pass : fail)(`${spacing}pt`, expected);
    },
  },
  {
    id: "F10.C2",
    gap: "F10",
    probe: "C",
    description: "ENGAGEMENT SUMMARY renders within 0.5mm of the design's width",
    evaluate(m) {
      const wantPt = round2(pxToPt(DESIGN_REGIONS["summary-label"].w));
      const tolPt = round2(mmToPt(LABEL_WIDTH_TOL_MM));
      const expected = `width ${wantPt} ± ${tolPt}pt (design ${round2(DESIGN_REGIONS["summary-label"].w)}px)`;
      if (!m.c) return probeMissing("c", expected);
      const r = m.c.regions?.["summary-label"];
      if (!r || typeof r.x !== "number" || typeof r.xEnd !== "number") {
        return fail("summary-label x/xEnd absent", expected, "not measured");
      }
      const width = r.xEnd - r.x;
      return (near(width, wantPt, tolPt) ? pass : fail)(`${round2(width)}pt`, expected);
    },
  },

  // --- F11 — vertical alignment in a cell ----------------------------------
  {
    id: "F11.A1",
    gap: "F11",
    probe: "A",
    description: "Band and footer cells carry vAlign center",
    evaluate(m) {
      const expected = "every band + footer cell vAlign=center";
      if (!m.a) return probeMissing("a", expected);
      const band = tableWithText(m.a, "Issue date");
      const footer = defaultFooterPart(m.a)?.tables?.[0];
      if (!band || !footer) return fail(`band ${band ? "found" : "missing"}, footer table ${footer ? "found" : "missing"}`, expected, "not found");
      const bandCells = (band.rows ?? []).flatMap((r) => r.cells ?? []);
      const footerCells = Array.isArray(footer.cells)
        ? footer.cells
        : (footer.rows ?? []).flatMap((r) => r.cells ?? []);
      const off = [...bandCells, ...footerCells].filter((c) => c.vAlign !== "center").length;
      const total = bandCells.length + footerCells.length;
      if (total === 0) return fail("no cells measured", expected, "not measured");
      return (off === 0 ? pass : fail)(`${off}/${total} cells not centred`, expected);
    },
  },
  {
    id: "F11.B1",
    gap: "F11",
    probe: "B",
    description: "The status pill's vertical centre matches the Issue date value's within 1px",
    evaluate(m) {
      const expected = `|Δcentre| <= ${PILL_CENTER_TOL_PX}px`;
      if (!m.b) return probeMissing("b", expected);
      const pill = m.b.regions?.["status-pill"];
      const anchor = m.b.regions?.["band-value"] ?? m.b.regions?.band;
      if (!pill || !anchor) return fail(`status-pill ${pill ? "found" : "missing"}, band ${anchor ? "found" : "missing"}`, expected, "not found");
      const delta = Math.abs((pill.y + pill.h / 2) - (anchor.y + anchor.h / 2));
      const which = m.b.regions?.["band-value"] ? "band-value" : "band";
      return (delta <= PILL_CENTER_TOL_PX ? pass : fail)(`Δ${round2(delta)}px vs ${which}`, expected);
    },
  },
  {
    id: "F11.C1",
    gap: "F11",
    probe: "C",
    description: "Word reads the band and footer cells as vertically centred",
    evaluate(m) {
      const expected = `every cell VerticalAlignment = ${WD_VALIGN_CENTER}`;
      if (!m.c) return probeMissing("c", expected);
      const v = m.c.vAlign;
      const band = v?.band;
      const footer = v?.footer;
      if (!Array.isArray(band) || !Array.isArray(footer) || band.length + footer.length === 0) {
        return fail("vAlign cells absent", expected, "not measured");
      }
      const off = [...band, ...footer].filter((x) => x !== WD_VALIGN_CENTER).length;
      return (off === 0 ? pass : fail)(`${off}/${band.length + footer.length} cells not centred`, expected);
    },
  },

  // --- F12 — a page break that doesn't leave a line ------------------------
  {
    id: "F12.A1",
    gap: "F12",
    probe: "A",
    description: "No lone-break paragraph; the payment page's first paragraph turns the page itself",
    evaluate(m) {
      // "Turns the page before itself" is the fact, not how it says so: the
      // break rides a style rather than the paragraph, because that is the one
      // form Word and docx-preview both read. Probe A counts either.
      const expected = "loneBreakParas=0, page-turning paragraphs=1";
      if (!m.a) return probeMissing("a", expected);
      const pb = m.a.pageBreaks;
      if (!pb) return fail("pageBreaks absent", expected, "not found");
      const before = Array.isArray(pb.pageBreakBefore) ? pb.pageBreakBefore.length : 0;
      const ok = pb.loneBreakParas === 0 && before === 1;
      return (ok ? pass : fail)(`lone=${pb.loneBreakParas}, pageBreakBefore=${before}`, expected);
    },
  },
  {
    id: "F12.A2",
    gap: "F12",
    probe: "A",
    description: "The break is Word's own break, never a run of empty paragraphs",
    evaluate(m) {
      const expected = "maxConsecutiveEmptyParas <= 1";
      if (!m.a) return probeMissing("a", expected);
      const runLen = m.a.pageBreaks?.maxConsecutiveEmptyParas;
      if (typeof runLen !== "number") return fail("maxConsecutiveEmptyParas absent", expected, "not measured");
      return (runLen <= 1 ? pass : fail)(`longest empty run ${runLen}`, expected);
    },
  },
  {
    id: "F12.B1",
    gap: "F12",
    probe: "B",
    description: "Page 1 carries no trailing empty paragraph; the footer closes the page",
    evaluate(m) {
      const expected = "sections[0].trailingEmptyPara = false";
      if (!m.b) return probeMissing("b", expected);
      const t = m.b.sections?.[0]?.trailingEmptyPara;
      if (typeof t !== "boolean") return fail("trailingEmptyPara absent", expected, "not measured");
      return (t === false ? pass : fail)(`trailingEmptyPara=${t}`, expected);
    },
  },
  {
    id: "F12.C1",
    gap: "F12",
    probe: "C",
    description: "Still exactly two pages, with the closer on page 1",
    evaluate(m) {
      const expected = `pages=${PAGES_EXPECTED}, closer on page 1`;
      if (!m.c) return probeMissing("c", expected);
      const closer = m.c.regions?.closer;
      if (typeof m.c.pages !== "number" || !closer) return fail(`pages=${m.c.pages ?? "?"}, closer ${closer ? "found" : "missing"}`, expected, "not found");
      const ok = m.c.pages === PAGES_EXPECTED && closer.page === 1;
      return (ok ? pass : fail)(`pages=${m.c.pages}, closer page ${closer.page}`, expected);
    },
  },

  // --- F13 — an image draws its variant; its placeholder is a picture ------
  {
    id: "F13.A1",
    gap: "F13",
    probe: "A",
    description: "The scan-to-pay image sits in a single-cell card table with D9DDEB border and 12pt padding",
    evaluate(m) {
      const expected = `inCardTable image; card border ${RULE_COLOR}, tcMar ${CARD_TCMAR_TWIPS} twips`;
      if (!m.a) return probeMissing("a", expected);
      const image = (m.a.images ?? []).find((i) => i.context === "body" && i.inCardTable === true);
      const card = m.a.cardTable;
      if (!image) return fail("no body image inCardTable", expected);
      if (!card) return fail("image in card, but cardTable border/margins not measured", expected, "not measured");
      const ok = card.borderColor === RULE_COLOR && card.tcMarTwips === CARD_TCMAR_TWIPS;
      return (ok ? pass : fail)(`border=${card.borderColor ?? "none"}, tcMar=${card.tcMarTwips ?? "none"}`, expected);
    },
  },
  {
    id: "F13.A2",
    gap: "F13",
    probe: "A",
    description: "No w:t anywhere matches /^\\[image: / — the text placeholder is gone",
    evaluate(m) {
      const expected = "imagePlaceholderTexts = []";
      if (!m.a) return probeMissing("a", expected);
      const texts = m.a.imagePlaceholderTexts;
      if (!Array.isArray(texts)) return fail("imagePlaceholderTexts absent", expected, "not found");
      return (texts.length === 0 ? pass : fail)(texts.length ? texts.map((t) => clip(t, 50)).join("; ") : "none", expected);
    },
  },
  {
    id: "F13.A3",
    gap: "F13",
    probe: "A",
    description: "Resolved and unresolved builds share the card box exactly; nothing after it moves",
    evaluate(m) {
      const expected = `card w/h identical; following regions move ${NO_REFLOW_TOL_PX}px`;
      if (!m.a) return probeMissing("a", expected);
      const base = m.a.cardTable;
      const variant = m.a.variants?.unresolvedImage?.cardTable;
      if (!base || !variant) return fail(`base card ${base ? "measured" : "missing"}, resolved variant ${variant ? "measured" : "missing"}`, expected, "not measured: the scan-to-pay image has no source until D14, so today's build IS the unresolved one and there is no resolved build to compare it against");
      const boxOk = base.widthTwips === variant.widthTwips && base.heightTwips === variant.heightTwips;
      const vRegions = m.b?.variants?.unresolvedImage?.regions;
      const baseRegions = m.b?.regions;
      let flowOk = null;
      const moved = [];
      if (vRegions && baseRegions) {
        flowOk = true;
        for (const name of ["terms", "footer-2"]) {
          const a = baseRegions[name];
          const b = vRegions[name];
          if (!a || !b) { flowOk = false; moved.push(`${name} unlocated`); continue; }
          if (Math.abs(a.y - b.y) > NO_REFLOW_TOL_PX) { flowOk = false; moved.push(`${name} Δy=${round2(Math.abs(a.y - b.y))}`); }
        }
      }
      if (flowOk === null) return fail(`box identical=${boxOk}; preview reflow not measured`, expected, "not measured");
      const ok = boxOk && flowOk;
      return (ok ? pass : fail)(`box identical=${boxOk}; ${moved.length ? moved.join(", ") : "nothing moved"}`, expected);
    },
  },
  {
    id: "F13.B1",
    gap: "F13",
    probe: "B",
    description: "The QR box is 143.6 +/- 4px square and the card 234.3 +/- 4px wide",
    evaluate(m) {
      const expected = `qr w,h ${QR_BOX_PX} ± ${QR_BOX_TOL_PX}px; card w ${CARD_WIDTH_PX} ± ${CARD_WIDTH_TOL_PX}px`;
      if (!m.b) return probeMissing("b", expected);
      const qr = m.b.regions?.["qr-canvas"];
      const card = m.b.regions?.["scan-card"];
      if (!qr || !card) return fail(`qr-canvas ${qr ? "found" : "missing"}, scan-card ${card ? "found" : "missing"}`, expected, "not found");
      const ok = near(qr.w, QR_BOX_PX, QR_BOX_TOL_PX)
        && near(qr.h, QR_BOX_PX, QR_BOX_TOL_PX)
        && near(card.w, CARD_WIDTH_PX, CARD_WIDTH_TOL_PX);
      return (ok ? pass : fail)(`qr ${round2(qr.w)}x${round2(qr.h)}px, card w ${round2(card.w)}px`, expected);
    },
  },
  {
    id: "F13.C1",
    gap: "F13",
    probe: "C",
    description: "Word reads the card table as 108pt wide (and as tall, +/- 2pt)",
    evaluate(m) {
      const expected = `w ${CARD_WIDTH_PT} ± ${CARD_WIDTH_TOL_PT}pt, h ${CARD_WIDTH_PT} ± ${CARD_HEIGHT_TOL_PT}pt`;
      if (!m.c) return probeMissing("c", expected);
      const card = m.c.scanCard;
      if (!card || typeof card.preferredWidth !== "number" || typeof card.heightPt !== "number") {
        return fail("scanCard width/height absent", expected, "not measured");
      }
      // 108pt is the picture, not the box around it: F13.A pins the card's
      // padding at 12pt a side, so a card holding a 108pt code measures 132.
      // Comparing the box against the picture was comparing two things.
      const padding = (CARD_TCMAR_TWIPS / 20) * 2;
      const inner = card.preferredWidth - padding;
      const innerHeight = card.heightPt - padding;
      const ok = near(inner, CARD_WIDTH_PT, CARD_WIDTH_TOL_PT)
        && near(innerHeight, CARD_WIDTH_PT, CARD_HEIGHT_TOL_PT);
      return (ok ? pass : fail)(
        `code ${round2(inner)} x ${round2(innerHeight)}pt inside a ${round2(card.preferredWidth)}pt card`,
        expected,
      );
    },
  },

  // --- F14 — a block can set its own measure -------------------------------
  {
    id: "F14.A1",
    gap: "F14",
    probe: "A",
    description: "The summary paragraph carries a 20mm right indent (w:ind 1134)",
    evaluate(m) {
      const expected = `ind right = ${SUMMARY_IND_RIGHT_TWIPS} twips`;
      if (!m.a) return probeMissing("a", expected);
      const ind = (m.a.inds ?? []).find((i) => typeof i.anchor === "string" && i.anchor.includes("Sprint 14"));
      if (!ind) return fail("summary ind absent", expected, "not found");
      return (ind.right === SUMMARY_IND_RIGHT_TWIPS ? pass : fail)(`right=${ind.right ?? "none"}`, expected);
    },
  },
  {
    id: "F14.B1",
    gap: "F14",
    probe: "B",
    description: "The summary renders 597.2 +/- 2px wide (the 158mm measure)",
    evaluate(m) {
      const expected = `w ${SUMMARY_WIDTH_PX} ± ${SUMMARY_WIDTH_TOL_PX}px`;
      if (!m.b) return probeMissing("b", expected);
      const r = m.b.regions?.summary;
      if (!r) return fail("summary region missing", expected, "not found");
      return (near(r.w, SUMMARY_WIDTH_PX, SUMMARY_WIDTH_TOL_PX) ? pass : fail)(`${round2(r.w)}px`, expected);
    },
  },
  {
    id: "F14.C1",
    gap: "F14",
    probe: "C",
    description: "Word reads RightIndent 56.7pt and no summary line exceeds 158mm",
    evaluate(m) {
      const maxLinePt = round2(mmToPt(SUMMARY_MAX_LINE_MM));
      const expected = `RightIndent ${SUMMARY_IND_RIGHT_PT} ± ${SUMMARY_IND_RIGHT_TOL_PT}pt; max line <= ${maxLinePt}pt`;
      if (!m.c) return probeMissing("c", expected);
      const s = m.c.summary;
      if (!s || typeof s.rightIndent !== "number" || typeof s.maxLineWidth !== "number") {
        return fail("summary rightIndent/maxLineWidth absent", expected, "not measured");
      }
      const ok = near(s.rightIndent, SUMMARY_IND_RIGHT_PT, SUMMARY_IND_RIGHT_TOL_PT)
        && s.maxLineWidth <= maxLinePt;
      return (ok ? pass : fail)(`indent ${s.rightIndent}pt, max line ${round2(s.maxLineWidth)}pt`, expected);
    },
  },

  // --- F15 — rounded blocks ------------------------------------------------
  {
    id: "F15.A1",
    gap: "F15",
    probe: "A",
    description: "The file holds real rounded geometry (roundrect/prstGeom) for the six rounded regions",
    evaluate(m) {
      const expected = `roundrects + prstRoundRects >= ${ROUNDED_REGIONS.length}`;
      if (!m.a) return probeMissing("a", expected);
      const g = m.a.roundedGeometry;
      if (!g) return fail("roundedGeometry absent", expected, "not found");
      const total = (g.roundrects ?? 0) + (g.prstRoundRects ?? 0);
      return (total >= ROUNDED_REGIONS.length ? pass : fail)(`vml=${g.roundrects ?? 0} + dml=${g.prstRoundRects ?? 0} = ${total}`, expected);
    },
  },
  {
    id: "F15.B1",
    gap: "F15",
    probe: "V",
    blockedBy: "P0",
    description: "Preview corner test: pill corners show the ground 2px inside the diagonal, centre shows the fill",
    evaluate: gatedByP0((m) => {
      const expected = "all corner-test entries rounded in the preview render";
      if (!m.v) return probeMissing("v", expected);
      return cornerVerdict(m.v.cornerTests?.preview, "preview", expected);
    }),
  },
  {
    id: "F15.C1",
    gap: "F15",
    probe: "V",
    blockedBy: "P0",
    description: "Word PDF raster corner test: same pixels, same verdict",
    evaluate: gatedByP0((m) => {
      const expected = "all corner-test entries rounded in the Word raster";
      if (!m.v) return probeMissing("v", expected);
      return cornerVerdict(m.v.cornerTests?.word, "word", expected);
    }),
  },
  {
    id: "F15.C2",
    gap: "F15",
    probe: "C",
    description: "Word's shape counts account for every rounded region",
    evaluate(m) {
      const expected = `Shapes + InlineShapes >= ${ROUNDED_REGIONS.length}`;
      if (!m.c) return probeMissing("c", expected);
      const shapes = m.c.shapesCount;
      const inline = m.c.inlineShapesCount;
      if (typeof shapes !== "number" || typeof inline !== "number") return fail(`shapes=${shapes ?? "?"}, inline=${inline ?? "?"}`, expected, "not measured");
      return (shapes + inline >= ROUNDED_REGIONS.length ? pass : fail)(`shapes=${shapes} + inline=${inline} = ${shapes + inline}`, expected);
    },
  },

  // --- D13 — totals through derivers, so the document publishes at all -----
  {
    id: "D13.A1",
    gap: "D13",
    probe: "A",
    blockedBy: "publish-build",
    description: "The publish-path build succeeds (reduce out of the initializer, into derivers)",
    evaluate(m) {
      const expected = "buildProjectEngineDocument succeeds";
      const p = m.build?.publish;
      if (!p) return blocked("publish build not run", expected, "requires --publish build");
      return p.ok === true
        ? pass("publish build ok", expected)
        : fail(`publish threw: ${clip(p.error)}`, expected);
    },
  },
  {
    id: "D13.A2",
    gap: "D13",
    probe: "A",
    blockedBy: "publish-build",
    description: "document.json carries the total as a derived token, not the baked figure",
    evaluate(m) {
      const expected = `contains ${DERIVED_TOTAL_TOKEN}; does not contain ${BAKED_TOTAL}`;
      const gate = publishBlocked(m.build, expected);
      if (gate) return gate;
      const text = publishDocumentText(m.build);
      if (text == null) return fail("document.json content not passed to evaluate", expected, "not measured");
      const hasToken = text.includes(DERIVED_TOTAL_TOKEN);
      const baked = text.includes(BAKED_TOTAL);
      return (hasToken && !baked ? pass : fail)(`token=${hasToken}, baked figure=${baked}`, expected);
    },
  },

  // --- D14 — the scan-to-pay code is a deriver, not a prompt ---------------
  {
    id: "D14.A1",
    gap: "D14",
    probe: "A",
    blockedBy: "publish-build",
    description: "document.json keeps the paymentQr deriver and its token, not a baked data URI",
    evaluate(m) {
      const expected = `contains "${QR_DERIVER_NAME}" and ${DERIVED_QR_TOKEN}`;
      const gate = publishBlocked(m.build, expected);
      if (gate) return gate;
      const text = publishDocumentText(m.build);
      if (text == null) return fail("document.json content not passed to evaluate", expected, "not measured");
      const hasDeriver = text.includes(QR_DERIVER_NAME);
      const hasToken = text.includes(DERIVED_QR_TOKEN);
      return (hasDeriver && hasToken ? pass : fail)(`deriver=${hasDeriver}, token=${hasToken}`, expected);
    },
  },
  {
    id: "D14.C1",
    gap: "D14",
    probe: "C",
    blockedBy: "publish-build",
    description: "The packed docx embeds a raster: Word reports exactly one inline shape in the card",
    evaluate(m) {
      const expected = "scanCard.inlineShapeCount = 1";
      if (!m.c) return probeMissing("c", expected);
      const n = m.c.scanCard?.inlineShapeCount;
      if (typeof n !== "number") return fail("scanCard.inlineShapeCount absent", expected, "not measured");
      return (n === 1 ? pass : fail)(`${n} inline shape(s)`, expected);
    },
  },

  // --- C6 — page numbers filled after layout -------------------------------
  {
    id: "C6.B1",
    gap: "C6",
    probe: "B",
    description: "The baked preview's footers read '1 / 2' and '2 / 2'",
    evaluate(m) {
      const expected = PAGE_FIELD_TEXTS.join(" then ");
      if (!m.b) return probeMissing("b", expected);
      const s = m.b.sections;
      if (!Array.isArray(s) || s.length < 2) return fail(`${Array.isArray(s) ? s.length : 0} sections`, expected, "not found");
      const texts = s.map((sec) => String(sec.pageNumber?.text ?? "").replace(/\s+/g, " ").trim());
      const ok = PAGE_FIELD_TEXTS.every((t, i) => texts[i] === t);
      return (ok ? pass : fail)(texts.map((t) => JSON.stringify(t)).join(", "), expected);
    },
  },
  {
    id: "C6.C1",
    gap: "C6",
    probe: "V",
    description: "Word's rendered pages agree: '1 / 2' on page 1, '2 / 2' on page 2",
    evaluate(m) {
      const expected = PAGE_FIELD_TEXTS.join(" then ");
      if (!m.v) return probeMissing("v", expected);
      const texts = m.v.footerTextByPage;
      if (!Array.isArray(texts) || texts.length < 2) return fail("footerTextByPage absent", expected, "not measured");
      const ok = PAGE_FIELD_TEXTS.every((t, i) => typeof texts[i] === "string" && texts[i].includes(t));
      return (ok ? pass : fail)(texts.map((t) => JSON.stringify(clip(t, 20))).join(", "), expected);
    },
  },
];

// ---------------------------------------------------------------------------
// Gates — plan §5 / contract "Objectives", same shape as OBJECTIVES.
// ---------------------------------------------------------------------------

/** @type {Array<{id: string, gap: string, probe: string, blockedBy?: string, description: string, evaluate: (m: object) => object}>} */
export const GATES = [
  {
    id: "G1",
    gap: "G1",
    probe: "C",
    description: "Word reports exactly two pages",
    evaluate(m) {
      const expected = PAGES_EXPECTED;
      if (!m.c) return probeMissing("c", expected);
      if (typeof m.c.pages !== "number") return fail("pages absent", expected, "not found");
      return (m.c.pages === PAGES_EXPECTED ? pass : fail)(m.c.pages, expected);
    },
  },
  {
    id: "G2",
    gap: "G2",
    probe: "B",
    description: "Preview section 1 height is 1000-1124px (fits a page, has not collapsed)",
    evaluate(m) {
      const expected = `h in [${G2_SECTION1_PX}]px`;
      if (!m.b) return probeMissing("b", expected);
      const h = m.b.sections?.[0]?.h;
      if (typeof h !== "number") return fail("section 1 height absent", expected, "not found");
      return (inRange(h, G2_SECTION1_PX) ? pass : fail)(`${round2(h)}px`, expected);
    },
  },
  {
    id: "G3",
    gap: "G3",
    probe: "B+C",
    description: "Every region within +/-1mm of the design, in the preview (px) and in Word (pt, x/y)",
    evaluate(m) {
      const expected = `all ${REGION_NAMES.length} regions ±${TOL_MM_PX}px in B; all ${REGION_NAMES.length - GEOMETRIC_REGIONS.size} anchored regions ±${TOL_MM_PT}pt in C`;
      if (!m.b && !m.c) return fail("measure-b and measure-c absent", expected, "probe missing");
      const bMisses = [];
      let bHits = 0;
      for (const name of REGION_NAMES) {
        const check = regionsClose(m.b?.regions?.[name], DESIGN_REGIONS[name], TOL_MM_PX);
        if (check.ok) bHits += 1;
        else bMisses.push(`${name}(${check.detail})`);
      }
      const cNames = REGION_NAMES.filter((n) => !GEOMETRIC_REGIONS.has(n));
      const cMisses = [];
      let cHits = 0;
      for (const name of cNames) {
        const d = DESIGN_REGIONS[name];
        const r = m.c?.regions?.[name];
        const ok = r
          && r.page === d.page
          && near(r.x, pxToPt(d.x), TOL_MM_PT)
          && near(r.y, pxToPt(d.y), TOL_MM_PT);
        if (ok) cHits += 1;
        else cMisses.push(name);
      }
      const measured = `B ${bHits}/${REGION_NAMES.length}, C ${cHits}/${cNames.length}`;
      const ok = bHits === REGION_NAMES.length && cHits === cNames.length;
      const note = ok ? undefined : clip(`off: B[${bMisses.map((s) => s.split("(")[0]).join(", ")}] C[${cMisses.join(", ")}]`, 300);
      return ok ? pass(measured, expected) : fail(measured, expected, note);
    },
  },
  {
    id: "G4",
    gap: "G4",
    probe: "V",
    blockedBy: "P0",
    description: "Per-visual-region pixel diff vs the design <= 3% (scan-card exempt)",
    evaluate: gatedByP0((m) => {
      const expected = `every visual region pctDiff <= ${G4_MAX_PCT}%`;
      if (!m.v) return probeMissing("v", expected);
      const entries = m.v.regions;
      if (!Array.isArray(entries)) return fail("region diffs absent", expected, "not found");
      const byName = new Map(entries.map((e) => [e.region, e]));
      const missing = VISUAL_REGIONS.filter((n) => !byName.has(n));
      const over = entries
        .filter((e) => !G4_EXEMPT.has(e.region) && !(typeof e.pctDiff === "number" && e.pctDiff <= G4_MAX_PCT))
        .map((e) => `${e.region}=${round2(e.pctDiff ?? NaN)}%`);
      const worst = Math.max(0, ...entries.filter((e) => !G4_EXEMPT.has(e.region)).map((e) => e.pctDiff ?? 0));
      const ok = missing.length === 0 && over.length === 0;
      const measured = `worst ${round2(worst)}%; over: ${over.length ? over.join(", ") : "none"}; unmeasured: ${missing.length ? missing.join(", ") : "none"}`;
      return (ok ? pass : fail)(clip(measured, 300), expected);
    }),
  },
  {
    id: "G5",
    gap: "G5",
    probe: "build",
    description: "Repo npm test and documents:check are green (runner fills this in under --full)",
    evaluate(m) {
      const expected = "testsOk and documentsCheckOk";
      const checks = m.build?.fullChecks;
      if (!checks) return blocked("not run", expected, "runs only under --full; the runner fills this in");
      const ok = checks.testsOk === true && checks.documentsCheckOk === true;
      return (ok ? pass : fail)(`tests=${checks.testsOk}, documents:check=${checks.documentsCheckOk}`, expected);
    },
  },
  {
    id: "G6",
    gap: "G6",
    probe: "build",
    description: "Invariants I1-I4 are clean",
    evaluate(m) {
      const expected = "I1..I4 all clean";
      const inv = m.build?.invariants;
      if (!inv) return fail("invariants not provided by the runner", expected, "not measured");
      const dirty = ["I1", "I2", "I3", "I4"].filter((k) => inv[k] !== "clean");
      return (dirty.length === 0 ? pass : fail)(dirty.length ? `dirty: ${dirty.map((k) => `${k}=${inv[k] ?? "missing"}`).join(", ")}` : "all clean", expected);
    },
  },
];

// ---------------------------------------------------------------------------
// Preconditions — plan §1.4. P0 gates every pixel-level objective.
// ---------------------------------------------------------------------------

/** @type {Array<{id: string, gap: string, probe: string, description: string, evaluate: (m: object) => object}>} */
export const PRECONDITIONS = [
  {
    id: "P0",
    gap: "P0",
    probe: "B+C",
    description: "The same face resolves in Chrome and in Word, proven by measurement in both",
    evaluate: p0Verdict,
  },
];
