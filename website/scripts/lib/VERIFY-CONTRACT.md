# Invoice v2 verification harness — the contract

The spec every probe and the runner are written against. The rationale lives in
`website/design/invoice-v2-plan.md` (§1 defines done, §3 the objectives, §5 the
loop); this file is the interfaces and the numbers. Where this contract and the
plan disagree on a number, **this contract wins** — its geometry comes from the
measured design fixture, the plan's early estimates came from a Segoe-fallback
render before Aptos was installed (design charge rows are 54.1px, not 44).

## Files

```text
website/scripts/verify-invoice.mjs        # runner: build -> probes -> objectives -> report; exit 1 unless all PASS
website/scripts/lib/verify-build.mjs      # builds .verify/invoice.docx + .verify/invoice.html (+ variant builds)
website/scripts/lib/objectives.mjs        # every objective + threshold; pure functions over measurements
website/scripts/lib/probe-ooxml.mjs       # probe A -> .verify/measure-a.json
website/scripts/lib/probe-preview.mjs     # probe B -> .verify/measure-b.json
website/scripts/lib/probe-word.ps1        # probe C (PowerShell/COM), emits JSON on stdout
website/scripts/lib/probe-word.mjs        # spawns the ps1, validates -> .verify/measure-c.json, exports .verify/invoice.pdf
website/scripts/lib/probe-visual.mjs      # probe V: pdf raster vs design fixtures -> .verify/measure-v.json
website/scripts/lib/page-only-style.fixture  # I1: byte copy of PAGE_ONLY_STYLE
website/.verify/                          # all outputs; gitignored
website/.verify/report.json               # the board
website/.verify/history.json              # prior statuses + objectives hash + attempts[]
```

npm script (website/package.json): `"verify:invoice": "node scripts/verify-invoice.mjs"`.
`--full` additionally runs gate G5 (repo `npm test` + `documents:check`); default
run skips G5 for loop speed. Add `website/.verify/` to the repo `.gitignore`.

## Environment facts (verified, do not re-litigate)

- Chrome: `C:/Program Files/Google/Chrome/Application/chrome.exe`. Headless
  pattern that works here: `--headless=new --disable-gpu --hide-scrollbars
  --window-size=W,H --virtual-time-budget=15000` with `--screenshot=` or
  `--dump-dom`. A measuring page MUST be served same-origin with what it
  iframes (file:// parent + http:// iframe is blocked — this bit us already).
- Local static server: copy the pattern in `website/scripts/lib/` from the demo
  scripts, or a ~20-line `node:http` server. Use **port 8901**. Always kill the
  server (and any Chrome you spawned) before exiting, success or failure.
- Word 16.0.20326 via COM from PowerShell. `New-Object -ComObject
  Word.Application`; `Visible=$false; DisplayAlerts=0`; open read-only with
  `AddToRecentFiles:=$false`; close `$d.Close(0)`, `$w.Quit()`, release COM
  refs; on any failure path `Stop-Process -Name WINWORD -Force` as last resort.
  Constants: Information(3)=page number, (5)=x from page edge in points,
  (6)=y from page edge in points; ComputeStatistics(2)=pages;
  Headers(1)=primary, Headers(2)=first page; ExportAsFixedFormat($pdf, 17);
  wdAlignParagraphRight=2; wdCellAlignVerticalCenter=1; wdLineStyleNone=0.
- Fonts: Aptos Regular/Bold/Italic/BoldItalic are installed per-user (copied
  from Office's cloud-font cache to `%LOCALAPPDATA%\Microsoft\Windows\Fonts` +
  HKCU registration). Word resolves Aptos AND Chrome resolves Aptos (measured:
  canvas width 297.2 vs 329.9 monospace fallback). P0 re-verifies each run.
- PDF rasterising: pdfjs-dist is installed in website/node_modules. **Render in
  headless Chrome**, not Node (Node has no canvas): serve a harness page that
  loads pdfjs from node_modules, renders each PDF page to canvas at scale
  96/72, `toDataURL()` into the DOM, then `--dump-dom` and decode in Node with
  pngjs. Set `GlobalWorkerOptions.workerSrc` to the served worker file.
- Node ESM on Windows: `import` specifiers must be relative or `file://` URLs —
  a bare `C:/...` path throws ERR_UNSUPPORTED_ESM_URL_SCHEME.
- docx-preview note: it reserialises styles, so colours read back as
  `rgb(r, g, b)`, never hex. Match text anchors, not style strings, wherever
  possible.

## The build step (verify-build.mjs)

Same bundling recipe as `website/scripts/build-demo-document.mjs` (esbuild,
`docxcelerateEsbuildTransform`, `packages: "external"`), pointed at
`website/src/demo/documents/invoice/document.project.ts`:

1. `buildProjectPreviewDocument(project)` -> model.
2. `createDocxBlob(model)` (from `docxcelerate/docx`) -> `.verify/invoice.docx`.
3. `renderDocxPage(model, { title, style: PAGE_ONLY_STYLE })` (from
   `./lib/docx-page.mjs`) -> `.verify/invoice.html`.
4. Variant builds behind flags (used by specific objectives; skip when the flag
   is off so the default loop stays fast):
   - `--lines=3` / `--lines=30`: clone previewData with truncated/extended
     lines (extend by repeating existing lines with numbered descs) -> F8.A3.
   - `--publish`: build via the publish path (see
     `skills/docxcelerate/references/publishing.md`; artifact build API in
     `src/project/artifact.ts`) -> writes `.verify/publish/` artifacts ->
     D13/F8.A2. Until D13 lands this build THROWS (reduce is refused) — the
     runner records that as status BLOCKED for the objectives that need it,
     with the thrown message as evidence.
   - `--unresolved-image`: previewData variant that leaves the scan-to-pay
     image unresolved -> F13.A3 no-reflow comparison.

## Probe outputs

Every probe writes one JSON file and never throws on a missing region: a region
that cannot be located is recorded as `null` (probe B/C) or the fact list notes
its absence (probe A). Objectives treat `null` as FAIL with note "not found".

### measure-a.json (OOXML facts)

Unzip `.verify/invoice.docx` (port the central-directory reader from
`tests/docx.ts` — ~60 lines, plus `inflateRawSync`). Parse with regex/string
scanning (no XML dependency). Collect at least:

```jsonc
{
  "parts": ["word/document.xml", "..."],
  "sectPr": { "titlePg": false, "headerRefs": [{"type": "default", "part": "header1.xml"}], "footerRefs": [] },
  "headings1": ["Invoice details", "Parties", "..."],            // text of every Heading1-styled w:p
  "headerParts": { "header1.xml": { "text": "...", "hasTable": true } },
  "footerParts": { "footer1.xml": { "text": "...", "tables": [{ "tblInd": null, "widthsTwips": [..], "cells": [...] }], "pageFieldPara": { "jc": null, "spacingAfter": 120, "hasPageField": true, "hasNumPagesField": true, "inlineDrawingWithText": false } } },
  "bodyTables": [ { "anchor": "Issue date", "tblInd": null, "widthsTwips": [...], "rows": [ { "header": true, "cells": [ { "shd": "2C3D8F", "vAlign": null, "borders": {"bottom": {"color": "D9DDEB", "sz": 6}}, "runs": [{"font": null, "sz": 20, "color": "FFFFFF", "caps": true, "spacing": null, "text": "Issue date"}], "paraSpacing": {"line": 348, "after": 0} } ] } ] } ],
  "images": [ { "context": "body|footer", "hasDrawing": true, "extentEmu": {"cx":0,"cy":0}, "inCardTable": false } ],
  "imagePlaceholderTexts": ["[image: Scan to pay invoice INV-2026-0142]"],  // every w:t matching /^\[image: /
  "pageBreaks": { "loneBreakParas": 1, "pageBreakBefore": [] },
  "inds": [ { "anchor": "Sprint 14", "left": null, "right": null } ],
  "stylesXml": { "heading1": { "sz": 15, "spacing": null, "caps": true, "color": "2C3D8F" } },
  "roundedGeometry": { "roundrects": 0, "prstRoundRects": 0 },
  "sectionsInModel": 7                                            // from the built model JSON, passed in by the runner
}
```

Shape may grow; keep keys stable once shipped (objectives read them).

### measure-b.json (preview geometry, px @ 96dpi, page-relative)

Serve `.verify/invoice.html` + a measuring harness (same-origin). docx-preview
emits one `section.docx` per page. For each region in the region table: locate
by anchor (smallest element containing the anchor text), record
`{x, y, w, h}` relative to its section's content box, plus which section index
it landed in. Also record: per-section `{w, h}`, header/footer element heights
per section, charge row heights, computed `font-family` of an amount cell and a
description cell, the font probe (Aptos vs monospace canvas widths), and
`pageNumberText` per section footer. Colour-strip regions (`rule`, `p2-rule`,
`charges-rule`) are located geometrically: full-width elements with background
`rgb(44, 61, 143)` / `rgb(217, 221, 235)` and height ≤ 5px.

### measure-c.json (Word ground truth, points, page-relative)

For each anchored region: `{ page, x, y }` from Find + Information(5/6) —
collapse the found range to its start for x,y; record `xEnd` (range collapsed
to end) where the objective needs a right edge (amount column). Also: `pages`,
`firstPageHeaderText`, `primaryHeaderText`, footer text + line count, table
count, per-table `PreferredWidth` (points) and cell shading (`RGB` ints) for
the charges table rows 1..8 col 1, cell `VerticalAlignment` for band/footer
cells, `Font.Name`/`Font.Spacing` samples (amount cell run, Heading 1 style),
paragraph `Alignment`/`SpaceAfter` for the page-number paragraph, fields text
after `Fields.Update()` (footer on each page), `InlineShapes.Count`,
`Shapes.Count`, and the two terms paragraphs' font size + colour. Write the
PDF export before quitting.

### measure-v.json (pixels)

Raster `.verify/invoice.pdf` at 96dpi -> page PNGs (expect 794x1123 ±2px;
record actual). For every region with `visual: true` in the region table, crop
the region rect (from `design-regions.json`) out of both the design fixture
and the Word raster, run pixelmatch (threshold 0.1), record
`{region, pctDiff, pixels}`. `scan-card` and `qr-canvas` are exempt (geometry
only). Also compare the docx-preview screenshot (probe B can save one) against
the Word raster for the same regions -> `crossEngine` diffs (informational
until G4, gating for "preview == Word" claims).

## The region table

Geometry ground truth: `website/design/invoice-v2/fixtures/design-regions.json`
(px @ 96dpi, page-relative, content-box). Conversions: mm = px / 3.77953;
pt = px * 0.75. Word tolerance ±1mm = ±2.83pt; preview tolerance ±1mm = ±3.78px
unless an objective narrows it.

| region | page | anchor (A/B/C locate by) | visual |
| --- | --- | --- | --- |
| letterhead | 1 | `Software consultancy · Manchester` | yes |
| rule | 1 | geometric: navy strip h≤5px | yes |
| band | 1 | `Issue date` | yes |
| status-pill | 1 | `Awaiting payment` | yes |
| parties | 1 | `Billed to` | yes |
| summary-label | 1 | `Engagement summary` | yes |
| summary | 1 | `Sprint 14 closed out` | yes |
| charges-head | 1 | `Description` | yes |
| charges-body | 1 | `Discovery and scoping workshop` … `Production support retainer` (union) | yes |
| charges-rule | 1 | geometric | no |
| totals-panel | 1 | `Subtotal` | yes |
| total-bar | 1 | `Total due` | yes |
| closer | 1 | `are on page 2` | yes |
| footer | 1 | `Registered in England` (footer part / page footer) | yes |
| p2-letterhead | 2 | `PAYMENT` | yes |
| p2-rule | 2 | geometric | no |
| bank-grid | 2 | `Sort code` | yes |
| reference-panel | 2 | `Payment reference` | yes |
| terms | 2 | `Payment within 14 days` | yes |
| scan-card | 2 | `Scan to pay` | **geometry only** |
| footer-2 | 2 | footer on page 2 | yes |

The design chips (`AI · summary.engagement`, `Placeholder · deriver:
payment.qr`) do NOT print (plan D12): mask their rects out of the design-side
crops for `summary-label` and `scan-card`… `summary-label`'s chip sits right of
the label — mask the chip rect (probe V computes it from the fixture once:
the bordered pill following the label) to neutral page colour on the design
side before diffing. Same for the QR caption chip inside `scan-card` (exempt
anyway). The design's canvas QR is random modules — never pixel-compare it.

## Objectives

IDs are `<gap>.<probe><n>` (`F1.A1`, `F1.B1`, `F1.C1`, `P0.B1`, `G1.C1`…).
Statuses: PASS / FAIL / BLOCKED / SUSPECT. An objective whose probe input is
missing (probe crashed) is FAIL with note "probe missing". `blockedBy` marks
dependencies: F8.A2, D13.*, D14.* carry `blockedBy: "publish-build"`; every
`visual` objective carries `blockedBy: "P0"`. BLOCKED counts as not-PASS.

Implement every objective from plan §3 (F1–F15), §4 (D13, D14, C6) and the
gates, with these fixture-corrected numbers where the plan estimated:

- F4.B: footer element height in [52, 58]px (design 54.78). F4.C: ≤ 14.8mm, 1 line.
- F7.B: every charge row height in [50, 58]px (design 54.11); `charges-body`
  height in [370, 388]px (design 378.77). F7.C: charges head+body ≤ 111mm
  (design head 31.84 + body 378.77 = 410.6px = 108.7mm).
- F7.A: charge-row paragraph `w:spacing` line < 348 (the body 1.45 default);
  muted note line < its cell sibling's.
- F5.B: page-number right edge within 2px of design footer text right inset
  (x = 733.2px page-relative).
- F13.B: qr box 143.6±4px square; card width 234.3±4px. F13.C: card ≈ 108pt wide.
- F14.B: summary width 597.2±2px.
- F15: rounded regions = {status-pill, charges-head, totals-panel, total-bar,
  reference-panel, scan-card} (6; the F mark tiles are SVG assets and carry
  their own radius). F15.B/V corner test: at each pill corner, sample the
  pixel 2px inside the bounding-box corner along the diagonal; it must match
  the ground (band fill / page white), and the centre must match the region
  fill. Run against BOTH the preview render and the Word PDF raster.
- G2: preview section 1 height in [1000, 1124]px.
- G3: every located region within ±1mm of design-regions, in B (px) and C (pt,
  x/y only). Regions not yet locatable are FAIL not error.
- G4: per-visual-region pixel diff ≤ 3%, none > 5%; requires P0; scan-card
  exempt.
- I4 (supersedes the plan's rough "≥52 chars"): preview-data.ts has ≥ 7 lines,
  `max(desc.length) ≥ 30`, `max(meta.length) ≥ 50`, and the grand total the
  fixture shows (`£22,380.00`) unchanged.

## Invariants (runner)

- I1: sha256 of the `PAGE_ONLY_STYLE` string exported by `docx-page.mjs` equals
  the checked-in `page-only-style.fixture` (create the fixture from the current
  value when authoring the runner; it is then frozen). Mismatch -> every B/V
  objective SUSPECT. The C6 field-filling change edits JS in `docx-page.mjs`
  but NOT this constant.
- I2: sha256 of `objectives.mjs` recorded per run in history.json. If an
  objective goes FAIL->PASS in a run whose objectives hash differs from the
  previous run's, mark it SUSPECT (note: "threshold changed in same run").
- I3: the board is one run; the runner never merges statuses across runs.
- I4: as above; failure -> every F7/G1/G2 objective SUSPECT (the numbers could
  have been reached by shrinking data).

## Report

`.verify/report.json`:

```jsonc
{
  "runId": "vr-<epoch>", "startedAt": "...", "finishedAt": "...",
  "objectivesSha256": "...",
  "preconditions": { "P0": { "status": "PASS", "chromeAptos": 297.2, "monoFallback": 329.9, "wordAptos": true } },
  "invariants": { "I1": "clean", "I2": "clean", "I3": "clean", "I4": "clean" },
  "probes": { "A": { "ok": true }, "B": { "ok": true }, "C": { "ok": true }, "V": { "ok": true } },
  "objectives": [ { "id": "F1.A1", "gap": "F1", "status": "FAIL", "measured": "...", "expected": "...", "note": "" } ],
  "gates": { "G1": { "status": "FAIL", "measured": 3, "expected": 2 } },
  "summary": { "pass": 0, "fail": 0, "blocked": 0, "suspect": 0 }
}
```

`history.json` keeps `{ runs: [{runId, objectivesSha256, statuses: {id: status}}], attempts: [] }`;
`attempts[]` entries are appended by hand (or by the fixing agent) per plan §5:
`{ objective, mechanism, probeRejected, measured, when }`.

Runner console output: one line per objective (`PASS F1.A1 — no unwanted
Heading1`), then the gate lines, then `BOARD: n pass / n fail / n blocked / n
suspect`. Exit 0 only when everything is PASS and invariants are clean.

## What the harness must never do

Edit anything under `src/`, `registry/`, `skills/`, or
`website/src/demo/documents/invoice/` — the harness measures the document, it
does not improve it. The only files it touches are listed under **Files**, plus
the `website/package.json` scripts block and `.gitignore`. The expected first
board is nearly all-red; that is correct and is not a bug to fix.
