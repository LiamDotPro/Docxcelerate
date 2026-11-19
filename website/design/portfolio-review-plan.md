# Portfolio review — the design, the gaps, and the build

Design: `PortfolioReport.dc.html` (13 A4 pages) in the Docxcelerate design
project `c749e446-5224-4fdc-a711-cc33093d48b6`, copied to
`design/portfolio-review/`. Implementation: `src/demo/documents/portfolio-review/`.
It replaces `offer-of-admission` in the homepage hero.

## Decisions taken

- **Fonts: Georgia and Arial.** The canvas asks for Cambria and Aptos, but on
  this Mac neither is installed and it renders Georgia (headings, letter) and
  Arial (everything else). Both ship with Word on Mac and Windows, so the file
  names them outright and the preview and Word draw the same faces.
- **Signature: Brush Script MT**, the canvas's first choice, present on both.
- **Cover photo: baked.** The canvas hot-links an Unsplash photo and tints it in
  CSS (grayscale, multiply #2f8a6b, a gradient to #0e3b2d). A `.docx` can do
  none of that, so the composite is rendered once from the canvas and committed
  as a JPEG asset.
- **Two builds.** `public/demo/portfolio-review.html` is the preview build
  (placeholders where the engine writes), compared with the canvas in
  `mode: "preview"`. `public/demo/portfolio-review.engine.html` is built with a
  fixture `AiClient` returning the canvas's ENGINE text, compared with
  `mode: "engine"`. Both compare against `highlightAi: false` — the green wash
  is an on-screen highlight, and no file prints it.

## How parity is measured

`npm run parity -- portfolio-review [--engine] [--pages 3,4]`
(`scripts/design-parity.mjs`) photographs the canvas and the preview in the
same headless Chrome at 96dpi (A4 = 794×1123), page by page, and writes
`.parity/<id>/index.html`: design, preview, pixel difference, and an overlay with
an opacity slider, per page, with a score.

- **pixel** — share of pixels pixelmatch calls the same (threshold 0.15).
- **layout** — both pages shrunk 8×, grey, mean difference. Forgives glyph
  rasterising; punishes wrong placement, wrong fill, missing blocks.
- **score** — their mean. The headline is the mean over pages.

Word itself cannot be measured on this Mac (no Word), so the `.docx` is checked
by its XML (Probe A idiom, `tests/`) and by the preview, which reads that file.

## Inventory: what each page needs, and what carries it

| Page | Blocks | Carried by |
| --- | --- | --- |
| 1 Cover | photo band, letterhead, title, prepared-for, value, decision stamp, headline, contents, sign-off, footer | **G1** floating image; first-page header; tables; static contents |
| 2 Letter | address/date, greeting, heading, drop-cap letter (AI ×4), signature, contact strip | **G2** drop cap; Text variant font for the signature; `signature-block` (edited) |
| 3 Summary | section heading, six key figures, recommendation (stamp + AI + trades), next steps, three AI points | tables with top rules; condition on rebalance; numbered list |
| 4 Performance | line chart with area + Q3 band + end labels, returns table, calendar-year clustered bars | **G3** chart features; table |
| 5 Attribution | two diverging bar lists, top/bottom table | **bars as table cells** (nested table per bar) |
| 6 Allocation | drift bars with range + target marker, region bars, 100% currency bar + key | bars as table cells |
| 7 Holdings | dark-header table, 8 rows (AI note each), account types | table; loop; key figures |
| 8–9 Focus | six fund panels: figure, sparkline, facts, position bars, AI "why" | small line chart; bars as cells; loop |
| 10 Income | 12-bar chart (quarter months dark), two tables, cash bar | **G3** per-point colour + data labels; bars as cells |
| 11 Activity | waterfall bridge, transactions table | bars as cells (waterfall = offset segment) |
| 12 Costs | costs table, five figures, before/after line chart with band, big figure | **G3** dashed series, band |
| 13 Plan | fan chart (band + central + dashed goal + end labels), decision, allowance bars, important info, glossary | **G3**; bars as cells; `terms-notice` (edited) |
| all | running header band (bleeds, dark), rule, footer rule, PAGE n OF N | header/footer; **G4** page-number styling |

## Framework gaps

Ordered by how much of the page they block. Each is a general capability, not
a portfolio special case.

- **G1 Floating image behind text.** `<Image float={{ xMm, yMm, behind }}>`
  packs a `wp:anchor` positioned from the page. Needed for the cover photo, and
  generally for watermarks and backgrounds. docx-preview reads anchors but
  positions them from the paragraph and ignores `behindDoc`; `docx-page.mjs`
  writes the page position back (the sanctioned "JS writing back what the file
  says" category).
- **G2 Drop cap.** A block property `dropCapLines` packs `w:framePr
  w:dropCap="drop"` on a one-letter paragraph. docx-preview floats it left.
- **G3 Chart styling.** Per-series `dash`, per-point `colors`, `gridlines:
  false`, value-axis hidden, data labels on bars; and **combo** — a line drawn
  over an area or band in one chart (performance, costs, plan). Needed in the
  packer (`chart_part.ts`) and in the preview drawer (`charts.ts`).
- **G4 Page numbers.** The run styling of the cell (colour, size, tracking) has
  to reach the field runs, and the preview's page-number fill only handles a
  `/` separator; "PAGE 3 OF 13" needs any separator.
- **G5 Dynamic labels in preview.** The canvas's preview mode prints
  "Dynamic: <node id>" above every engine-written block. A preview build marks
  each dynamic node with that label when the style declares a `dynamicLabel`
  block — opt-in, so the other demos do not change.
- **Not a gap, by construction:** drawn bars (drift, diverging, waterfall,
  stacked, progress, positions) are table cells with fills and mm widths, which
  Word and docx-preview both draw exactly. Shapes stay out of it.
- **Deferred:** a real TOC field (docx-preview evaluates no fields; the contents
  list is static, with the design's page numbers); per-section margins (the
  cover's different top is carried by the first-page header instead).

## Registry

Reused (edited into the document, the way `dxcl add` intends): `signature-block`
(page 2 sign-off), `terms-notice` (page 13 important information),
`next-steps` (shape replaced: the canvas wants three numbered steps, not one
paragraph). Others in the registry (`balance-trend`, `period-comparison`,
`status-banner`) do not match what the canvas draws.

New registry components, promoted from the document once each one renders at
parity — the first three are document-agnostic:

- `key-figures` — a grid of labelled figures under a top rule (pages 3, 7, 12).
- `bar-list` — labelled horizontal bars as table cells, with optional diverging
  axis and value column (pages 5, 6, 8–10, 13).
- `value-bridge` — opening value, signed steps, closing value as a waterfall,
  with a check that the steps sum (page 11).
- `allocation-drift` — current vs target vs agreed range, with post-trade mix
  (page 6).
- `cost-disclosure` — the regulated costs-and-charges table and its
  twelve-month summary (page 12).

## The loop

1. Build all 13 pages with what exists; bars as cells; charts as plain Graphs.
2. Run parity; take the lowest-scoring page; find the largest visible
   difference in its diff; fix the document or close the gap that causes it.
3. After each framework change: `npm test` at the root, and the docx checked
   by XML for the property it adds.
4. Repeat until no page improves.

## Where it stands (2026-11-19, second round)

| Build | Round 1 | Round 2 | Pages |
| --- | ---: | ---: | --- |
| Preview (`portfolio-review.html`) vs canvas in preview mode | 95.5% | **96.9%** | 13 |
| Engine (`portfolio-review.engine.html`) vs canvas in engine mode | 95.0% | **96.5%** | 13 |
| Word's own PDF of the engine `.docx` vs canvas | 95.1% | **96.5%** | 13 in Word |

Every page now scores 96–98% in the preview.

Round 2 framework changes, tested in `tests/layout_fidelity.test.tsx`:

- **G3 charts**: per series `dash`, `band` (packed as Word's stacked-area band), `colors` per point, `widthPt` (0 draws no line), `endLabel`; per chart `markers`, `gridlines`, `axes`, `textSizePt`, `plotArea` (a manual plot layout both Word and the preview honour), `fromEdge` (midCat). The preview reader and the echarts drawer read all of them back, and the drawer now prints number-format suffixes, `;` sections and `[>=N]` conditions.
- **The site's paginator script never ran.** It inlined `docx_paginate.js`, whose import of `docx_fields.js` cannot resolve in an inline module, and called `applyTabStops`, which it never included. So on main every hero preview has had no tab stops and no repagination. `docx-page.mjs` now inlines all three modules, each in its own scope.
- Right tab stops measure trailing text to the end of the paragraph, not of the tab's own span.
- The preview sets each paragraph at its runs' size, so a line is exactly its stated leading (the CSS strut added ~1px a row).
- Nested tables share `"auto"` columns from their cell's width, not the page's.
- A cell's `align` reaches its text; stand-in labels take their paragraph's space-before and print the id as written, not the loop-numbered one.

Measurement: the canvas draws a 1px border round pages 2–13; the parity tool now measures from inside it.

Word-specific finding: Word adds a cell's bottom margin to a row of exact height. The cover's title block is held off the photograph's foot by a paragraph, not padding, for that reason.

## Round 3 (2026-11-19)

| Build | Round 2 | Round 3 | Pages |
| --- | ---: | ---: | --- |
| Preview | 96.9% | **97.2%** | 13 |
| Engine preview | 96.5% | **96.9%** | 13 |
| Word's own PDF | 96.5% | **96.7%** | 13 in Word |

- **G2 drop cap**: block `dropCap: { lines, variant }` packs a `w:framePr w:dropCap="drop"` first letter; written text only.
- Numbered lists: `listStyle: "decimalPlain"` (a bare figure) and block `marker` (the variant the numerals are set in). A written passage split into paragraphs counts on as one list.
- A paragraph border's `w:space` is taken per side from `paddingSidesPt` (it read only `paddingPt`, so a hairline's padding below was silently dropped).
- The preview adds a paragraph's space after to the next one's space before, as Word does (CSS was collapsing them to the larger).
- The site's font fallbacks no longer split a quoted family name at the `;` of `&quot;` (the signature was falling back to a serif); script faces fall back to `cursive`.
- Bars inside cells give back the point Word's trailing paragraph adds; the parity tool rasterises both sides with the same font smoothing.

Left, largest first:

1. The performance chart's Q3 shaded region (a category-range highlight; no chart option for it yet).
2. The canvas's preview mode puts "Dynamic: cover-headline" above the "This quarter" label; the label is per paragraph, so it sits below.
3. Per-page spacing, a few points at a time, with `parity:steps` (pages 3, 5, 6, 8, 11, 13 sit at 96%).
4. Word sets tracked capitals a little wider than the browser; the holdings page is where it shows most (94.7% in Word).
6. Per-page spacing tune with `parity:steps`.
7. Promote `key-figures`, `bar-list`, `value-bridge`, `allocation-drift`, `cost-disclosure` to the registry.
