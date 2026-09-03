# Word conformance — a plan

**The problem.** We have one renderer (pack a `.docx`, read it back with
docx-preview) and one very good harness that measures exactly one document: the
invoice. `website/scripts/verify-invoice.mjs` and its four probes know the
invoice's regions, the invoice's design fixture, the invoice's thresholds. When
we want to answer "does Word do the right thing with a vertically merged cell",
there is nowhere to put the question.

**The plan.** Lift the harness out of the invoice and turn it into a
conformance suite: one small document per Word feature, each producing the same
three artefacts — **the code**, **what the preview shows**, **what Word makes of
it** — with a screenshot at both render stages and machine-readable facts behind
both. The suite is simultaneously the test suite, the compatibility matrix, and
the roadmap: a feature we don't support yet is a case that is red on purpose.

---

## 1. What one case is

A case is a `.case.tsx` file. It is deliberately tiny — one feature, the
smallest document that exercises it — because the whole value is that when it
goes red you know what broke without bisecting a page.

```tsx
// conformance/cases/tables/header-row-repeat.case.tsx
import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.ts";

const lines = Array.from({ length: 60 }, (_, i) => i + 1);

export default defineCase({
  id: "tables/header-row-repeat",
  feature: "tables.headerRow",
  title: "A header row repeats on every page the table runs onto",
  word: "Table Properties -> Row -> Repeat as header row (w:tblHeader)",

  /** The claim under test. `unsupported` means the case is expected to fail. */
  claim: "supported",

  document: template(() => (
    <Document id="header-row-repeat" title="Header row repeat">
      <Table id="charges" columns={[{ width: 120 }, { width: 30, align: "right" }]}>
        <Row id="head" header>
          <Cell id="h1"><Paragraph id="h1p">Description</Paragraph></Cell>
          <Cell id="h2"><Paragraph id="h2p">Amount</Paragraph></Cell>
        </Row>
        {lines.map((n) => (
          <Row id={`r${n}`}>
            <Cell id={`r${n}c1`}><Paragraph id={`r${n}p1`}>{`Line ${n}`}</Paragraph></Cell>
            <Cell id={`r${n}c2`}><Paragraph id={`r${n}p2`}>10.00</Paragraph></Cell>
          </Row>
        ))}
      </Table>
    </Document>
  )),

  /** Anchors both geometry probes locate by. No design fixture — see §3. */
  regions: [
    { id: "head", anchor: "Description" },
    { id: "tail", anchor: "Line 60" },
  ],

  expect: {
    // Tier A — the file. Pure, fast, runs on any machine, lives in `npm test`.
    ooxml: (a, is) => {
      is.equal(a.table(0).row(0).tblHeader, true, "header row declares w:tblHeader");
      is.equal(a.table(0).widthsTwips.length, 2, "two declared columns");
    },
    // Tier C — Word's own reading. Windows + Word only.
    word: (c, is) => {
      is.equal(c.pages, 2, "the table spills onto a second page");
      is.equal(c.occurrences("Description").length, 2, "Word repeats the header on page 2");
    },
    // Tier B — the preview, held to Word rather than to an opinion of its own.
    preview: (b, is, c) => {
      is.within(b.region("head").y, c.region("head").y, "1mm", "header sits where Word puts it");
    },
  },
});
```

Every case, whatever it asserts, produces the same evidence:

```text
conformance/.out/tables/header-row-repeat/
  case.docx                 the file
  case.pdf                  Word's export
  preview-p1.png  -p2.png   docx-preview, headless Chrome screenshot
  word-p1.png     -p2.png   the PDF rastered at 96dpi
  measure-a.json            OOXML facts
  measure-b.json            preview geometry, px @96dpi, page-relative
  measure-c.json            Word geometry, points, page-relative
  measure-x.json            cross-engine pixel diff, preview vs Word
  result.json               one entry per assertion
```

The board reads that directory. **Code, preview, Word** — the triptych the whole
exercise is for.

---

## 2. Four probes, and what each one is allowed to prove

Lifted wholesale from `website/scripts/lib/VERIFY-CONTRACT.md`, which already
got these right. The probes stay; what changes is that they take a path and a
region list instead of knowing about an invoice.

| Probe | Reads | Proves | Runs where |
| --- | --- | --- | --- |
| **A — OOXML** | unzip `case.docx`, scan the XML | the packer emitted the element/attribute at all | anywhere |
| **B — preview** | `docx-preview` in headless Chrome, measured + screenshotted | what a person sees in the app we ship | anywhere with Chrome |
| **C — Word** | Word 16 over COM, page/x/y via `Information(3/5/6)` | the only ground truth there is | Windows + Word |
| **X — cross-engine** | pixel diff of preview PNG vs Word PDF raster, per region | that the preview is not lying | needs B and C |

The ordering matters and is the point of the whole design: **A can pass while
the document is wrong, B can pass while Word disagrees, only C settles it, and X
is the number that says whether the thing we ship people is honest.** `radiusPt`
was a property that did nothing in the file; a preview patched with CSS Word
cannot reproduce is the same failure wearing a nicer coat.

**Tiering, so the suite is usable off the Windows box:**

- Tier A assertions run inside `npm test` at the repo root, as ordinary
  `node --test` cases. No Chrome, no Word, no network. This is CI.
- Tier B adds Chrome. Cheap; runs on any dev machine.
- Tiers C and X need Word. They run on demand (`npm run conf -- --word`) and in
  the pre-release gate. A case whose C tier could not run is `SKIPPED`, never
  silently green.

---

## 3. What conformance cases do *not* need

The invoice harness carries a design fixture (`design-regions.json`, two
reference PNGs) because there was an external design to match. A conformance
case has no design. Its ground truth is Word itself, which means:

- **No fixture files.** Regions are declared by anchor text in the case, and
  probes B and C locate them the same way they already do.
- **No absolute geometry thresholds.** The assertions are *relative*: preview
  within 1mm of Word, this row taller than that one, this fill equals that hex.
  Absolute numbers only where the feature is an absolute number (a 30mm column
  is 30mm).
- **Baseline PNGs are evidence, not gates.** Word version and font version move
  under us; a pixel baseline that gates the build is a build that breaks when
  Office updates. Approved baselines (`--approve`) are stored per case and
  diffed at **warn** level, so a human sees "this looks different now" without
  the suite going red for a reason nobody can act on. The gates are the measured
  facts.

---

## 4. Red on purpose: how a case becomes an objective

`claim` is the mechanism that turns this suite into a roadmap.

| `claim` | Meaning | All assertions pass | Any fails |
| --- | --- | --- | --- |
| `supported` | we say this works | `PASS` | `FAIL` — a regression, break the build |
| `partial` | works in the file, not in the preview (or vice versa); the case names which tier | tier-by-tier | expected tiers red are `KNOWN`, others `FAIL` |
| `unsupported` | a gap we have written down | **`STALE` — promote it** | `KNOWN` — the roadmap |

Two properties fall out of this, and both are the reason for doing it this way:

1. **A gap you can't measure isn't on the roadmap.** To add an objective you
   write the case that fails. "Word conformance for numbered lists" becomes
   `lists/ordered-numbering` with concrete assertions about `w:numPr` and about
   what Word prints. It is now a small, checkable problem with an id.
2. **The suite cannot be made green by deletion.** A case that starts passing is
   reported `STALE`, not quietly ignored; the fix is to promote its claim, which
   is a diff a reviewer sees. Carry over the invoice harness's I2 discipline:
   hash the assertion source per run, and if a case goes `FAIL -> PASS` in the
   same run its hash changed, mark it `SUSPECT`.

The loop protocol from `invoice-v2-plan.md` §5 transfers unchanged: no objective
may be edited, weakened, deleted or skipped to make a run green; after three
failed attempts at the same mechanism, record the attempt and move to the next.

---

## 5. The compatibility inventory

This is the breakdown to work against. Current state is from the source as it
stands today (`src/render/docx_document.ts`, `src/domain/types.ts`); the seed
case is the first case to write for that row.

### Page and section

| Feature | State | Evidence | Seed case |
| --- | --- | --- | --- |
| Page size A4 / Letter | yes | `DocumentPageStyle.size` | `page/size-a4`, `page/size-letter` |
| Orientation | yes | `DocumentPageStyle.orientation` | `page/orientation-landscape` |
| Margins (mm, four sides) | yes | `DocumentPageMargins` | `page/margins` |
| Multiple sections | no — one `sectPr` for the whole document | `createDocxDocument` builds one section | `page/second-section` |
| Text columns | no | no `columns` in the model | `page/two-columns` |
| Page borders / background | no | — | `page/border` |
| Different first page | yes | `firstHeader` / `firstFooter`, `titlePg` | `furniture/first-page` |
| Even/odd furniture | **yes — added** | `evenHeader` / `evenFooter`, `w:evenAndOddHeaders` | `furniture/even-odd` |

### Running furniture — **done**

| Feature | State | Case | Note |
| --- | --- | --- | --- |
| Header and footer on every page | yes | `furniture/running` | A·B·C·X green |
| **Distance from the paper's edge** | **yes — added** | `furniture/distance` | was a silent 12.5mm chosen by the packing library |
| Page number counted by the renderer | **yes** | `furniture/page-numbers` | matches Word's count exactly |
| Different first page | **yes — both** | `furniture/first-page` | the running strip is asked of docx-preview, not built |
| Even / odd (recto and verso) | **yes — added** | `furniture/even-odd` | the preview alternates the strips too |

### The preview's own layout

| Feature | State | Case | Note |
| --- | --- | --- | --- |
| **Paginates by content** | **yes — added** | `preview/content-pagination` | `paginateDocxPreview`; matches Word exactly on ordinary paragraphs |
| Splits a paragraph across a break | no | — | breaks between blocks only, so a paragraph moves whole |
| **Splits a table between its rows** | **yes — added** | `preview/table-pagination` | `splitTable`; three pages where Word takes three, the seam within one row |
| **Holds a page to the size of the paper** | **yes — was silently false** | `preview/table-pagination` | a block that would not fit used to make its page taller: a first sheet 593mm deep above a second of ordinary A4 |
| **Repeats a heading row onto a new page** | **yes — added** | `tables/header-row-repeat` | `readPackedTables` reads `w:tblHeader`, settle moves those rows into a `<thead>`, and the split copies it |

### Paragraphs — **done**

The suite exists and this area is through it. Every row below is a case under
`conformance/cases/text/`, measured on all four tiers against Word 16.

| Feature | State | Case | Note |
| --- | --- | --- | --- |
| Body font, size, colour | yes | `text/typography` | A·B·C green |
| Space after | yes | `text/spacing-after` | A·B·C·X green |
| Shading, borders, padding | yes | `text/block-box` | X green after two preview fixes |
| Uppercase, letter spacing | yes | `text/caps-tracking` | X green after the tracking fix |
| **Alignment** (left/centre/right/justify) | **yes — added** | `text/align` | was: only inside a table cell |
| **Indentation** (left/right/first-line/hanging) | **yes — added** | `text/indent` | was: only the negative bleed indent |
| **Space before** | **yes — added** | `text/spacing-before` | was: headings only |
| **Keep-with-next / keep-lines** | **yes — added** | `text/keeps` | Word confirms the pair is not split |
| **Tab stops** | **yes — file and preview** | `text/tab-stops` | `applyTabStops` places them where there is a layout |
| Line height | yes, ±1.6mm at 2.0 leading | `text/line-height` | `partial` — measured, see the case |
| **Preview paginates by content** | **yes — added** | `preview/content-pagination` | matches Word exactly on ordinary paragraphs |

### Inline runs — the largest gap

`ParagraphNode.text` is **one string**. There is no run model, so nothing below
the paragraph can vary. Everything in this block is a single design decision
(introduce inline runs) with many cases hanging off it.

| Feature | State | Seed case |
| --- | --- | --- |
| Bold / italic / underline / strike in a sentence | no | `runs/bold-span` |
| A coloured or differently-sized span | no | `runs/coloured-span` |
| Superscript / subscript | no | `runs/superscript` |
| Hyperlink (internal and external) | no | `runs/external-hyperlink` |
| Line break within a paragraph | no | `runs/soft-break` |
| Non-breaking space, symbol runs | no | `runs/nbsp` |
| Inline image among words | yes | `runs/inline-image` |

### Headings, lists, references

| Feature | State | Evidence | Seed case |
| --- | --- | --- | --- |
| Heading 1 (from `Section`) | yes | `HeadingLevel.HEADING_1` | `headings/section-heading` |
| Heading 2–6 / nesting | no — every section is H1 | `renderNode` section branch | `headings/nested-levels` |
| Named Word styles | no | — | `headings/named-style` |
| Bulleted list | no — no numbering part emitted | — | `lists/bulleted` |
| Numbered list, nested | no | — | `lists/ordered-numbering` |
| **Table of contents** | **no — the node exists and packs as a plain line of text** | `renderNode` fallback, `src/render/docx_document.ts:437` | `fields/toc` |
| Bookmarks, cross-references | no | — | `fields/bookmark-crossref` |
| Footnotes / endnotes | no | — | `fields/footnote` |

### Tables — **done**

The suite has been through this area on all four tiers against Word 16. Every
row below is a case under `conformance/cases/tables/`. Three of them were red
when they were written and are green now: writing the case found the bug, and
the bug was one or two lines each.

| Feature | State | Case | Note |
| --- | --- | --- | --- |
| Column widths (mm and `auto`) | **yes — was silently wrong** | `tables/column-widths` | A·B·C·X green since `w:tblLayout`, see below |
| Column / cell alignment | **yes — added** | `tables/cell-align` | A·B·C·X green |
| Header row repeat | **yes — all four tiers** | `tables/header-row-repeat` | the preview repeats it too now, see below |
| Horizontal span | **yes — added** | `tables/colspan` | A·B·C·X green; Word reads the table as non-uniform |
| **Vertical merge** | no — no `rowSpan` in the model | `tables/rowspan` | nothing writes `w:vMerge`; the label is repeated or left blank |
| Cell fill, per-side borders, padding | **yes — was spaced twice** | `tables/cell-borders` | A·B·C·X green since the `w:space` fix, see below |
| Vertical alignment in a cell | **yes — added** | `tables/cell-valign` | A·B·C·X green |
| Bleeding table (negative `tblInd`) | yes — the preview needed a fix | `tables/bleed` | A·B·C·X green; `applyTableIndents` holds |
| Fixed row height | no | `tables/row-height` | no height on a row; `heightPt` sets leading, which is not the same thing |
| Nested table | **yes — measured** | `tables/nested` | was untested; A·B·C·X green |
| Banded rows | **yes — added** | `tables/banded-rows` | per-cell `w:shd` counted by the renderer, never a `w:tblStyle` |
| Table alignment and float | no | `tables/float` | no `w:jc` on a table, no `w:tblpPr` |

Three bugs came out of writing these, in the order they cost most. All three
are fixed, and the case that found each one is what keeps it fixed.

1. **Word reworked the column widths.** The packer emitted a correct
   `w:tblGrid` and no `w:tblLayout`, and a table without one is autofit to its
   contents — so 60mm / 40mm / 70mm printed as 67.5mm / 50.9mm / 51.6mm. The
   total was right, which is why nobody had noticed: the table still filled the
   column and only the boundaries inside it were wrong. The preview honoured
   the grid, so the two engines disagreed by up to 18mm about where a column
   starts. `renderTable` now declares `TableLayoutType.FIXED`, and so does the
   card a picture stands in. `tables/column-widths` holds the numbers.
2. **A cell's padding was written twice.** `blockBorders` set `w:space` from
   the block's padding, which is right for `w:pBdr` — where `w:space` is the
   gap between the border and the text and the only room a paragraph can have
   — and wrong for `w:tcBorders`, where the room inside is `w:tcMar`'s job.
   Word added the two and drew a padded panel 12pt lower than the preview did,
   and everything below it 12pt lower again. `blockBorders` now takes which
   element it is drawing for. `tables/cell-borders` holds it.
3. **The preview could not break a table.** Covered in the preview section
   below, along with the second failure hiding behind it.


### Images and graphics

| Feature | State | Evidence | Seed case |
| --- | --- | --- | --- |
| Inline picture, sized in points | yes | `imageRunOf`, `ptToPx` | `images/inline-sized` |
| `data:` URI source | yes | `src/render/image_source.ts` | `images/data-uri` |
| SVG with raster fallback | yes | `fallbackPath` | `images/svg-fallback` |
| Alt text | yes | `ImageNode.alt` | `images/alt-text` |
| Card / variant box around a picture | yes | `imageCard` | `images/card-variant` |
| Floating / text-wrapped image | no | — | `images/float-wrap` |
| Crop, rotate | no | — | `images/crop` |
| **Native chart** | **no — `GraphNode` packs as `[bar graph: ...]` text** | `renderNode` graph branch | `charts/bar-native` |

### Fields, furniture, breaks

| Feature | State | Evidence | Seed case |
| --- | --- | --- | --- |
| PAGE / NUMPAGES | yes — real fields | `pageNumberRuns` | `fields/page-of-total` |
| Preview shows the page number | partial — patched in after render | `fillPageFields` | `fields/page-of-total` (tier X) |
| DATE, DOCPROPERTY | no | — | `fields/date` |
| Header / footer | yes | `header`, `footer` | `furniture/header-footer` |
| Page break | yes — via a style, to avoid a stray line | `BREAK_STYLE_ID` | `breaks/page-break` |
| Column / section break | no | — | `breaks/section-break` |
| Core properties (author, subject, keywords) | title only | `createDocxDocument` | `meta/core-properties` |

**29 features work, 4 are partial, 29 are absent.** That count is the thing
this plan exists to move, and the board is where it moves in public.

It moved twice through the tables area, and the first move was downward, which
is the suite working. Three features written down as working turned out to work
in the file and not in Word — declared column widths, a padded cell, a header
row on page two — and one written down as untested turned out to work. Then
each of the three was one or two lines to fix, because a case with a failing
assertion and a number beside it is a small problem. A count that only ever
improves is a count nobody is measuring.

---

## 6. The runner

```sh
npm run conf                              # every case, every tier that can run here
npm run conf -- --case tables/rowspan     # one case
npm run conf -- --area tables             # one area
npm run conf -- --tier a                  # OOXML only: no Chrome, no Word
npm run conf -- --word                    # include Word + cross-engine
npm run conf -- --approve                 # write the PNG baselines for what just ran
npm run conf -- --board                   # emit conformance/report/index.html
```

Output mirrors the invoice runner: one line per assertion, then per-case status,
then `BOARD: n pass / n fail / n known / n stale / n skipped`. Exit non-zero when
anything is `FAIL`, `STALE` or `SUSPECT`. `KNOWN` does not fail the build —
otherwise the roadmap is a broken build.

**The board** is one static HTML page, one row per feature from §5, expanding to
the triptych: source on the left, preview screenshot in the middle, Word raster
on the right, with the tier chips and each failing assertion's measured-vs-
expected underneath. Published as an artifact, so a compatibility question gets a
link instead of a discussion.

---

## 7. Order of work

1. **Generalise the probes.** `website/scripts/lib/{probe-ooxml, probe-preview,
   probe-word.mjs, probe-word.ps1, probe-visual}.mjs` move to `conformance/lib/`,
   with the invoice's regions and thresholds replaced by parameters.
   `docx-page.mjs`'s jsdom render moves too; `PAGE_ONLY_STYLE` keeps its I1
   fixture hash. `verify-invoice.mjs` becomes a thin caller of the generic
   probes so the invoice board keeps working — that is the test that the lift
   was clean.
2. **`conformance/` as its own package**, `docxcelerate: file:..` the way
   `website/` does, carrying jsdom / pdfjs-dist / pixelmatch / pngjs /
   docx-preview as its own devDependencies. The root package keeps its single
   runtime dependency.
3. **`defineCase` + the runner + tier A into `npm test`.** Ten green cases for
   things we already claim: `page/size-a4`, `text/spacing-after`,
   `tables/column-widths`, `tables/header-row-repeat`, `tables/cell-valign`,
   `images/inline-sized`, `fields/page-of-total`, `furniture/first-page-differs`,
   `breaks/page-break`, `tables/bleed`. **Expect one or two of these to be red.**
   That is the first thing the suite buys us.
4. **Tiers B, C and X, and the board.** At this point every case has its
   screenshots and the cross-engine number is real rather than informational.
5. **Write the gaps as red cases**, in the order they cost us most:
   `text/align-centre` (a body paragraph cannot be centred — the cheapest and
   most surprising), `runs/bold-span` and the rest of the inline-run block,
   `lists/ordered-numbering`, `fields/toc`, `tables/rowspan`,
   `headings/nested-levels`.
6. **Then fix them, one red case at a time.** Each is now a small problem with an
   id, a failing assertion, and a picture of what Word does instead.

Steps 1 to 4 are done, and steps 5 and 6 are through the paragraph and table
areas. What the tables pass turned up has been fixed, in the same pass:

- **`w:tblLayout` on every table.** Word autofits a grid it was not told to
  keep, so every declared column width was a suggestion it ignored. Held by
  `tables/column-widths`.
- **No `w:space` on a cell border.** `blockBorders` is written for `w:pBdr`,
  where the space is the gap to the text; on `w:tcBorders` it was a second
  helping of the padding, and Word served both. Held by `tables/cell-borders`.
- **A table broken between its rows, on a page the size of the paper.** The
  paginator let a block it could not split make its page taller instead, so a
  three-page table was drawn as one sheet 593mm deep. `splitTable` breaks it at
  the last row that fits and carries the heading over. Held by
  `preview/table-pagination` and `tables/header-row-repeat`.

What is left in this area is the two features nothing can express yet — a
vertical merge (`tables/rowspan`) and a row with a height of its own
(`tables/row-height`) — both of which need a field on the model rather than a
line in the renderer. Table alignment (`tables/float`) is small and behaves;
floating a table is a question rather than a task, because `w:tblpPr` takes a
table out of the flow and makes the paragraphs wrap around it, and a preview
reproducing that is a preview doing its own line breaking.

## 8. What the suite must never do

Same rule the invoice harness runs under: the suite measures the framework, it
does not improve it. A conformance run touches nothing under `src/`,
`registry/`, `skills/` or `website/src/`. A green board is evidence, not proof —
the last step before a release is still to open a `.docx` in Word and look at it.
