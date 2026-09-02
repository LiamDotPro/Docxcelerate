# Invoice v2 — closing the gap to the design, and proving it closed

Design: `Invoice v2.dc.html` in the Docxcelerate design project
(`a093f680-868b-4b80-af4a-cfc4e02953ce`).
Implementation: `website/src/demo/documents/invoice/`.

**The target is 100%.** Not "close enough for a demo" — every region of the
packed `.docx` placed where the design puts it, filled and set the way the
design sets it, in Word *and* in the preview. Nothing in the design is dropped
as unreachable. Where Word needs a different mechanism than CSS to get the same
result, we build that mechanism.

This document has three parts: what "done" means and how it is measured (§1),
the fifteen framework gaps with an objective each (§3), and the document work
(§4). Every objective is written so that it passes or it doesn't — there is no
partial credit, and §5 is the rule for what to do when one fails.

---

## 1. What "done" means

**Done is:** for each of the nineteen named regions below, the packed file puts
it within **±1mm** of where the design puts it, at the same size, with the same
fill, stroke, face, size and tracking; Word paginates the document as **exactly
two pages**; and docx-preview shows the same thing.

The regions, which are the units every geometric objective is written against:

| Page | Regions |
| --- | --- |
| 1 | `letterhead`, `rule`, `band`, `parties`, `summary-label`, `summary`, `charges-head`, `charges-body`, `charges-rule`, `totals-panel`, `closer`, `footer` |
| 2 | `p2-letterhead`, `p2-rule`, `bank-grid`, `reference-panel`, `terms`, `scan-card`, `footer-2` |

### 1.1 The preview is not the file, and not Word

One premise needs correcting before the harness is built, because it changes
what the harness has to do.

**We do not render with the same package we write with.** `docx` (v9.6) writes
the file; `docx-preview` (v0.3.7) reads it back. Different packages, different
authors, different jobs. docx-preview is a good reader, but it is not Word:

- It does not evaluate fields. `PAGE` and `NUMPAGES` come out empty — which is
  why our footer currently shows a bare `/`.
- It does not paginate the way Word does. It honours explicit breaks and
  `min-height`, and lets a section grow past A4 rather than spilling. That is
  how page one is 1553px tall in the preview instead of becoming three pages.
- It resolves fonts through the browser. Word resolves them through Office.
  **These disagree on this machine** — see §1.4.

So agreement with the preview is necessary and not sufficient. Every geometric
objective is asserted **three times**, and only counts when all three agree.

### 1.2 The three probes

**Probe A — the file.** Unzip the `.docx` and assert on `word/document.xml`,
`styles.xml`, `header*.xml`, `footer*.xml`. This is ground truth for what we
wrote. `tests/docx.ts` already does the unzipping (`documentXml`, `partXml`,
`partNames`), and `tests/word_output.test.tsx` is already written in exactly
this idiom — new assertions go alongside, in
`tests/invoice_fidelity.test.tsx`. Deterministic, fast, runs in CI.

**Probe B — the preview.** Bake the page with the existing
`renderDocxPage()`, then measure it in a real layout engine: headless Chrome
loading a harness that reads `getBoundingClientRect()` for each region and
writes the numbers into the DOM, read back with `--dump-dom`. (This is how the
1553px measurement in §2 was taken; the technique works.) Catches anything that
is right in the XML and wrong on the page.

**Probe C — Word.** Word is installed here (16.0.20326) and drives from
PowerShell COM. This is the probe that makes "compiled Word document" real:

```powershell
$w = New-Object -ComObject Word.Application
$d = $w.Documents.Open($path, [ref]$false, [ref]$true)   # read-only
$d.ComputeStatistics(2)                                   # wdStatisticPages
$r.Information(3)                                         # wdActiveEndPageNumber
$r.Information(5)                                         # wdHorizontalPositionRelativeToPage, pt
$r.Information(6)                                         # wdVerticalPositionRelativeToPage, pt
$d.Tables(4).Cell(2,1).Shading.BackgroundPatternColor
$d.Sections(1).Headers(2).Range.Text                      # wdHeaderFooterFirstPage
$d.ExportAsFixedFormat($pdf, 17)                          # wdExportFormatPDF
```

`Information(5)` and `Information(6)` give the position of any range **in
points from the page corner**, which is precisely what a ±1mm objective needs.
Page count, per-page placement, cell shading, run fonts and shape counts all
come straight from Word's own model. The exported PDF feeds the visual probe.

**The visual probe** rasterises the PDF with `pdfjs-dist` in headless Chrome at
150dpi (A4 → 1240×1754), renders the design canvas at the same size, and
compares per region with `pixelmatch`. New devDependencies for `website/`:
`pdfjs-dist`, `pixelmatch`, `pngjs`. All pure JS, no native builds.

### 1.3 The harness

```text
website/scripts/verify-invoice.mjs        # the runner; exit 1 if anything is not PASS
website/scripts/lib/objectives.mjs        # every objective and every threshold, in one file
website/scripts/lib/probe-ooxml.mjs       # Probe A
website/scripts/lib/probe-preview.mjs     # Probe B
website/scripts/lib/probe-word.ps1        # Probe C, emits JSON on stdout
website/scripts/lib/probe-visual.mjs      # PDF -> PNG -> per-region diff
website/.verify/report.json               # the record: per objective, per probe, per attempt
```

`npm run verify:invoice` in `website/`. The report is the loop's memory — see §5.

### 1.4 P0 — the font precondition (do this first)

**Word can resolve Aptos on this machine. Chrome cannot.** Checked, not
assumed: Word's `FontNames` lists Aptos, Aptos Display, Aptos Mono; Chrome
measures `Aptos` at exactly its monospace fallback width, the same as a
nonsense face. No `Aptos*` file exists in `C:\Windows\Fonts`, the per-user font
store or Office's VFS — Word is listing it as a cloud font.

That is not a detail. Different faces mean different advance widths, different
line breaks, different block heights — so the preview and Word would disagree
about whether a page fits, and any pixel comparison of text would be noise.

**P0 objective.** The same face resolves in Word *and* in Chrome, proven by
measurement in both, before any pixel-level objective is allowed to report PASS.

Two ways to satisfy it, in order of preference:

1. Install Aptos properly on the verification machine so both engines resolve
   it, and re-run the two measurements.
2. Failing that, pin `invoice-style.ts` and the design file to a face installed
   here — Segoe UI is present to both — and record the substitution in the
   report. This is a change to the design as well as the document, so it needs
   saying out loud rather than doing quietly.

Until P0 is PASS, the visual objectives report **BLOCKED**, never PASS.
Structural and geometric objectives still run: they are asserted from twips in
the file and from Word's own measurements, neither of which depends on the
browser's font list.

### 1.5 Invariants — the things that must not be "fixed" the easy way

Four ways an objective could be made to pass without the document getting any
better. The runner checks all four, and reports `SUSPECT` rather than `PASS`.

- **I1 — no CSS patching.** `PAGE_ONLY_STYLE` in `docx-page.mjs` must stay
  byte-identical to the checked-in fixture. If the file is wrong, fix the
  packer, not the stylesheet. (The one sanctioned change to `docx-page.mjs` is
  filling `PAGE`/`NUMPAGES` from the laid-out section count — see C6 in §4.3.
  That is JS, not CSS, and it writes the number Word itself prints.)
- **I2 — no threshold fiddling.** The runner records a hash of
  `objectives.mjs`. An objective that goes red → green in the same run as a
  threshold change is `SUSPECT`.
- **I3 — one run.** Every objective must be green in the *same* run. No
  collecting greens across runs.
- **I4 — no shrinking the data.** `preview-data.ts` keeps the longest name and
  largest figure. The runner asserts the longest description is ≥ 52 characters
  and the line count is ≥ 7, so "make it fit" cannot mean "print less".

### 1.6 What the preview is allowed to stand in for

**Anything the preview can honestly compute, it computes.** A placeholder is for
what genuinely cannot exist yet, and the list of those is shorter than the
current document assumes.

The rule the framework already has: a deriver stands in for itself in a preview
**only when it declared a `placeholder`** — `standsInForPreview()` is exactly
`this.#placeholders.has(name)` (`src/runtime/derivers.ts:131`). Declare one for a
deriver that renders, reads or fetches; leave it off and it runs, and the preview
shows the real value.

After the work below, one thing in this document is a placeholder in preview:

| | Preview shows | Why |
| --- | --- | --- |
| Engagement summary | The placeholder prose | Genuinely needs a model. Previews never call one — nothing leaves the machine, and the same build gives the same page every time. The placeholder text is word-for-word what the design canvas shows, so the page still matches. |
| Scan-to-pay code | **A real, scannable QR** | It is a deterministic encoding of the account, amount and reference. A deriver computes it, and a deriver with no declared placeholder runs. See D14. |
| Page numbers | **Real numbers** | Filled from the laid-out section count. See C6. |
| Totals, VAT, amounts | **Real figures** | Already real, and about to become derivers so they are also real per document. See D13. |

That table is the reason two items below change shape: the QR stops being a
prompt (§4.1, D14), and the image placeholder stops being a line of text
(§3, F13).

---

## 2. Where we start

Page one does not fit on a page. The rendered first section is **1553px against
A4's 1123px** at 96dpi — 430px over, about 114mm. In the preview it grows; in
Word it spills, and the invoice prints on three pages.

| Block on page one | Packed | Design |
| --- | ---: | ---: |
| Letterhead table | 103px | as drawn |
| "Invoice details" heading | 25px | 0px |
| Meta band | 104px | ~57px |
| "Parties" heading | 25px | 0px |
| Parties table | 153px | ~89px |
| "Engagement summary" heading | 25px | 25px |
| Summary paragraph | 73px | 73px |
| "Charges" heading | 25px | 0px |
| Charges table | 488px | ~334px |
| Totals table | 138px | ~98px |
| Line left by `<PageBreak>` | 23px | 0px |
| **Body total** | **1182px** | **~676px** |

Budget: `1123 − 120 margins − 69 running header − 114 running footer = 820px`.

---

## 3. The framework gaps

F1–F13 are as originally filed. **F14 and F15 are new**: they were on the old
"Word cannot do this" list, and with the target at 100% they stop being excuses
and become work. The other four items on that list turned out to need no
framework change at all and have moved to §4.

Each gap carries an **Objective** (what must be true), the **Probes** that
decide it, and **Mechanisms** in the order to try them. A gap is done when every
probe is green in one run.

---

### F1 — A section that doesn't print its heading

`renderNode` emits a `HEADING_1` for every section, unconditionally
(`src/render/docx_document.ts:120`). `showTitle` exists but is document-wide.
The invoice needs sections for what sections are for — addressing a node by id,
grouping for the engine, a diff that lines up — while the design draws the
heading itself as a 7pt label inside a table, or not at all. Three unwanted
headings cost 75px of page one.

Note which headings the design *keeps*: "Engagement summary", "Pay by bank
transfer", "Terms & notes", "Scan to pay" are all in the design as 7pt tracked
labels, which is what `sectionHeading` already produces. Only "Invoice
details", "Parties" and "Charges" go.

**Mechanism.** `<Section title="Charges" showTitle={false}>`. `title` stays
required, so the section keeps its name for the model and the TOC; only the
printing is off, mirroring the document-level flag so there is one idea rather
than two.

**Objective F1.**

- **A** — `document.xml` contains no `w:pStyle w:val="Heading1"` paragraph whose
  text is `Invoice details`, `Parties` or `Charges`; and still contains one for
  each of the four the design keeps.
- **A** — the built `DocumentModel` still carries all eight sections with their
  titles intact. Suppressing the heading must not cost the name.
- **B** — no text node between the letterhead and the band; region `band` starts
  ≤ 3px below `rule`.
- **C** — Word's `Content.Find` for each of the three strings returns no match.
- **Fail action.** If `showTitle` reaches the renderer but the heading still
  prints, the bug is in `renderNode`, not the prop. Do not work around it by
  deleting the section.

---

### F2 — Different furniture on the first page

A document gets one header and one footer (`docx_document.ts:73–79`). The
design's page one has no running header — its letterhead *is* the top of the
page — and page two does. We print both, so page one names the sender twice and
loses 69px doing it.

Word has this natively: `w:titlePg` with a first-page header/footer. `docx`
exposes it as `properties.titlePage` plus `headers: { first, default }`.
docx-preview understands `titlePg` and `headerReference` (both appear in its
source), so this should show correctly in both — **verify, don't assume.**

**Mechanism.** `<Document firstHeader={…} firstFooter={…}>`, accepting `false`
to mean "nothing on the first page".

**Objective F2.**

- **A** — `sectPr` contains `<w:titlePg/>`; a `headerReference w:type="first"`
  exists and its part holds no text; the `default` header part holds the sender
  and the reference.
- **B** — page 1's rendered `header` element has height ≤ 2px; page 2's is
  between 40px and 80px.
- **C** — `Sections(1).Headers(2).Range.Text` (first page) is empty;
  `Headers(1).Range.Text` (primary) contains `INV-2026-0142`.
- **Fail action.** If Probe B fails while A and C pass, docx-preview is ignoring
  `titlePg` — record it in the report as a preview-reader limitation with the
  evidence, and raise it. It is not licence to skip the objective, and it is
  never fixed in `PAGE_ONLY_STYLE` (I1).

---

### F3 — A table that bleeds

`blockIndent` is applied in the paragraph branch only. A block style's
`bleed: true` works on the navy rule, which is a paragraph, and does nothing on
the footer bar, which is a table — so the design's edge-to-edge footer stops at
the margins.

**Mechanism.** Honour `bleed` in `renderTable`: negative `tblInd`, with the
margins added back into the column widths, exactly as the paragraph does it.

**Objective F3.**

- **A** — the footer table carries `w:tblInd` of `-907` twips, and its column
  widths total `11906` twips (210mm).
- **B** — the footer table's left edge is at x = 0 of the section, and its width
  equals the section width, both ±1px.
- **C** — `Information(5)` for the footer table's range is `0pt ± 1pt`, and its
  width is `595.3pt ± 1pt`.
- **Mechanisms in order.** (1) negative `tblInd`; (2) if Word clips a table
  indented past the margin, widen the section's text column and inset every
  other block instead. Do not settle for a bar that stops at the margin.

---

### F4 — An image inline with text

`cellContent` gives every child its own paragraph, so the footer's mark and
"Generated with Docxcelerate" stack, and a one-line footer bar becomes three
lines — 114px where the design spends about 50px. Word has no trouble with an
inline drawing in a run; we never build one.

**Mechanism.** Let `<Paragraph>` take an `<Image>` among its children and pack
it as an `ImageRun` in the same `w:p` as the text runs.

**Objective F4.**

- **A** — exactly one `w:p` in `footer*.xml` contains both a `w:drawing` and the
  text `Generated with Docxcelerate`.
- **B** — the rendered `footer` element is ≤ 53px tall.
- **C** — footer height ≤ 14mm, measured as `Information(6)` of the footer's
  first range minus the page height less the bottom margin; and the footer range
  reports one line.
- **Fail action.** The image must stay a picture. A footer that hits the height
  by dropping the mark fails.

---

### F5 — Non-paragraph cell children obey the cell

`cellContent` special-cases images and paragraphs and hands everything else to
`renderNode`, which knows nothing about the column. A `<PageNumber>` in a
right-aligned cell prints left and takes the body's 6pt spacing-after inside a
cell where every other child is set to zero. That is the rest of why the footer
is three lines deep.

**Mechanism.** Apply the resolved alignment and the cell's zero spacing to every
child kind, not to two of them.

**Objective F5.**

- **A** — the `w:p` holding the `PAGE` field carries `<w:jc w:val="right"/>` and
  `<w:spacing w:after="0" .../>`.
- **B** — the page-number element's right edge is within 1px of the footer
  table's right edge.
- **C** — the page-number range's `ParagraphFormat.Alignment` is `2` (right),
  and `SpaceAfter` is `0`.
- **C** — with fields updated (`$d.Fields.Update()`), the footer on page 1 reads
  `1 / 2` and on page 2 reads `2 / 2`.

---

### F6 — A block can name a face

`DocumentBlockStyle` has no `font`. Qty, rate, amount, sort code, IBAN and the
reference are all Consolas in the design, for the usual reason: proportional
digits in a money column don't line up. Today a theme can say a figure is small
and grey but not that it is tabular.

**Mechanism.** `font?: string` on `DocumentBlockStyle`, passed through
`blockRun` to the run. Word substitutes a missing face the way it already does
for the body font, so this carries the same guarantee as everything else in
`blocks`.

**Objective F6.**

- **A** — every `w:r` in the qty, rate and amount columns carries
  `<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>`; body cells do not.
- **B** — computed `font-family` of those cells resolves to Consolas (which
  Chrome *does* have here, unlike Aptos).
- **C** — `.Cell(2,4).Range.Font.Name` is `Consolas`.
- **C** — the right edge of every amount is within **0.35mm** of every other.
  This is the objective that actually matters; the font name is only how it is
  achieved.

---

### F7 — A block can set its own leading

Every paragraph, in a cell or out, takes `typography.bodyLineHeight`. The design
sets the body at 1.5, table rows near 1.2, and the 8pt note under a description
tighter still. One leading for the whole document is what makes our charge rows
66px against the design's 44px — 150px across seven lines — and what makes the
address blocks airier than drawn.

**Mechanism.** `lineHeight?: number` on `DocumentBlockStyle`, overriding the
body's wherever a variant is in play. Word stores line spacing per paragraph.

**Objective F7.**

- **A** — charge-row paragraphs carry the theme's line value, not `348`
  (1.45 × 240); the muted note carries its own, tighter.
- **B** — every charge row is ≤ 46px tall; region `charges-body` ≤ 340px.
- **C** — the charges table's height is ≤ 90mm, from the difference of
  `Information(6)` at its first and last rows.
- **C** — no charge row wraps to a third line at any data length in
  `preview-data.ts` (I4 keeps that data long).

---

### F8 — Alternating row fill, from the theme

The obvious `.map((line, i) => <Row variant={i % 2 ? "zebra" : undefined}>)`
works in preview and **breaks on publish**: the map becomes a loop the engine
walks, and a variant is a static string, so every row gets whatever the build
decided once. Striping is a rendering decision in any case — the row does not
know it is odd.

**Mechanism.** If the theme defines `blocks.rowAlt`, the renderer fills every
other body row with it. No node changes, and it survives publishing because the
renderer counts rows as it draws them.

**Objective F8.**

- **A** — body rows 2, 4 and 6 of the charges table carry
  `w:shd w:fill="F7F8FD"`; rows 1, 3, 5 and 7 carry no `w:shd`. The header row
  keeps the accent fill.
- **A (publish)** — build the same template through the publish path and pack
  the artifact directly: the striping must still alternate. A variant baked at
  build time fails this.
- **A (length)** — build with 3 lines and with 30 lines: striping alternates in
  both, and the header row is still the only accent row.
- **C** — `.Cell(3,1).Shading.BackgroundPatternColor` is `RGB(247,248,253)` and
  `.Cell(2,1)` is `wdColorAutomatic`.

---

### F9 — Turning the row hairline off

`separatorBorder` draws a rule under every unfilled, non-header body row and
there is no way to say no. With F8 in place the invoice wants stripes *instead
of* rules, and a table with both is wearing belt and braces. The only way out
today is to give the row a border in the page colour, which is a lie written
into the file.

**Mechanism.** Treat `borderSides: []` on a block as "no borders", rather than
falling through to the default because no border colour was set.

**Objective F9.**

- **A** — no `w:tcBorders` in the charges table carries the rule colour
  `D9DDEB`, except the single closing rule under the last row.
- **A** — a block with `borderSides: []` and no `border` emits no `w:tcBorders`
  at all, asserted on a minimal fixture so the behaviour is pinned outside the
  invoice.
- **C** — `.Cell(3,1).Borders(wdBorderBottom).LineStyle` is
  `wdLineStyleNone`.

---

### F10 — Tracking on headings

`DocumentBlockStyle` has `letterSpacingEm`; `DocumentTextBlockStyle` (title and
section headings) does not. The design's labels are tracked 0.12em and the
wordmark 0.14em. `invoice-style.ts` already sets the title to 23pt uppercase to
match the design and then cannot open it up — so the one place the wordmark is
defined disagrees with the block that duplicates it.

**Mechanism.** `letterSpacingEm?: number` on `DocumentTextBlockStyle`, applied
in `createDocxStyles`.

**Objective F10.**

- **A** — `styles.xml`'s Heading1 run carries `<w:spacing w:val="18"/>`
  (0.12em × 7.5pt × 20).
- **C** — `Styles("Heading 1").Font.Spacing` is `0.9pt ± 0.05`.
- **C** — the rendered width of `ENGAGEMENT SUMMARY` is within **0.5mm** of the
  design's.

---

### F11 — Vertical alignment in a cell

Cells always sit top-aligned. The design centres the band's contents against the
status pill, and the footer's three items against each other. Word has
`w:vAlign` per cell, and docx-preview reads it.

**Mechanism.** `valign` on `<Cell>`, or on the block style if it reads better as
a theme decision — pick one and pin it in a test.

**Objective F11.**

- **A** — band and footer cells carry `<w:vAlign w:val="center"/>`.
- **B** — the vertical centre of the status pill is within 1px of the vertical
  centre of the `Issue date` value.
- **C** — `.Cell(1,4).VerticalAlignment` is `1` (centre).

---

### F12 — A page break that doesn't leave a line

`<PageBreak>` packs as a paragraph containing only a break, leaving an empty
line at the foot of the outgoing page — 23px of page one, and 23px we are over
budget by.

**Mechanism.** Attach the break to the first paragraph of the next node with
`pageBreakBefore`, rather than giving it a paragraph of its own.

**Objective F12.**

- **A** — no `w:p` whose only run is `<w:br w:type="page"/>`; the payment page's
  first paragraph carries `<w:pageBreakBefore/>`.
- **A** — the existing test "a page break is Word's own break, not a run of
  empty paragraphs" still passes.
- **B** — page 1 has no trailing empty paragraph; region `footer` is the last
  element.
- **C** — still exactly 2 pages, and the closer line's
  `Information(3)` is `1`.

---

### F13 — An image draws its variant, and its placeholder is a picture

Two things, one node.

**One.** `imageRunOf` never consults `blockOf`, so `<Image variant="card">` is
silently ignored — the scan-to-pay card has no border, no padding, no card.

**Two.** An image with no bytes yet packs as the text `[image: …]` on a line of
its own. That is wrong twice over: the page is laid out around a string and
re-laid-out around a 38mm square when the real one arrives, and a line of prose
is not what a missing picture looks like. **A placeholder for a picture should be
a picture** — a box at the declared size, with the label set inside it, so the
page has the right shape before the content exists and a reader can see what is
coming.

For this document the placeholder mostly stops mattering, because the QR becomes
a deriver that actually runs (D14). It still matters for the general case, and it
is what the `scan-card` region falls back to if a deriver is ever given a
placeholder of its own.

**Mechanism.** Draw the block's fill, border and padding around the image — a
single-cell table is the honest packing, since Word has no box around a run —
and give an unresolved image a frame of its declared `width` × `height` with the
label centred inside, drawn in the `muted` variant.

**Objective F13.**

- **A** — the scan-to-pay image sits in a single-cell table with border
  `D9DDEB` and `w:tcMar` of 12pt.
- **A** — an unresolved image emits a cell of exactly its declared `width` ×
  `height` containing the placeholder text. **No `w:t` anywhere in the file
  matches `/^\[image: /`** — that string is the current behaviour and its
  absence is the objective.
- **A (no reflow)** — build twice, once with a real data URI and once
  unresolved. The outer box's width and height are **identical** in both, and
  the following region's `y` differs by **0px**.
- **B** — region `scan-card` is 38mm ± 1mm square, centred in the right column,
  in both builds.
- **C** — the card table's width is `108pt ± 1pt` and height `108pt ± 2pt`.
- **Fail action.** A card that only appears once the image resolves fails, and
  so does a placeholder that is still a run of text at the declared size. The
  box has to be a box.

---

### F14 — A block can set its own measure *(new)*

The design sets the engagement summary to a 158mm measure inside a 178mm text
column, so the paragraph reads at a sensible line length rather than running the
full width. This was previously written off as unreachable; it isn't. Word
indents paragraphs natively, and 178 − 158 = a 20mm right indent.

`DocumentBlockStyle` can express a *negative* indent (`bleed`) and no positive
one, which is a strange asymmetry now that it is noticed.

**Mechanism.** `indentRightMm` / `indentLeftMm` on `DocumentBlockStyle`, sharing
the code path `bleed` already uses.

**Objective F14.**

- **A** — the summary paragraph carries `<w:ind w:right="1134"/>` (20mm).
- **B** — region `summary` is 158mm ± 1mm wide.
- **C** — `.ParagraphFormat.RightIndent` is `56.7pt ± 0.5`, and no line of the
  summary exceeds 158mm.

---

### F15 — Rounded blocks *(new)*

The design rounds seven things: the status pill at 999px, the charges header's
top corners at 3px, the totals panel, the reference panel and the scan-to-pay
card at 4px, and the inner total bar at 3px. (The two `F` mark tiles are also
rounded, but they are supplied SVGs — the radius is already in the asset, and
those cost nothing.)

This was the one item genuinely written off. It should not have been. Word draws
rounded rectangles natively; we simply have no node that emits one. Note the
history here: `radiusPt` once existed on a block style and did nothing in Word,
and was removed for exactly that reason (see the note at the head of
`tests/word_output.test.tsx`). Re-adding it as a property that *does* something
is the fix; re-adding it as one that doesn't is the mistake we already made.

**Mechanisms in order.** Each must be proven in Word *and* in docx-preview
before it is accepted:

1. **VML round-rect** — `w:pict` holding a `v:roundrect` with `arcsize`, behind
   or around the block. Word renders VML natively, and docx-preview has a
   `VmlElement` renderer that draws to SVG. `docx` exposes `Textbox` (VML) and
   `ImportedXmlComponent` for raw XML, so this is emittable today.
2. **DrawingML shape** — a `WpsShapeRun` with `prstGeom prst="roundRect"`.
   Native to Word; check whether docx-preview draws it.
3. **A packed picture** — generate the rounded ground as a PNG at pack time and
   place the text over it. Last resort: it works everywhere and costs
   selectable text under the pill, so only if 1 and 2 both fail their probes.

**Objective F15.**

- **A** — the file contains real rounded geometry for each of the seven regions
  — a `v:roundrect` with a non-zero `arcsize`, or `prstGeom prst="roundRect"`.
  A square block with a radius recorded only in the style fails.
- **B** — at the pill's four corners, the preview's rendered pixel is the page
  colour, not the badge fill, at 2px inside the bounding box diagonal; and the
  fill is present at the pill's centre.
- **C** — same corner test against the Word PDF raster.
- **C** — `.Shapes.Count + .InlineShapes.Count` accounts for every rounded
  region.
- **Anti-cheat.** The corner test runs against the *rendered file* in both
  engines. It cannot be satisfied from `PAGE_ONLY_STYLE` (I1), and a preview
  that rounds while Word squares fails on Probe C.
- **Fail action.** Work down the mechanism list, recording each attempt and the
  probe that rejected it in `report.json`. Dropping the radius from the design
  is not an outcome — if all three mechanisms fail with evidence, escalate.

---

## 4. Document, theme and preview work

### 4.1 No framework change needed

- **D1 — Totals.** Drop `header` from the total row — it is what turns "Total
  due" into tracked capitals via `headingRun` — and move `variant="totalRow"`
  onto the two cells carrying words. The spacer cell then draws nothing, and the
  navy bar sits over the right-hand 84mm instead of crossing the page. Give the
  subtotal and VAT cells `variant="panel"`.
- **D2 — Header rows that aren't navy bars.** `variant="label"` on the header
  rows of the band and the parties grid. The navy default only applies when no
  variant resolves, so naming one both removes the bar and picks up the 7pt
  tracked label the theme already defines.
- **D3 — The band as a band.** One row, not two: each cell holds a label
  paragraph over a value paragraph.
- **D4 — The reference panel.** A one-cell table with `variant="panel"` holding
  three paragraphs. Consecutive shaded paragraphs each draw their own box with a
  6pt gap; one cell draws one box.
- **D5 — Page two in two columns.** A two-cell table, `[auto, 62]`, with nested
  tables inside. `cellContent` already renders a nested table.
- **D6 — The page-one closer.** "Payment details, terms and a scan-to-pay code
  are on page 2." beside "Due 4 September 2026".
- **D7 — Two names for the sender.** Trading name in the letterhead, legal name
  in the From block, rather than "Ltd" twice.
- **D8 — Page two's letterhead.** The reduced mark and the "PAYMENT" wordmark,
  as body content at the top of the payment page.
- **D9 — A muted second paragraph** in the terms.
- **D10 — Solid ink for the footer.** The design's
  `rgba(255,255,255,0.85)` over `#1E2A66` resolves to `#D5D8E4`. Put that in the
  theme. Composited at build time this is not an approximation — it is the same
  colour, exactly.
- **D11 — Copy.** "Terms & notes", not "Terms and notes"; the design's
  ampersand is the design's.
- **D12 — The annotation chips do not print.** "AI · summary.engagement" beside
  the summary label, and "Placeholder · deriver: payment.qr" under the code.
  These are the design canvas labelling its own elements — notes *about* the
  document rather than content *in* it. The second one settles it: it says
  "Placeholder", and after D14 there is no placeholder there, so printing it
  would put a false statement on the invoice. **This is the one deliberate
  departure from the canvas**, and it is a departure from annotation, not from
  design — every mark that is part of the invoice is reproduced.

- **D13 — Totals through derivers, so the document can publish at all.**
  `Totals` and `Payment` both compute with `data.lines.reduce(...)` in a
  `useState` initializer. That works locally and **throws under the publish
  stand-in** — `reduce` and `reduceRight` are on the refused list
  (`src/template/publish.ts:435`), because the entries do not exist until a
  request does. As written, this document is preview-only.

  Move the subtotal, VAT and total into derivers (`deriver_module.ts:91` has
  almost exactly this example), leave the placeholders off so they run in
  preview, and refer to them as `{{derived.total}}`. Two things depend on this:
  F8's publish-path probe, which cannot run until the document publishes at all,
  and D14, which needs the total to encode.

  Objective: `buildDocument` through the publish path succeeds; `document.json`
  carries `{{derived.total}}` rather than a baked figure; `preview.json` carries
  the real figure; and the two agree once the deriver runs.

- **D14 — The scan-to-pay code is a deriver, not a prompt.** `ScanToPay`
  currently carries a `generalPrompt` asking an engine to *draw* a payment QR.
  That is the wrong instrument: a QR is a deterministic encoding of a string,
  not a picture a model should invent, and a model-drawn one would not scan.

  The design says so itself — the chip under the card reads
  "deriver: payment.qr". Registering `paymentQr` and leaving its `placeholder`
  off means it runs in the preview too (§1.6), so the preview shows a real,
  scannable code rather than a stand-in:

  ```tsx
  <Image
    id="scan-to-pay"
    variant="card"
    src="{{derived.qrSvg}}"
    fallbackSrc="{{derived.qrPng}}"
    alt={`Scan to pay invoice ${state.reference}`}
    width={108}
    height={108}
    derivers={[
      derive("paymentQr", { output: "qrSvg", inputs: [...payTo, literalValue("svg")] }),
      derive("paymentQr", { output: "qrPng", inputs: [...payTo, literalValue("png")] }),
    ]}
  />
  ```

  Two outputs because Word will not embed an SVG alone: the screen gets the
  vector, the `.docx` gets the raster, which is what `fallbackSrc` is for. A
  static image's `path` goes through `renderTemplate`
  (`src/runtime/resolver.ts:224`), so a `{{derived.…}}` token in `src` resolves —
  and `{{derived.…}}` is the one token form that survives publishing, because
  derivers are preserved rather than run.

  Note what this changes about the node: it stops being dynamic. It has a `src`
  and no prompts, which is the correct classification — the QR is computed, not
  composed.

  Needs a QR encoder. `qrcode` is pure JS with no native build; vendoring a
  minimal encoder is the alternative if a dependency in a demo's derivers is
  unwelcome. Encode a payment URI carrying the IBAN, the amount and the
  reference.

  Objective D14:

  - The preview's `scan-card` region contains an image, not text, and decoding
    it yields the payment URI with the right IBAN, amount and reference.
  - The packed `.docx` embeds a raster (`imageRunOf` takes the `png` branch), and
    Word reports one inline shape in the card.
  - Built through the publish path, `document.json` keeps the deriver invocation
    and the `{{derived.qrSvg}}` token rather than a baked data URI.
  - Changing the reference in `preview-data.ts` changes the decoded payload.

  **The design's own QR is a stand-in** — the canvas draws random modules around
  three finder patterns. So the `scan-card` region is compared on geometry, not
  pixels: the card, the label, the caption and the code's box must match to
  ±1mm, and the code itself must be *real*, which the canvas's is not. This is
  the one place the document is deliberately better than the drawing.

### 4.2 Preview harness

- **C6 — Page numbers.** docx-preview does not evaluate `PAGE`/`NUMPAGES`, so
  the footer shows a bare `/`. Fill both from the section index and the rendered
  section count in `docx-page.mjs`, after layout. This is the one sanctioned
  change under I1: it is JS, not CSS, and it writes the number Word itself
  prints. Objective: the baked page reads `1 / 2` and `2 / 2`, and Probe C
  agrees after `Fields.Update()`.

### 4.3 Out of scope, and why

The design canvas's page shadow, ruled workspace ground and 28px gap between
sheets are canvas furniture, not document content. The comparison crops to the
page, so they never enter a diff.

---

## 5. The loop

The objectives exist so that "done" is not a judgement call. The protocol:

1. Run `npm run verify:invoice`. It writes `.verify/report.json` and prints one
   line per objective: `PASS` / `FAIL` / `BLOCKED` / `SUSPECT`.
2. **Work is not finished while any objective is not PASS.** `BLOCKED` (a
   precondition such as P0 is unmet) and `SUSPECT` (an invariant tripped) both
   count as not-PASS.
3. On `FAIL`, fix and re-run. After three consecutive failures of the same
   objective with the same approach, move to the next mechanism listed for that
   gap and record the attempt — which mechanism, which probe rejected it, and
   the measured value — in `report.json`'s `attempts[]`.
4. **No objective may be edited, weakened, deleted or skipped to make a run
   green.** Thresholds are hashed (I2); the stylesheet is fixtured (I1); the
   preview data is length-checked (I4).
5. Escalate to a human only when every listed mechanism for a gap has failed
   with recorded evidence. "It doesn't seem possible" without a populated
   `attempts[]` is not an escalation.
6. The final run must be a single run in which every objective, every global
   gate and every invariant is green, with P0 satisfied.

### Global gates

| Gate | Assertion |
| --- | --- |
| G1 | Word reports **exactly 2 pages** |
| G2 | Preview section 1 height ≤ **1123px**, and ≥ 1000px (it has not collapsed) |
| G3 | All 19 regions within **±1mm** of the design, from Word's own measurements |
| G4 | Per-region pixel diff vs the design ≤ **3%** of pixels, none above 5% (requires P0). `scan-card` is exempt from the pixel test and checked on geometry only — the canvas's QR is a stand-in and ours is real (D14) |
| G5 | `npm test` green at the repo root; `npm run documents:check` green |
| G6 | Invariants I1–I4 clean |

---

## 6. Order of work

Build the harness first. Every objective below is measured by it, and an
objective you cannot measure is an opinion.

1. **P0 and the harness.** Settle the font question, then
   `verify-invoice.mjs` with all three probes wired and every objective
   registered as FAIL. A red board you trust beats a green one you don't.
2. **F1, F2, F12.** The cheapest three; they buy back ~190px of page one.
3. **F7, F6.** Leading and face — the largest remaining saving (~150px), and
   what makes the money columns line up.
4. **F4, F5, F3.** The footer bar: one line, edge to edge, page number correct.
5. **D1–D4, D9–D12.** Page-one document work, now that the page fits. Expect G1
   and G2 to go green here.
6. **D13, then D14.** Totals through derivers first — the document cannot
   publish until that is done, and D14 needs the total to encode. Getting these
   in before F8 means F8's publish-path probe can actually run.
7. **F13, F8, F9, F11, F10, F14.** The card, the stripes, the details.
8. **F15.** The rounded blocks, worked down the mechanism list. Left late
   because it is the one with genuine unknowns, and everything else should be
   green before its unknowns are the only thing moving.
9. **D5–D8.** Page two.
10. **The final run.** One run, everything green, P0 satisfied — then open the
   file in Word and look at it, because a green board is evidence and not proof.

```sh
npm run build                 # repo root
npm test                      # repo root — Probe A objectives live here
cd website && npm run demo    # -> website/public/demo/invoice.html
npm run verify:invoice        # the board
npm run documents:check
```
