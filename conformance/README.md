# Conformance

One small document per Word feature. Each is packed to a `.docx`, read back by
docx-preview, opened in Word, and asserted against all three — plus a fourth
tier that holds the preview and Word against each other.

```sh
cd conformance
npm install

node runner.mjs                          # every case, every tier that can run here
node runner.mjs --tier a                 # the file only: no Chrome, no Word
node runner.mjs --word                   # include the Word and parity tiers
node runner.mjs --case text/align        # one case
node runner.mjs --area text              # one area
node runner.mjs --word --no-screenshots  # faster, when the pictures are not the point
node -e "import('./lib/board.mjs').then(m => m.writeBoard())"   # report/index.html
```

Everything a run produces lands in `.out/<case-id>/`:

```text
case.docx     the file            measure-a.json   what the packed XML says
case.pdf      Word's export       measure-b.json   what the preview laid out
case.html     the preview page    measure-c.json   what Word made of it
preview-p*.png                    measure-x.json   how far apart those two are
word-p*.png                       result.json      one entry per assertion
```

## The four tiers

| | reads | proves | needs |
| --- | --- | --- | --- |
| **A** `ooxml` | the packed XML | the packer emitted the element at all | nothing |
| **B** `preview` | docx-preview in headless Chrome | what a person sees in the app we ship | Chrome |
| **C** `word` | Word 16 over COM | the only ground truth there is | Windows + Word |
| **X** `parity` | B against C | that the preview is not lying | both |

A can pass while the document is wrong. B can pass while Word disagrees. Only C
settles it, and X is the number that says whether the thing we hand people is
honest. A tier that cannot run here is `SKIPPED`, never silently green.

## Claims

`claim` is what turns a test suite into a roadmap.

| claim | all tiers pass | something fails |
| --- | --- | --- |
| `supported` | `PASS` | `FAIL` — a regression, break the build |
| `partial` + `knownRed: [...]` | `STALE` | `KNOWN` if only the named tiers are red |
| `unsupported` | `STALE` — promote it | `KNOWN` — the roadmap |

`KNOWN` does not fail the build; `FAIL` and `STALE` do. A case that starts
passing reports `STALE` rather than going quiet, so the suite cannot be made
green by deletion — the fix is to promote its claim, which is a diff a reviewer
sees.

## Writing a case

```tsx
export default defineCase({
  id: "text/align",
  title: "A paragraph set centred, ranged right, or justified",
  word: "Paragraph → Alignment (w:jc)",
  claim: "supported",

  style: withBlocks({ standfirst: { align: "center" } }),
  document: template(<Document id="align" title="Alignment">…</Document>),

  expect: {
    ooxml: (a, is) => is.equal(a.para("Centred").jc, "center", "centre writes w:jc center"),
    preview: (b, is) => is.within(b.textCentre("Centred"), b.mm(85), "2mm", "on the column's axis"),
    word: (c, is) => is.equal(c.para("Centred").alignment, "center", "Word reads it as centred"),
    parity: (p, is) =>
      is.within(p.previewX("Centred"), p.wordX("Centred"), "1mm", "where Word starts it"),
  },
});
```

A few things worth knowing before you write assertions.

**Measure drawn text, not boxes.** A centred paragraph's `<p>` is the full width
of the column whether it is centred or not; only the lines inside it move. The
preview view's `textLeft` / `textCentre` / `textRight` / `firstLineLeft` read
line rectangles for this reason.

**Anchors are matched case-insensitively.** `w:caps` prints a label in capitals
without changing the text, so the file says `Invoice reference` and Word says
`INVOICE REFERENCE`. An anchor names what a reader sees, and both spellings are
the same words.

**A cell is not a paragraph, and `para` will not find one.** Every probe leaves
a cell's paragraphs out of its paragraph slice deliberately, so that a
paragraph case's index means what it looks like it means. Tables have their own
vocabulary, and it is the same three words on all three tiers:

```tsx
a.table(0).row(1).cell(2)      // by position, in the file
a.cell("1,250.00")             // by the words a reader sees in it
b.cell("1,250.00").w           // the same cell, as the preview drew it
c.cell("1,250.00").width       // and as Word made it
p.previewCellX("1,250.00")     // where its words start, on screen
p.wordCellX("1,250.00")        // and in Word
```

`b.cellTextLeft` / `cellTextRight` / `cellTextCentre` / `cellTextTop` are the
cell's version of the paragraph lookups above, and read line rectangles for the
same reason.

**A chart is a part, not an element, and all three tiers say so.** What
`document.xml` holds is a frame and a relationship id; the plot and every
cached value are in `word/charts/chartN.xml`, and the numbers a reader would
open are in a workbook beside it. Probe A follows the whole chain, so a case
asserts on what the chart *is* rather than on the drawing that reserves its
space:

```tsx
a.chart(0).plot            // "bar", "line", "pie" …, from the chart part
a.chart(0).partPresent     // and that the part the drawing names is there
a.chart(0).workbookPresent // and the workbook "Edit Data" opens
a.chartSeries(0, 1).values // the second series, cached — what Word draws from
b.chart(0).w               // the frame, as the preview laid it out
b.chart(0).plotted         // and whether anything was drawn inside it
c.chart(0).typeName        // the chart Word actually built
c.chartSeries(0, 1).values // and the figures Word reads back out of it
p.previewChartWidth(0)     // the frame on screen against the frame in Word
```

**The preview draws a chart's plot and Word draws its own, so parity compares
frames.** docx-preview has no reading of a chart part at all, so it leaves an
empty span at the frame's size; the suite fills it with the same ECharts drawer
a scaffolded workspace uses, bundled straight out of
`templates/workspace/preview/charts.ts` so the tier measures what a person
sees. Holding those pixels against Word's would be measuring two chart
libraries. The frame is what both sides take from the file and what the page is
laid out around, and that is what tier X asserts.

**A cell's text includes its nested table's, so the innermost cell wins.** A
table is the only shape in a document that contains itself: the text of a cell
holding a table is everything printed inside it, so an outer cell matches every
anchor its inner table matches. All three probes therefore search the deepest
tables first, and an anchor names the cell a reader would point at. Where that
is not specific enough, `cellAt(table, row, column)` says exactly which.

**`column` is the grid column, not the cell's place in its row.** They are the
same number until something spans. On a row whose first cell spans two of three
columns, the second cell is sibling 1 and column 2, and a case that asserted
the first would pass while proving nothing.

## Things that bit us, so they need not bite again

- **`Range.Information(5)` always returns the left margin.** It reports where a
  line may begin, not where the glyphs do — identical for a centred, a
  right-ranged and a left paragraph. Only `Selection.Information` after
  `Select()` gives the real position. `Get-Info` in `probe-word.ps1` does that
  as its only path, not as a fallback.
- **Word reports run properties over a paragraph's whole range as "mixed".**
  The paragraph mark carries its own formatting, so a uniformly tracked label
  comes back `wdUndefined` for both caps and spacing. Measure the range with
  the mark moved off the end.
- **Word's `TabStops` includes its own default grid.** One custom stop reports
  four. Filter on `CustomTab`.
- **docx-preview puts run properties on a `<span>`, not the `<p>`.** Reading a
  face or a size off the paragraph returns the browser's default — Times New
  Roman at 16px, black — which looks like a document and is not one.
- **pdf.js will not render under Chrome's `--virtual-time-budget`.** The virtual
  clock does not advance a worker thread, so the render sits pending until the
  budget expires and the page comes back empty with no error. Hand the worker
  module over as `window.pdfjsWorker` to run it on the main thread.
- **Pixels do not fit through `--dump-dom`.** Carrying rendered pages back as
  base64 data URIs worked for one page and for two and silently produced
  nothing at three — no error, no partial output. Each page is now drawn alone
  at its own size and captured with `--screenshot`, which writes straight to
  disk. One Chrome run per page, plus a cheap first run that asks only how many
  pages there are and how big they are.
- **A chart needs Word to open the document writable *and* with a window.**
  Both were wrong at once and each looked like the other. Opened read-only,
  Word builds the chart object and reports the right type and title, and
  `SeriesCollection` comes back with the right count and no names and no values
  — the data load is deferred while the document cannot be edited. Opened with
  the `Visible` argument false, Word never builds the chart at all:
  `InlineShape.Type` still says 12 (`wdInlineShapeChart`), `HasChart` returns
  nothing and `.Chart` throws. Measured on one file: read-only gives `2024=`,
  writable gives `2024=12,18,9,22`; invisible gives `HasChart=[]`, visible gives
  `HasChart=[-1] ChartType=51`. `probe-word.ps1` now opens writable and
  visible, closes with `wdDoNotSaveChanges`, and leaves the *application*
  invisible, so nothing appears on screen.
- **Word driven over COM does not lay an inline shape out, so the Word column
  is the wrong evidence for anything a chart's *height* does.** The exported
  PDF shows a chart's plot missing and its line collapsed to about 25pt, and
  `Selection.Information(6)` on the paragraphs either side agrees — before and
  after `Repaginate()`, with a document window open. It is not the file and it
  is not charts: a control document of twelve 200pt *images* reports the same
  one page as twelve 200pt charts, and both are four pages in Word itself.
  Meanwhile `InlineShape.Height` is 200, `PlotArea.Height` is a real 117, and
  `Chart.Export` writes a correct PNG. So a chart case asserts through the
  chart object — its type, its series, its figures, its frame — and never
  through where the page put it. A pagination case involving an inline shape
  cannot be measured here at all, which is why there is not one.
- **`InlineShape.Type` and `Shape.Type` are different enumerations.** 12 is
  `wdInlineShapeChart` and 3 is `wdInlineShapePicture`; in `msoShapeType` 3 is
  a chart and 13 a picture. `Read-Shapes` walks both collections through one
  name table, so an inline chart reports `type12`. Charts have their own reader
  and do not depend on it.
- **Word holds a document open a moment after it is closed.** A run that
  measures a case and then rebuilds it can land inside that moment and get
  `EBUSY` for a reason that has nothing to do with the document. `buildCase`
  retries three times over a second and a half.
- **A non-greedy regex cannot read a table.** `<w:tbl>[\s\S]*?</w:tbl>` ends at
  the first closing tag, which for a table holding a table is halfway through
  the outer one — and `element(cellXml, "w:tcPr")` returns the *inner* cell's
  properties for an outer cell that holds one. Both give a plausible number for
  the wrong element, which is the worst kind. Probe A scans children by depth
  and reads an element's properties only from its head, the part before its
  first child, because OOXML always writes `w:tblPr`, `w:trPr` and `w:tcPr`
  first.
- **A range over a cell's contents reports the cell, not the text.** The range
  holds a block element — the paragraph docx-preview put there — and a range
  holding a block reports that block's box, which is the full width of the
  column whichever way the text inside it is set. Measured that way a
  right-ranged cell and a left one start in the same place. The paragraphs are
  measured instead, and the cell only directly when it holds none.
- **Word's cells are walked through the table's range, not through its rows.**
  `Table.Cell(r, c)`, `Table.Rows` and `Table.Columns` are all documented to
  fail on a table whose cells do not line up, and that is exactly the table a
  merge case is about — so the one reading that does not depend on the table
  being a plain grid is the one to use. `Table.Range.Cells` yields every cell
  with the `RowIndex` and `ColumnIndex` the other reading would have given.
  Measured so far: a horizontally spanned table reports `Uniform` as false and
  still gives up its rows, so this is a precaution rather than a scar — the
  suite cannot yet build a vertically merged table to find out (see
  `tables/rowspan`).
- **`settleDocxPreview` reads the packed file, never the theme.** Two of
  docx-preview's omissions — a run's letter spacing and a border's gap to its
  text — can only be put back from the file. Recomputing them from the style
  that produced the document would be a second copy of the packer's arithmetic
  living beside the first, and the copy nobody opens in Word is the one free to
  drift. `readPackedParagraphs` recovers them; settle is handed the result.

## What the suite must never do

It measures the framework; it does not improve it. A run touches nothing under
`src/`, `registry/`, `skills/` or `website/src/`. And a green board is evidence,
not proof — the last step before a release is still to open a `.docx` in Word
and look at it.
