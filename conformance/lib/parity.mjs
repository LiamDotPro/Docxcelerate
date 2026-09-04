/**
 * Tier X — the preview and Word, held against each other.
 *
 * The other three tiers each measure one thing well. This one measures the
 * thing the framework actually promises: that what you see while you are
 * writing a document is what comes out of it. A property can be written
 * correctly into the file (tier A), drawn plausibly on screen (tier B), and
 * read correctly by Word (tier C), and the preview can still be showing the
 * page half a centimetre out — because nothing in those three ever compares
 * the two engines to each other.
 *
 * Everything here reduces both sides to the same frame — millimetres from the
 * top-left of the text column — so an assertion never converts anything and
 * never has a chance to convert it wrongly.
 *
 * @module
 */

import { PT_PER_MM, PX_PER_MM } from "./case.mjs";

/**
 * The two measurements, side by side.
 *
 * @param preview A `previewView`.
 * @param word A `wordView`.
 */
export function parityView(preview, word) {
  const pxToMm = (value) => (value === null || value === undefined ? null : round(value / PX_PER_MM));
  const ptToMm = (value) => (value === null || value === undefined ? null : round(value / PT_PER_MM));

  return {
    preview,
    word,

    // --- where a paragraph starts ------------------------------------------

    /** The left edge of the preview's drawn text, in mm from the text column. */
    previewX: (anchor) => pxToMm(preview.firstLineLeft(anchor)),
    /** Word's own answer to the same question. */
    wordX: (anchor) => ptToMm(word.para(anchor).x),

    /**
     * How far down the page the preview draws it, in mm.
     *
     * The first *line's* top, not the box's: a paragraph's box includes
     * whatever leading sits above its first line, and Word reports where the
     * character is. Comparing box to character is how a harness invents a
     * divergence that is really its own frame mismatch.
     */
    previewY: (anchor) => {
      const para = preview.para(anchor);
      return pxToMm(para.lines[0]?.y ?? para.y);
    },
    /** Word's answer. */
    wordY: (anchor) => ptToMm(word.para(anchor).y),

    // --- where it ends ------------------------------------------------------

    /** The right edge of the preview's drawn text, in mm. */
    previewRight: (anchor) => pxToMm(preview.textRight(anchor)),
    /** Word's, from the end of the paragraph's range. */
    wordRight: (anchor) => ptToMm(word.para(anchor).xEnd),

    /** How wide the drawn text is on screen, in mm. */
    previewWidth: (anchor) => pxToMm(preview.textWidth(anchor)),
    /**
     * And in Word — but only for a paragraph on one line, because Word reports
     * the end of the range, and on a wrapped paragraph that is the end of the
     * *last* line rather than the width of the first.
     */
    wordWidth(anchor) {
      const para = word.para(anchor);
      if (para.lineCount !== 1) return null;
      if (para.x === null || para.xEnd === null) return null;
      return round((para.xEnd - para.x) / PT_PER_MM);
    },

    // --- pagination ---------------------------------------------------------

    /** How many pages the preview broke the document into. */
    previewPages: () => preview.pageCount(),
    /** How many Word did. */
    wordPages: () => word.pageCount(),

    /** Which page the preview put a paragraph on, counting from one. */
    previewPage(anchor) {
      const page = preview.para(anchor).pageIndex;
      return page === null || page === undefined ? null : page + 1;
    },
    /** Which page Word put it on. */
    wordPage: (anchor) => word.para(anchor).page,

    // --- cells --------------------------------------------------------------
    //
    // The paragraph pair above will not find a cell: the preview's paragraph
    // slice is the article's own children and Word's leaves in-table
    // paragraphs out, both deliberately. A table's parity is asked of the cell
    // instead, and the two sides are matched by the words a reader sees in it.
    //
    // Word reports where a cell's *text* begins, which is inside whatever
    // padding the cell has. So the preview's side is its drawn text too, never
    // the cell box — comparing a box to a character is how a harness invents a
    // divergence the size of the padding.

    /** Where a cell's words start, in mm from the text column. */
    previewCellX: (anchor) => pxToMm(preview.cellTextLeft(anchor)),
    /** Word's answer to the same question. */
    wordCellX: (anchor) => ptToMm(word.cell(anchor).x),

    /** How far down the page the preview draws a cell's first line, in mm. */
    previewCellY(anchor) {
      const cell = preview.cell(anchor);
      return pxToMm(cell.lines[0]?.y ?? cell.y);
    },
    /** Word's. */
    wordCellY: (anchor) => ptToMm(word.cell(anchor).y),

    /** The furthest right a cell's words reach on screen, in mm. */
    previewCellRight: (anchor) => pxToMm(preview.cellTextRight(anchor)),

    /** How wide the preview draws the cell itself, in mm. */
    previewCellWidth: (anchor) => pxToMm(preview.cell(anchor).w),
    /** How wide Word makes it. */
    wordCellWidth: (anchor) => ptToMm(word.cell(anchor).width),

    /** Which page the preview put a cell on, counting from one. */
    previewCellPage(anchor) {
      const page = preview.cell(anchor).pageIndex;
      return page === null || page === undefined ? null : page + 1;
    },
    /** Which page Word put it on. */
    wordCellPage: (anchor) => word.cell(anchor).page,

    /** How wide the preview draws a whole table, in mm. */
    previewTableWidth: (index = 0) => pxToMm(preview.table(index).w),
    /** Where the preview puts its left edge, in mm from the text column. */
    previewTableX: (index = 0) => pxToMm(preview.table(index).x),
    // --- shapes --------------------------------------------------------------
    //
    // A shape is the one thing here whose size is stated in the file rather
    // than laid out from its contents, so the pair worth comparing is the box
    // itself: the preview's <svg> against Word's Shape.

    /** How wide the preview draws a shape, in mm. */
    previewShapeWidth: (index = 0) => pxToMm(preview.shape(index).w),
    /** How wide Word makes it. */
    wordShapeWidth: (index = 0) => ptToMm(word.shape(index).width),

    /** How deep the preview draws it, in mm. */
    previewShapeHeight: (index = 0) => pxToMm(preview.shape(index).h),
    /** How deep Word makes it. */
    wordShapeHeight: (index = 0) => ptToMm(word.shape(index).height),


    // --- running furniture --------------------------------------------------
    //
    // Both sides reduced to millimetres from the top of the *sheet*. A strip is
    // drawn outside the margins, so the text column — the frame everything else
    // here uses — is the wrong one: measured against it, every header would be
    // a negative number and every footer a number larger than the page.

    // Word's side of these is its *declared* distance rather than a measured
    // glyph position, and that is not a shortcut — it is the only reading
    // available. Asking Word where a header's range sits means selecting it,
    // and selecting it means moving the window's SeekView into the header pane;
    // done to a hidden instance that call kills Word outright ("the remote
    // procedure call failed", and the RPC server is gone for the rest of the
    // run, PDF export included). HeaderDistance is where Word puts the strip's
    // top, read from Word, without touching the view.

    /** How far down the sheet the preview draws the header, in mm. */
    previewHeaderY: (page = 1) => pxToMm(preview.furniture("header", page).y),
    /** Where Word puts the top of the header, from the top of the sheet. */
    wordHeaderY: () => ptToMm(word.headerDistance()),

    /** How far the preview's footer stops short of the foot of the sheet, in mm. */
    previewFooterFromBottom: (page = 1) =>
      pxToMm(preview.furniture("footer", page).fromBottom),
    /** Where Word puts the bottom of the footer, from the foot of the sheet. */
    wordFooterFromBottom: () => ptToMm(word.footerDistance()),

    /** Both answers at once, for a report that wants to show the pair. */
    compare(anchor) {
      return {
        anchor,
        previewX: pxToMm(preview.firstLineLeft(anchor)),
        wordX: ptToMm(word.para(anchor).x),
        previewY: pxToMm(preview.para(anchor).lines[0]?.y ?? preview.para(anchor).y),
        wordY: ptToMm(word.para(anchor).y),
        previewPage: (preview.para(anchor).pageIndex ?? -1) + 1,
        wordPage: word.para(anchor).page,
      };
    },
  };
}

function round(value) {
  return Math.round(value * 1000) / 1000;
}

/**
 * Every paragraph both engines found, paired by text.
 *
 * This is the report the board draws its parity column from: one row per
 * paragraph, with how far apart the two engines put it. Written as data rather
 * than assertions because it is a *survey* — a case says which paragraphs it
 * cares about, and this says what happened to all of them.
 */
export function parityTable(preview, word) {
  const rows = [];

  for (const para of preview.paragraphs) {
    const text = (para.text ?? "").trim();
    if (text.length === 0) continue;

    const match = word.paragraphs.find((other) => (other.text ?? "").trim() === text);
    if (match === undefined) {
      rows.push({ text: text.slice(0, 60), previewOnly: true });
      continue;
    }

    const setup = word.pageSetup ?? { leftMargin: 0, topMargin: 0 };
    const previewXmm = (para.lines[0]?.x ?? para.x) / PX_PER_MM;
    const previewYmm = (para.lines[0]?.y ?? para.y) / PX_PER_MM;
    const wordXmm = match.x === null ? null : (match.x - setup.leftMargin) / PT_PER_MM;
    const wordYmm = match.y === null ? null : (match.y - setup.topMargin) / PT_PER_MM;

    rows.push({
      text: text.slice(0, 60),
      previewPage: (para.pageIndex ?? 0) + 1,
      wordPage: match.page,
      dxMm: wordXmm === null ? null : round(previewXmm - wordXmm),
      dyMm: wordYmm === null ? null : round(previewYmm - wordYmm),
    });
  }

  // And every cell, by the same rule. Without this a table case's survey is
  // empty — the paragraph slices on both sides leave a cell's paragraphs out
  // on purpose — and the board would show a case with four green tiers and
  // nothing measured underneath.
  for (const cell of previewCells(preview)) {
    const text = (cell.text ?? "").trim();
    if (text.length === 0) continue;

    const match = word.cells(text)[0];
    if (match === undefined) {
      rows.push({ text: text.slice(0, 60), cell: true, previewOnly: true });
      continue;
    }

    const previewXmm = (cell.lines[0]?.x ?? cell.x) / PX_PER_MM;
    const previewYmm = (cell.lines[0]?.y ?? cell.y) / PX_PER_MM;

    rows.push({
      text: text.slice(0, 60),
      cell: true,
      previewPage: (cell.pageIndex ?? 0) + 1,
      wordPage: match.page,
      // Word's side is already content-relative: `wordView.cells` reduces it.
      dxMm: match.x === null ? null : round(previewXmm - match.x / PT_PER_MM),
      dyMm: match.y === null ? null : round(previewYmm - match.y / PT_PER_MM),
    });
  }

  return rows;
}

/** Every cell the preview drew, innermost tables first, in one flat list. */
function previewCells(preview) {
  return [...(preview.tables ?? [])]
    .sort((a, b) => (b.depth ?? 0) - (a.depth ?? 0))
    .flatMap((table) =>
      table.rows.flatMap((row) =>
        row.cells.map((cell) => ({ ...cell, pageIndex: table.pageIndex }))
      )
    );
}
