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

  return rows;
}
