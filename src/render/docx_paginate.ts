/**
 * Breaking a preview into pages, the way Word breaks the file into pages.
 *
 * docx-preview breaks where the *file* says to — an explicit `w:br` or a
 * `w:pageBreakBefore` — and nowhere else. It does not lay text against a page
 * height and start a new sheet when the old one is full, because it has no
 * paginator. So a document Word prints on five pages is drawn as one very long
 * sheet, and everything read off that is wrong with it: which page a paragraph
 * is on, whether a header repeats, what a footer's "1 of 3" should say.
 *
 * This is the paginator. It is deliberately not part of
 * {@linkcode settleDocxPreview}: settling writes back facts the file declares
 * and needs no layout, so it can run in a jsdom at build time. Pagination needs
 * to know how tall things actually drew, which means it only works where
 * something has laid them out — a browser, or headless Chrome. Calling it
 * without layout is a no-op rather than a wrong answer, because every height it
 * would measure is zero.
 *
 * ## What it is, and is not
 *
 * It flows the body's blocks into boxes the height of the page and starts a new
 * sheet when one is full, carrying the running furniture onto it. That is what
 * Word does, from numbers the file already carries — the page size, the
 * margins, the header distance — so it is finishing the reading rather than
 * inventing an appearance.
 *
 * What it is not is Word's own paginator. It breaks *between* blocks, never
 * through the middle of one, so a paragraph that Word would split across a
 * sheet moves whole here instead. That is a real difference and it is measured
 * rather than assumed: `conformance/cases/preview/content-pagination` reports
 * how far the two engines' page counts and placements diverge, and the honest
 * summary is that they agree on documents made of ordinary paragraphs and drift
 * on documents made of very long ones.
 *
 * @module
 */

import { fillPageFields } from "./docx_fields.ts";

/** A point, in the CSS pixels a browser lays out in at 96dpi. */
const PX_PER_PT = 96 / 72;

/** How many sheets one document may be broken into before something is wrong. */
const PAGE_LIMIT = 200;

/** What one pagination pass did, for a caller that wants to say so. */
export interface PaginationResult {
  /** How many sheets the preview now holds. */
  pages: number;
  /** Whether anything was moved — false when the body already fitted. */
  changed: boolean;
  /**
   * Why nothing happened, when nothing did.
   *
   * `"no-layout"` means every height measured zero, which is what a DOM with
   * nothing laying it out returns. It is the one failure worth naming: the
   * answer looks like "the document fits on one page" and is not.
   */
  reason?: "no-layout" | "fits" | "limit";
}

/**
 * Breaks a rendered preview into pages the size of the document's own.
 *
 * Call it after `renderAsync` and after {@linkcode settleDocxPreview}, in a
 * context that lays out — a browser, or headless Chrome. Settling first
 * matters: it changes heights (a border's padding, a picture's wrapper), and
 * paginating against the unsettled heights would put the breaks in the wrong
 * places.
 *
 * @param container The element `renderAsync` rendered into.
 * @returns What the pass did.
 *
 * @example
 * ```ts
 * await renderAsync(bytes, body, head, { breakPages: true });
 * settleDocxPreview(body, model, await readPackedParagraphs(bytes));
 * paginateDocxPreview(body);
 * ```
 */
export function paginateDocxPreview(container: Element): PaginationResult {
  const sections = [...container.querySelectorAll("section.docx")] as HTMLElement[];

  if (sections.length === 0) {
    return { pages: 0, changed: false, reason: "fits" };
  }

  // Nothing here can work without layout, and a DOM that is not laying out
  // answers every measurement with zero — which reads as "it all fits" and is
  // the one wrong answer worth refusing to give.
  if (sections.every((section) => section.getBoundingClientRect().height === 0)) {
    return { pages: sections.length, changed: false, reason: "no-layout" };
  }

  let changed = false;
  let guard = 0;

  // A queue rather than a loop over the snapshot: splitting a section produces
  // another section that may itself need splitting, and the overflow of page
  // two is page three's problem.
  const queue = [...sections];

  while (queue.length > 0) {
    const section = queue.shift() as HTMLElement;

    if ((guard += 1) > PAGE_LIMIT) {
      return { pages: container.querySelectorAll("section.docx").length, changed, reason: "limit" };
    }

    const overflow = splitSection(section);

    if (overflow !== null) {
      changed = true;
      queue.push(overflow);
    }
  }

  // "1 of 5" is only right once there are five pages.
  //
  // `fillPageFields` fills the page number in from the pages it can see, and
  // settle runs before this does — so on a document that needed paginating it
  // filled in "1 / 1", and the split then copied that onto every sheet. The
  // numbers are recomputed here because this is the step that changed them.
  if (changed) {
    fillPageFields(container);
  }

  return {
    pages: container.querySelectorAll("section.docx").length,
    changed,
    reason: changed ? undefined : "fits",
  };
}

/**
 * Moves whatever will not fit on one sheet onto a new one.
 *
 * @returns The sheet the overflow went onto, or null when it all fitted.
 */
function splitSection(section: HTMLElement): HTMLElement | null {
  const article = section.querySelector(":scope > article");

  if (article === null) {
    return null;
  }

  const limit = contentBottomOf(section);

  if (limit === null) {
    return null;
  }

  const blocks = [...article.children] as HTMLElement[];

  // The first block whose foot falls past the sheet's. Never the first block on
  // the page: something taller than a whole page has to go somewhere, and
  // moving it to a sheet of its own would move it for ever.
  const breakAt = blocks.findIndex((block, index) =>
    index > 0 && block.getBoundingClientRect().bottom > limit
  );

  if (breakAt === -1) {
    return null;
  }

  const next = newSheet(section);
  const nextArticle = next.querySelector(":scope > article") as HTMLElement;

  for (const block of blocks.slice(breakAt)) {
    nextArticle.appendChild(block);
  }

  section.insertAdjacentElement("afterend", next);

  return next;
}

/**
 * How far down the viewport this sheet's text may reach.
 *
 * From the sheet's *declared* height rather than the height it currently has:
 * a section whose content overflowed has already grown past the page, and
 * measuring the grown box would say everything fits, for ever. docx-preview
 * writes the page height as the section's `min-height`, in points, which is the
 * one number here that comes straight from the file.
 */
function contentBottomOf(section: HTMLElement): number | null {
  const declared = Number.parseFloat(section.style.minHeight);

  if (!Number.isFinite(declared)) {
    return null;
  }

  const view = section.ownerDocument.defaultView;
  const styles = view === null ? null : view.getComputedStyle(section);

  if (styles === null) {
    return null;
  }

  const rect = section.getBoundingClientRect();
  const height = section.style.minHeight.endsWith("pt") ? declared * PX_PER_PT : declared;
  const padding = Number.parseFloat(styles.paddingBottom) || 0;

  // A footer is drawn inside the sheet and below the text, so the text stops
  // above it rather than behind it.
  //
  // Its *drawn* depth, not its box: docx-preview gives a footer a fixed reserve
  // — a min-height cancelled by an equal negative margin — so the box is the
  // same size on a page showing a bar as on one showing nothing. Subtracting
  // the reserve took a whole extra sheet out of a ninety-line document, which
  // is how this was found.
  // …and the strip sits *inside* the bottom margin, not below it. docx-preview
  // draws running furniture in the padding, which is what a margin is for on a
  // printed page. Subtracting both took a sliver off every sheet, and on a
  // ninety-line document the slivers added up to a whole extra page: six where
  // Word printed five.
  const footer = section.querySelector(":scope > footer");

  return rect.top + height - Math.max(padding, drawnHeightOf(footer));
}

/**
 * How tall the content of a running strip actually is.
 *
 * The union of its children's boxes rather than its own, for the reserve reason
 * above. A strip with nothing in it is nothing tall.
 */
function drawnHeightOf(strip: Element | null): number {
  if (strip === null) {
    return 0;
  }

  let top = Infinity;
  let bottom = -Infinity;

  for (const child of [...strip.querySelectorAll("*")]) {
    const rect = child.getBoundingClientRect();
    if (rect.height <= 0) continue;
    top = Math.min(top, rect.top);
    bottom = Math.max(bottom, rect.bottom);
  }

  return bottom > top ? bottom - top : 0;
}

/**
 * A fresh sheet with the same page and the same running furniture.
 *
 * The header and footer are cloned rather than moved: they are drawn on every
 * page, which is what makes them running furniture, and the sheet they came
 * from still needs them. The article is cloned empty — its blocks are what the
 * caller is about to move across.
 */
function newSheet(section: HTMLElement): HTMLElement {
  const sheet = section.cloneNode(false) as HTMLElement;

  for (const child of [...section.children]) {
    if (child.tagName === "ARTICLE") {
      sheet.appendChild(child.cloneNode(false));
    } else {
      sheet.appendChild(child.cloneNode(true));
    }
  }

  // A sheet that inherited no article — a section holding only furniture —
  // still needs one to receive the overflow.
  if (sheet.querySelector(":scope > article") === null) {
    sheet.appendChild(section.ownerDocument.createElement("article"));
  }

  return sheet;
}
