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

/**
 * What a caller can tell the paginator that the DOM cannot.
 *
 * Only one thing so far, and it is the running furniture of a title-page
 * document. docx-preview picks one header and one footer per page it renders,
 * and it renders one page — so for a document whose first page differs it draws
 * the *first* page's strip and never asks for the other. There is nothing left
 * in the DOM for a new sheet to inherit but the letterhead, which is the one
 * thing that must not repeat.
 *
 * A host that can render the running parts — anything holding docx-preview and
 * the file — passes them here, and every sheet after the first takes those
 * instead. Elements rather than markup, because they come from docx-preview's
 * own rendering of the document's own parts: this is the file's running header,
 * not a second opinion about what one should look like.
 */
export interface PaginationOptions {
  /** The header every sheet after the first should carry. */
  runningHeader?: Element | null;
  /** The footer every sheet after the first should carry. */
  runningFooter?: Element | null;
}

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
export function paginateDocxPreview(
  container: Element,
  options: PaginationOptions = {},
): PaginationResult {
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

    const overflow = splitSection(section, options);

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
function splitSection(section: HTMLElement, options: PaginationOptions): HTMLElement | null {
  const article = section.querySelector(":scope > article");

  if (article === null) {
    return null;
  }

  const limit = contentBottomOf(section);

  if (limit === null) {
    return null;
  }

  const blocks = [...article.children] as HTMLElement[];

  // The first block whose *text* falls past the sheet's foot.
  //
  // The text, not the space under it: measured on eighteen two-line paragraphs
  // set 10pt apart, Word ends the last one 780.8pt down a page whose text stops
  // at 785.2 — and lets its trailing 10pt hang past that into the margin rather
  // than turning the page. So the border box is the right edge to test, and
  // adding the margin to it costs a paragraph a page.
  //
  // Never the first block on the page: something taller than a whole sheet has
  // to go somewhere, and moving it to a sheet of its own would move it for ever.
  const breakAt = blocks.findIndex((block, index) =>
    index > 0 && block.getBoundingClientRect().bottom > limit
  );

  if (breakAt === -1) {
    return null;
  }

  const next = newSheet(section, options);
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
  // The footer eats into the text only when it is bigger than the margin it
  // sits in, which is exactly Word's rule: the text stops at the bottom margin,
  // or at the top of the footer, whichever comes first.
  //
  // Measured both ways round. A ninety-paragraph document with no footer fits
  // eighteen to a page in Word; the same document with a one-line footer fits
  // seventeen, because that footer's own space-after pushes it up past the
  // margin. Getting this wrong is a page of disagreement by the fifth sheet.
  const footer = section.querySelector(":scope > footer");
  const reserve = footer === null
    ? 0
    : footerDistanceOf(footer, padding) + drawnHeightOf(footer, true);

  return rect.top + height - Math.max(padding, reserve);
}

/**
 * How far the footer stands from the foot of the paper, in pixels.
 *
 * Not written anywhere directly, but recoverable. docx-preview reserves the
 * strip a box of `margin − distance` — the slice of the margin left beneath it
 * — cancelled by an equal negative margin so it does not disturb the flow. The
 * distance is therefore the margin less that reserve, which is the number the
 * file declared in `w:pgMar/@w:footer` and the number Word lays out from.
 */
function footerDistanceOf(footer: Element, padding: number): number {
  const declared = Number.parseFloat((footer as HTMLElement).style.minHeight);

  return Number.isFinite(declared) ? Math.max(0, padding - declared) : padding;
}

/**
 * How tall the content of a running strip actually is.
 *
 * The union of its children's boxes rather than its own, for the reserve reason
 * above. A strip with nothing in it is nothing tall.
 */
function drawnHeightOf(strip: Element | null, withTrailingSpace = false): number {
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

  if (bottom <= top) {
    return 0;
  }

  // A footer's own space-after counts toward how much room it takes, even
  // though a body paragraph's does not. The difference is that a body paragraph
  // can let its trailing space hang into the margin, and a strip already lives
  // there — measured, it is what makes a one-line footer cost a page its
  // eighteenth paragraph.
  const view = strip.ownerDocument.defaultView;
  const last = strip.lastElementChild;
  const trailing = withTrailingSpace && view !== null && last !== null
    ? Number.parseFloat(view.getComputedStyle(last).marginBottom) || 0
    : 0;

  return bottom - top + trailing;
}

/**
 * A fresh sheet with the same page and the same running furniture.
 *
 * The header and footer are cloned rather than moved: they are drawn on every
 * page, which is what makes them running furniture, and the sheet they came
 * from still needs them. The article is cloned empty — its blocks are what the
 * caller is about to move across.
 */
function newSheet(section: HTMLElement, options: PaginationOptions): HTMLElement {
  const sheet = section.cloneNode(false) as HTMLElement;

  const running: Record<string, Element | null | undefined> = {
    HEADER: options.runningHeader,
    FOOTER: options.runningFooter,
  };

  for (const child of [...section.children]) {
    if (child.tagName === "ARTICLE") {
      sheet.appendChild(child.cloneNode(false));
      continue;
    }

    // The running strip, where the caller supplied one, in place of the strip
    // this sheet happens to be carrying. On a title-page document that is the
    // difference between every sheet repeating the letterhead and every sheet
    // after the first carrying the running header, which is what the file says
    // and what Word draws.
    const replacement = running[child.tagName];

    if (replacement !== undefined && replacement !== null) {
      const strip = replacement.cloneNode(true) as HTMLElement;
      // The reserve docx-preview sized for this sheet's strip, kept: it comes
      // from the page's own furniture distance, and the replacement was
      // rendered into a container that had no page to take it from.
      strip.setAttribute("style", (child as HTMLElement).getAttribute("style") ?? "");
      sheet.appendChild(strip);
      continue;
    }

    sheet.appendChild(child.cloneNode(true));
  }

  // A sheet that inherited no article — a section holding only furniture —
  // still needs one to receive the overflow.
  if (sheet.querySelector(":scope > article") === null) {
    sheet.appendChild(section.ownerDocument.createElement("article"));
  }

  return sheet;
}
