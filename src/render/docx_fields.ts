/**
 * The fields docx-preview drops, filled in from the pages that exist.
 *
 * Its own module because two steps need it and neither owns it. Settling calls
 * it, because a document that fits on one page is finished the moment it is
 * read back. Pagination calls it again, because pagination is what changes the
 * answer — a footer settled as "1 / 1" on a sheet that then became five is a
 * footer that lies on all five.
 *
 * @module
 */

import type { DocumentPageStyle, PageNumberStyle } from "../domain/style.ts";

/**
 * PAGE and NUMPAGES, filled in from the layout that just happened.
 *
 * docx-preview drops a field run on the floor — its `renderRun` returns null
 * for one — so a footer Word prints as "1 / 2" arrives as the bare "/" left
 * between the two dropped fields. The numbers come back from the rendered
 * section count, which is the count Word reaches too: it is the same file and,
 * once the preview paginates, the same pagination.
 *
 * Only the separator the document itself wrote between the two fields is
 * touched, and only inside a footer. A page number written as a literal is left
 * exactly as the document wrote it.
 *
 * @param container The element the preview was rendered into.
 * @param numbering Where the page numbers start and how they are written. Left
 * out, it is read back from the sheets, where an earlier call stamped it.
 */
export function fillPageFields(
  container: Element,
  numbering?: Pick<DocumentPageStyle, "firstPageNumber" | "numberStyle">,
): void {
  const pages = [...container.querySelectorAll("section.docx")];

  // `newSheet` clones the section with `cloneNode(false)`, which copies
  // attributes, and the settled DOM is what gets serialised into the pagination
  // frame. So the second pass sees what the first pass was told.
  if (numbering !== undefined) {
    for (const page of pages) {
      if (numbering.firstPageNumber !== undefined) {
        page.setAttribute("data-first-page-number", String(numbering.firstPageNumber));
      }
      if (numbering.numberStyle !== undefined) {
        page.setAttribute("data-number-style", numbering.numberStyle);
      }
    }
  }
  const stamped = pages[0];
  const start = Number(stamped?.getAttribute("data-first-page-number") ?? 1);
  const style = (stamped?.getAttribute("data-number-style") ?? "decimal") as PageNumberStyle;

  pages.forEach((page, index) => {
    const footer = page.querySelector(":scope > footer");
    if (footer === null) {
      return;
    }

    // The innermost element holding the separator, so the number lands beside
    // the slash rather than replacing the whole footer.
    //
    // Either the bare separator docx-preview left behind, or a number this
    // function already wrote. It has to be both, because it runs twice: once
    // when the preview is settled, and again after pagination, which is what
    // changes the answer. Matching only the bare form meant the second pass
    // found nothing to correct and every sheet of a five-page document kept the
    // "1 / 1" the first pass had written.
    const holders = [...footer.querySelectorAll("*")]
      .filter((element) => element.children.length === 0)
      .filter((element) => /^\s*(?:[\dA-Za-z]+\s*)?\/(?:\s*\d+)?\s*$/.test(element.textContent ?? ""));

    for (const holder of holders) {
      holder.textContent = `${formatPageNumber(start + index, style)} / ${pages.length}`;
    }
  });
}

const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"],
  [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
];

function formatPageNumber(n: number, style: PageNumberStyle): string {
  // shortcut: page 0 in roman/letter is written as 0; Word's own rendering of it is unverified
  if (style === "decimal" || n < 1) {
    return String(n);
  }
  const glyphs = style.endsWith("Roman")
    ? ROMAN.reduce(
      ([rest, out], [value, glyph]) => [rest % value, out + glyph.repeat(Math.floor(rest / value))] as [number, string],
      [n, ""] as [number, string],
    )[1]
    // Word doubles the letter after Z (AA, BB) rather than counting in base 26.
    : String.fromCharCode(65 + (n - 1) % 26).repeat(Math.ceil(n / 26));
  return style.startsWith("lower") ? glyphs.toLowerCase() : glyphs;
}
