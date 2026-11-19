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
import type { DocumentModel, DocumentNode, PageNumberFormat } from "../domain/types.ts";

/** What a holder stands for: the form its node asked for, and the words it held. */
type PageField = { format: PageNumberFormat; text: string };

/**
 * PAGE and NUMPAGES, filled in from the layout that just happened.
 *
 * docx-preview drops a field run on the floor — its `renderRun` returns null
 * for one — so a footer Word prints as "1 / 2" arrives as the bare "/" left
 * between the two dropped fields. The numbers come back from the rendered
 * section count, which is the count Word reaches too: it is the same file and,
 * once the preview paginates, the same pagination.
 *
 * Only the words the document itself wrote beside the fields are touched, and
 * only inside a footer: the separator between two numbers, or the prefix before
 * one, found from `model`. Without one, the bare `/` a default `<PageNumber />`
 * leaves behind. A single number with no prefix, or a separator of only
 * whitespace, leaves nothing to anchor on and stays blank in the preview; Word
 * is always right. A page number written as a literal is left as written.
 *
 * @param container The element the preview was rendered into.
 * @param numbering Where the page numbers start and how they are written. Left
 * out, it is read back from the sheets, where an earlier call stamped it.
 * @param model The document that was packed, for the words each page number
 * wrote. Without it only the default `/` form is found.
 */
export function fillPageFields(
  container: Element,
  numbering?: Pick<DocumentPageStyle, "firstPageNumber" | "numberStyle">,
  model?: DocumentModel,
): void {
  const anchors = pageFieldAnchors(model);
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

    for (const holder of footer.querySelectorAll("*")) {
      if (holder.children.length > 0) continue;

      // Marked when first found: once filled, its words are a number nothing
      // could find again, and pagination fills it again without a model. The
      // marks travel with the clone pagination makes of a footer.
      if (!holder.hasAttribute("data-page-field")) {
        const field = anchors.get((holder.textContent ?? "").trim());
        if (field === undefined) continue;
        holder.setAttribute("data-page-field", field.format);
        holder.setAttribute("data-page-text", field.text);
      }

      const current = formatPageNumber(start + index, style);
      const format = holder.getAttribute("data-page-field");
      const text = holder.getAttribute("data-page-text") ?? "";
      holder.textContent = format === "current"
        ? `${text}${current}`
        : format === "total"
        ? `${text}${pages.length}`
        : `${current}${text}${pages.length}`;
    }
  });
}

/** The words each footer page number leaves once its fields are dropped, keyed by what they trim to. */
function pageFieldAnchors(model: DocumentModel | undefined): Map<string, PageField> {
  const anchors = new Map<string, PageField>([["/", { format: "currentOfTotal", text: " / " }]]);
  const walk = (nodes: readonly DocumentNode[] | undefined): void => {
    for (const node of nodes ?? []) {
      if (node.kind === "pageNumber") {
        const format = node.format ?? "currentOfTotal";
        // Between the numbers for "3 of 10", before the one number otherwise.
        const text = format === "currentOfTotal" ? node.separator ?? " / " : node.prefix ?? "";
        if (text.trim() !== "") anchors.set(text.trim(), { format, text });
      } else if ("children" in node) {
        walk(node.children);
      }
    }
  };

  walk(model?.footer);
  walk(model?.firstFooter);
  walk(model?.evenFooter);
  return anchors;
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
