/**
 * Settling a preview that docx-preview has just laid out.
 *
 * A preview is the packed `.docx` read back — that is the only way it cannot
 * drift from what Word shows. But docx-preview does not read everything the
 * file says: it drops a field run entirely, it looks for a table's indent
 * under an attribute that never carries it, and it wraps a picture in an
 * element a paragraph is not allowed to hold. Each one leaves the preview
 * showing something the document does not say.
 *
 * So the reading is finished here. Every function below writes back a fact the
 * packed file already declares and Word already draws — none of them invents
 * anything, and none of them is a stylesheet. The distinction matters: a
 * preview patched with CSS Word cannot reproduce is a preview that lies, and
 * the whole point of reading the file back is to stop it lying.
 *
 * Call {@linkcode settleDocxPreview} after `renderAsync` and before the
 * container is measured, screenshotted or serialised. It takes the container
 * docx-preview rendered into, so it works the same in a browser as it does in
 * a jsdom the build runs in.
 *
 * @module
 */

import type { DocumentModel } from "../domain/types.ts";

/** A millimetre, in points. */
const PT_PER_MM = 72 / 25.4;

/** How close to the paper's edge a page's own margin sits, by page size. */
const PAGE_WIDTH_MM: Record<string, number> = { A4: 210, Letter: 215.9 };

/**
 * Finishes docx-preview's reading of a document it has already laid out.
 *
 * @param container The element `renderAsync` rendered into.
 * @param model The document that was packed, for the facts the DOM cannot
 * supply on its own — page size and margins.
 *
 * @example
 * ```ts
 * await renderAsync(bytes, body, head, { breakPages: true });
 * settleDocxPreview(body, model);
 * ```
 */
export function settleDocxPreview(container: Element, model?: DocumentModel): void {
  inlinePictureWrappers(container);
  fillPageFields(container);

  if (model !== undefined) {
    applyTableIndents(container, model);
  }
}

/**
 * The box docx-preview puts a picture in, made something a paragraph can hold.
 *
 * A drawing is wrapped in a `<div>`, and docx-preview appends it to the `<p>`
 * its run belongs to. In a DOM built by hand that stands; serialised to HTML
 * and parsed again it does not, because a `<div>` may not sit inside a `<p>`.
 * The parser closes the paragraph at the picture, strands the words that
 * followed it, and leaves an empty paragraph where the close tag lands — one
 * taking the document's default leading rather than the line the file asked
 * for. Measured on an invoice, that turned a one-line footer bar 29px deep in
 * Word into 73px in the preview.
 *
 * The wrapper does the job of a `<span>` set `inline-block`, and a span is what
 * a paragraph may contain. Nothing about the layout changes; what changes is
 * that the layout survives being written down.
 */
export function inlinePictureWrappers(container: Element): void {
  for (const wrapper of [...container.querySelectorAll("p div")]) {
    const span = wrapper.ownerDocument.createElement("span");

    span.setAttribute("style", wrapper.getAttribute("style") ?? "");
    while (wrapper.firstChild !== null) {
      span.appendChild(wrapper.firstChild);
    }
    wrapper.replaceWith(span);
  }
}

/**
 * PAGE and NUMPAGES, filled in from the layout that just happened.
 *
 * docx-preview drops a field run on the floor — its `renderRun` returns null
 * for one — so a footer Word prints as "1 / 2" arrives as the bare "/" left
 * between the two dropped fields. The numbers come back from the rendered
 * section count, which is the count Word reaches too: it is the same file and
 * the same pagination.
 *
 * Only the separator the document itself wrote between the two fields is
 * touched, and only inside a footer. A page number written as a literal is
 * left exactly as the document wrote it.
 */
export function fillPageFields(container: Element): void {
  const pages = [...container.querySelectorAll("section.docx")];

  pages.forEach((page, index) => {
    const footer = page.querySelector(":scope > footer");
    if (footer === null) {
      return;
    }

    // The innermost element holding the bare separator, so the number lands
    // beside the slash rather than replacing the whole footer.
    const holders = [...footer.querySelectorAll("*")]
      .filter((element) => element.children.length === 0)
      .filter((element) => (element.textContent ?? "").replace(/\s+/g, "") === "/");

    for (const holder of holders) {
      holder.textContent = `${index + 1} / ${pages.length}`;
    }
  });
}

/**
 * The indent a bleeding table declares, put back.
 *
 * docx-preview reads a table's indent with `parseIndentation`, which looks for
 * a `w:left` attribute — and `w:tblInd` does not carry one, its value is in
 * `w:w`. So the indent is parsed and then dropped, and a table the file says
 * reaches the paper's edge stops at the margin instead. Word honours it.
 *
 * A bleeding table is known by its own width: docx-preview writes the table's
 * declared width onto the element, and nothing but a bleed makes a table wider
 * than the text column it stands in. That keeps this from needing to know
 * which variant a table was given.
 */
export function applyTableIndents(container: Element, model: DocumentModel): void {
  const margins = model.style?.page?.margins;
  if (margins === undefined) {
    return;
  }

  const pageMm = PAGE_WIDTH_MM[model.style?.page?.size ?? "A4"] ?? PAGE_WIDTH_MM.A4;
  const columnPt = (pageMm - margins.leftMm - margins.rightMm) * PT_PER_MM;

  for (const table of [...container.querySelectorAll("table")]) {
    const element = table as HTMLElement;
    const declared = Number.parseFloat(element.style.width);

    // A point of slack: the declared width is written back from twips, so a
    // table sitting exactly on the column can round either way.
    if (!Number.isFinite(declared) || declared <= columnPt + 1) {
      continue;
    }

    element.style.marginInlineStart = `-${margins.leftMm}mm`;
    element.style.marginInlineEnd = `-${margins.rightMm}mm`;
  }
}
