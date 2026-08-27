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
import { fillPageFields } from "./docx_fields.ts";
import type { PackedParagraph } from "./docx_packed.ts";

export { fillPageFields } from "./docx_fields.ts";
export { paginateDocxPreview, type PaginationResult } from "./docx_paginate.ts";

export {
  type PackedBorderSpace,
  type PackedParagraph,
  type PackedRun,
  readPackedParagraphs,
  readPart,
} from "./docx_packed.ts";

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
 * @param packed What the file says about its own paragraphs, from
 * {@linkcode readPackedParagraphs}. Two of docx-preview's omissions can only be
 * put back from the file itself; without this they stay omitted, and the
 * preview is as complete as docx-preview left it rather than wrong.
 *
 * @example
 * ```ts
 * const bytes = new Uint8Array(await blob.arrayBuffer());
 * await renderAsync(bytes, body, head, { breakPages: true });
 * settleDocxPreview(body, model, await readPackedParagraphs(bytes));
 * ```
 */
export function settleDocxPreview(
  container: Element,
  model?: DocumentModel,
  packed?: readonly PackedParagraph[],
): void {
  inlinePictureWrappers(container);
  fillPageFields(container);

  if (model !== undefined) {
    applyTableIndents(container, model);
  }

  if (packed !== undefined) {
    applyPackedParagraphs(container, packed);
  }
}

/**
 * The DOM paragraph a packed paragraph was drawn into, found by its words.
 *
 * By text rather than by position: docx-preview renders headers and footers
 * into the same container, and the packer inserts paragraphs of its own — a
 * break carrier, a separator between two tables — so counting would drift, and
 * a drifted match writes one paragraph's appearance onto another's words. Each
 * element is claimed once, so two paragraphs that happen to say the same thing
 * take one match each.
 *
 * A paragraph that cannot be found is skipped. That is the safe direction: a
 * miss leaves the preview exactly as docx-preview drew it, where a wrong match
 * would make it say something the document never did.
 */
function matchParagraphs(
  container: Element,
  packed: readonly PackedParagraph[],
): Array<{ packed: PackedParagraph; element: HTMLElement }> {
  const wanted = packed.filter((paragraph) => paragraph.text.trim().length > 0);
  if (wanted.length === 0) {
    return [];
  }

  const elements = [...container.querySelectorAll("p")] as HTMLElement[];
  const claimed = new Set<Element>();
  const pairs: Array<{ packed: PackedParagraph; element: HTMLElement }> = [];

  for (const paragraph of wanted) {
    const text = normalise(paragraph.text);
    const element = elements.find(
      (candidate) => !claimed.has(candidate) && normalise(candidate.textContent ?? "") === text,
    );

    if (element !== undefined) {
      claimed.add(element);
      pairs.push({ packed: paragraph, element });
    }
  }

  return pairs;
}

/** Whitespace-insensitive, because a renderer may rewrap but never rewords. */
function normalise(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/**
 * The two things docx-preview drops on the way from the file to the DOM.
 *
 * Both are read out of the packed bytes rather than recomputed from the theme
 * that produced them. That distinction is the whole discipline of this module:
 * recomputing would be a second copy of the packer's arithmetic living beside
 * the first, and the copy that nobody opens in Word is the one free to drift.
 * What is written back here is the number in the file, which is the number
 * Word draws.
 */
function applyPackedParagraphs(container: Element, packed: readonly PackedParagraph[]): void {
  for (const pair of matchParagraphs(container, packed)) {
    applyTracking(pair.element, pair.packed);
    applyBorderPadding(pair.element, pair.packed);
  }
}

/**
 * The tracking a run declares, put back.
 *
 * docx-preview reads `w:spacing` only when its parent is `w:pPr` — the branch
 * is guarded `if (elem.localName == "pPr")` — so the identical element inside
 * `w:rPr`, which is where Word keeps character spacing, is dropped on the
 * floor. Word draws it.
 *
 * The cost is not subtle. Measured on a 7pt label tracked 0.12em, Word set it
 * 29.1mm wide and the preview 23.8mm: a fifth narrower, which is the difference
 * between a label that fits its column and one that wraps.
 *
 * Runs are matched to spans in order, since docx-preview draws one span per
 * run. Where the two counts disagree — a picture among the words, a run split
 * on a tab — every span takes the first tracking the paragraph declares, which
 * is right for the paragraphs that carry tracking at all: a label is set in one
 * face throughout, or it is not a label.
 */
function applyTracking(element: HTMLElement, packed: PackedParagraph): void {
  const tracked = packed.runs.filter((run) => run.letterSpacingPt !== undefined);
  if (tracked.length === 0) {
    return;
  }

  const spans = [...element.querySelectorAll("span")] as HTMLElement[];
  const targets = spans.length === 0 ? [element] : spans;
  const aligned = targets.length === packed.runs.length;

  targets.forEach((target, index) => {
    const spacing = aligned
      ? packed.runs[index].letterSpacingPt
      : tracked[0].letterSpacingPt;

    if (spacing !== undefined) {
      target.style.letterSpacing = `${spacing}pt`;
    }
  });
}

/**
 * The gap between a paragraph's border and its words, put back.
 *
 * `w:pBdr` carries a `w:space` on each edge — the distance Word holds the text
 * off the rule — and docx-preview's `parseBorderProperties` writes only the
 * `border-*` shorthand, never the space. So a panel padded 8pt draws its border
 * tight against its own words on screen and 8pt clear of them in Word, and
 * every paragraph below it sits about 3mm too high.
 *
 * Only the edges the file actually draws a border on carry a gap, because those
 * are the only edges Word pads: `w:space` lives on the border element, so an
 * edge with no border has nowhere to record one. A block that asks for padding
 * without a border gets none here *and* none in Word — a limitation of the
 * format rather than of the reading, and the reason a filled panel that must
 * hold its words off its own edge has to be a table cell.
 *
 * The two axes are not the same shape, which is the part only a measurement
 * tells you. Measured on an 8pt-padded panel: Word moves the text *down* from a
 * top border, so the vertical gap is padding as CSS means it; but it does not
 * move the text *in* from a left border — it moves the border out into the
 * margin and leaves the words on the column. So the horizontal gap is padding
 * with a matching negative margin, which draws the same picture: the rule steps
 * outward, the text does not move.
 */
function applyBorderPadding(element: HTMLElement, packed: PackedParagraph): void {
  const sides = [
    ["top", "Top"],
    ["right", "Right"],
    ["bottom", "Bottom"],
    ["left", "Left"],
  ] as const;

  for (const [side, capitalised] of sides) {
    const gap = packed.borderSpacePt[side];
    if (gap === undefined || gap === 0) {
      continue;
    }

    element.style[`padding${capitalised}`] = `${gap}pt`;

    // Left and right: the border steps out, the text stays where it was.
    if (side === "left" || side === "right") {
      element.style[`margin${capitalised}`] = `-${gap}pt`;
    }
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
