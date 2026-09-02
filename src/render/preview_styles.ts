/**
 * How a run of pages is presented: the ground they sit on, the space between
 * them, and the edge that makes a sheet look like a sheet.
 *
 * This is chrome, and only chrome. Nothing here reaches inside a page — a
 * page's content is whatever the packed file says, and a preview that drew
 * something Word will not would be lying about the document, which is the one
 * thing reading the file back exists to prevent. What it styles is the space
 * *between* sheets, which is not in the file because it is not part of the
 * document: Word draws it too, and calls it the desk.
 *
 * It is one function rather than a rule in each preview because there are two
 * of them — the site's baked embeds and the app a scaffolded workspace runs —
 * and they had drifted. One showed pages on a desk with a gap and a shadow;
 * the other ran them together on white, so a two-page document read as one
 * long sheet with a seam in it.
 *
 * @module
 */

/** What {@linkcode previewPageStyles} lets a viewer decide for itself. */
export interface PreviewPageStyleOptions {
  /**
   * The ground the sheets sit on.
   *
   * Anything but white: a sheet is only distinct from the desk if the desk is
   * a different colour from the sheet.
   */
  ground?: string;
  /** Space between one sheet and the next. Defaults to `28px`. */
  gap?: string;
  /**
   * Space around the whole run.
   *
   * `0` for an embed framed to the width of a page, where padding would inset
   * the first sheet inside a frame drawn to fit it. A scrolling viewer wants
   * room at the top and rather more at the bottom, so the last page can be
   * scrolled clear of the edge.
   */
  inset?: string;
  /**
   * Whether a sheet is lifted off the desk.
   *
   * A hairline is enough where the frame is small and a shadow would be most
   * of what you see; a full viewer wants the shadow.
   */
  raised?: boolean;
}

/**
 * The stylesheet a preview needs around its pages.
 *
 * Written against docx-preview's own structure — `section.docx` per page,
 * inside a `.docx-wrapper` when it was asked for one — so it works whether or
 * not the caller rendered `inWrapper`.
 *
 * @param options What this viewer wants of the desk.
 * @returns CSS, ready for a `<style>` element.
 *
 * @example
 * ```ts
 * const style = document.createElement("style");
 * style.textContent = previewPageStyles({ inset: "32px 0 72px" });
 * head.append(style);
 * ```
 */
export function previewPageStyles(options: PreviewPageStyleOptions = {}): string {
  const ground = options.ground ?? "#e6e8ed";
  const gap = options.gap ?? "28px";
  const inset = options.inset ?? "0";
  const raised = options.raised ?? true;

  // A hairline under the shadow rather than instead of it: a shadow alone
  // leaves the top edge of a sheet on a light desk with nothing to sit on.
  const edge = raised
    ? "box-shadow: 0 0 0 1px rgb(15 23 42 / 8%), 0 1px 2px rgb(15 23 42 / 10%), " +
      "0 14px 34px rgb(15 23 42 / 18%);"
    : "box-shadow: 0 0 0 1px rgb(15 23 42 / 12%);";

  return [
    `html, body { margin: 0; padding: 0; background: ${ground}; }`,
    `body { padding: ${inset}; }`,
    // docx-preview's own wrapper paints a desk of its own and pads it. Ours is
    // on the body, so the two would be a desk inside a desk.
    ".docx-wrapper { background: transparent !important; padding: 0 !important; }",
    "section.docx {",
    "  margin: 0 auto !important;",
    "  background: #ffffff;",
    "  box-sizing: border-box;",
    `  ${edge}`,
    "}",
    `section.docx + section.docx { margin-top: ${gap} !important; }`,
  ].join("\n");
}
