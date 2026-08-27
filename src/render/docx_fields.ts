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
 */
export function fillPageFields(container: Element): void {
  const pages = [...container.querySelectorAll("section.docx")];

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
      .filter((element) => /^\s*(?:\d+\s*)?\/(?:\s*\d+)?\s*$/.test(element.textContent ?? ""));

    for (const holder of holders) {
      holder.textContent = `${index + 1} / ${pages.length}`;
    }
  });
}
