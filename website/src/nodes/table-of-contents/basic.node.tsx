/** @jsxImportSource docxcelerate/template */
import { TableOfContents } from "docxcelerate/template";

/**
 * A heading standing in for the contents of the document around it.
 *
 * The shipped renderers print the title and stop — the entries themselves are
 * a renderer's job, and neither the browser preview nor the DOCX packer builds
 * them yet.
 */
export const Contents: TableOfContents = () => (
  <TableOfContents id="contents" title="What is in this letter" />
);
