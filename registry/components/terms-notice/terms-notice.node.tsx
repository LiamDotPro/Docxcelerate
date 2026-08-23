import { Paragraph, type Section as SectionComponent, Section, useState } from "docxcelerate/template";

/**
 * The small print, as a section rather than a footnote.
 *
 * Boilerplate is the part of a document nobody rereads and everybody copies, so
 * it is worth having one node that owns it: the clauses arrive as data, each
 * becomes its own paragraph, and the id of each is stable, which is what lets a
 * later version of the terms be diffed against this one clause by clause.
 *
 * Reads well in the Legal Serif theme, which is what it was drawn against.
 *
 * Installed by `dxcl add terms-notice`.
 */

/** What this component reads. Add these fields to your document data type. */
export interface TermsData {
  terms: {
    /** The heading above the clauses. Defaults to "Terms". */
    heading?: string;
    /** One entry per clause, in the order they should be printed. */
    clauses: string[];
    /** Version or date of the terms, printed last so a reader can cite it. */
    version?: string;
  };
}

export const TermsNotice: SectionComponent = () => {
  const [terms] = useState((data: TermsData) => data.terms);

  return (
    <Section id="terms" title={terms.heading ?? "Terms"}>
      {terms.clauses.map((clause, index) => (
        <Paragraph id={`terms-clause-${index + 1}`}>
          {index + 1}. {clause}
        </Paragraph>
      ))}
      {terms.version && (
        <Paragraph id="terms-version">These terms are version {terms.version}.</Paragraph>
      )}
    </Section>
  );
};
