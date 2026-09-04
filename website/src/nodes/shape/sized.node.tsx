import { Paragraph, type Shape as ShapeComponent, Shape } from "docxcelerate/template";

/**
 * The same element given a width as well as a height.
 *
 * A shape with no width fills the text column, which is what a banner wants.
 * Saying one makes a block that sits in the column rather than spanning it —
 * a stamp, a callout, a badge beside prose.
 */
export const PaidStamp: ShapeComponent = () => (
  <Shape id="paid-stamp" variant="stamp" width={160} height={56}>
    <Paragraph id="paid-stamp-line">Paid in full</Paragraph>
  </Shape>
);
