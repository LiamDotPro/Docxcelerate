import { Cell, Document, Image, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A picture set between words inside a table cell, as a status tick sits in a
 * statement line.
 *
 * The cell paragraph is cut at the picture's offset, so the drawing run lands
 * between two text runs of one `w:p` rather than after the line.
 *
 * **The Word tier is not written, deliberately.** It was written on a machine
 * without Word, and a measurement nobody took is not one to claim. Until it
 * goes in on the Windows machine, "supported" here means the file and the
 * preview, which is everything this case can see.
 */

// A one-pixel PNG: the shortest thing that is really an image.
const pngPixel = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

export default defineCase({
  id: "images/inline-in-cell",
  feature: "image.inlineInCell",
  title: "A picture mid-sentence in a table cell",
  word: "Insert → Picture → In Line with Text, between words inside a table cell (w:drawing run between w:r runs of one w:p)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="inline-in-cell" title="Picture in a cell">
      <Table id="payments" columns={[{ width: "auto" }]}>
        <Row>
          <Cell>
            <Paragraph id="paid">
              Paid <Image id="tick" src={pngPixel} alt="" width={8} height={8} /> on 3 March 2026 by bank transfer, reference INV-0042
            </Paragraph>
          </Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "paid", anchor: "Paid" },
  ],

  expect: {
    ooxml: (a, is) => {
      const x = a.cell("Paid").paragraphs[0].xml;

      is.equal(a.cell("Paid").paragraphs[0].hasDrawing, true, "the picture is packed in the cell paragraph");
      is.less(x.indexOf("Paid"), x.indexOf("<w:drawing"), "after the words before it");
      is.less(x.indexOf("<w:drawing"), x.indexOf("on 3 March"), "and before the words after it");
    },

    preview: (b, is) => {
      is.equal(b.chart(0).picture, true, "the preview draws the picture");
      is.less(
        b.chart(0).x,
        b.cell("Paid").x + b.cell("Paid").lines[0].w / 2,
        "in the first half of the line, not at its end",
      );
      is.equal(b.cell("Paid").lines.length, 1, "on a single line");
    },
  },
});
