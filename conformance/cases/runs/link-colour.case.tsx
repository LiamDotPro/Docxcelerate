import { Document, Paragraph, template, Text } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A link drawn in the ink of the block it sits in.
 *
 * A link used to be set in the palette's accent wherever it stood, which on a
 * dark band is a link nobody can read. It now keeps the colour its block gave
 * the words, and falls back to the accent only when the block gave none.
 *
 * **The Word tier is not written**, for the reason the hyperlink case gives.
 */
export default defineCase({
  id: "runs/link-colour",
  feature: "run.hyperlink.color",
  title: "A link in a coloured block takes the block's colour",
  word: "Insert → Link, in a paragraph whose font colour is white (w:hyperlink, w:color)",
  claim: "supported",

  style: withBlocks({ band: { fill: "1F2933", color: "FFFFFF" } }),

  document: template(
    <Document id="link-colour" title="Link colour">
      <Paragraph id="contact" variant="band">
        Questions? <Text href="mailto:billing@example.com">Email billing</Text>.
      </Paragraph>
      <Paragraph id="pay">
        You can <Text href="https://example.com/pay">pay online</Text> any time.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "band", anchor: "Questions?" },
    { id: "plain", anchor: "You can pay online" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.run("Email billing").color, "FFFFFF", "a link in a band is drawn in the band's colour");
      is.equal(a.run("pay online").color, "2F5FBD", "a link in a plain paragraph is still the accent");
    },

    preview: (b, is) => {
      is.equal(b.run("Email billing").color, b.hex("FFFFFF"), "the preview draws the band's link in the same colour");
    },
  },
});
