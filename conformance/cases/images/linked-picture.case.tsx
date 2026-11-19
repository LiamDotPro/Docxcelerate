import { Document, Image, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A picture a reader can follow to an address, as a QR code or a Pay-now
 * button does.
 *
 * Word holds the link twice: a `w:hyperlink` round the run, which is what
 * Ctrl+click follows, and an `a:hlinkClick` inside the drawing, which is what
 * makes the picture itself the target. Both name one External relationship.
 *
 * **The Word tier is not written, deliberately.** The Word probe has no
 * reading of a hyperlink, and one nobody has run is a measurement nobody took.
 * Until it goes in on the Windows machine, "supported" here means the file and
 * the preview, which is everything this case can see.
 */
const ADDRESS = "https://pay.example.com/inv/123";

// A one-pixel PNG: the shortest thing that is really an image.
const pngPixel = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

export default defineCase({
  id: "images/linked-picture",
  feature: "image.hyperlink",
  title: "A picture that links to an address",
  word: "Insert → Link on a selected picture (w:hyperlink + a:hlinkClick, TargetMode=External)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="linked-picture" title="Linked picture">
      <Paragraph id="pay">
        Scan to pay <Image id="qr" src={pngPixel} href={ADDRESS} alt="Pay online" width={40} height={40} />
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "pay", anchor: "Scan to pay" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Scan to pay").hasDrawing, true, "the picture is packed");

      // The link has no words, so "" matches the only link in the document.
      is.equal(a.link("").text, "", "the link holds a picture, not words");
      is.equal(a.link("").target, ADDRESS, "which points at the address");
      is.equal(a.link("").external, true, "through a relationship marked external");
    },

    preview: (b, is) => {
      is.equal(b.chart(0).picture, true, "the preview draws the picture");
      is.equal(b.chart(0).href, ADDRESS, "inside a link to the same address");
    },
  },
});
