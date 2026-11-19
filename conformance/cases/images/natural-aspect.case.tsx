import { Document, Image, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A picture given one side keeps its shape.
 *
 * Word's own Lock aspect ratio does this when a picture is dragged; here the
 * other side is worked out from the picture's header when the file is packed.
 * What is measured is the extent the file states, then that each engine drew
 * the frame at it.
 */
// 4x2 and 2x4 PNGs.
const pngWide = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAC0lEQVR42mNgwAUAABoAAS+Yl6YAAAAASUVORK5CYII=";
const pngTall = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAAECAIAAAArjXluAAAAC0lEQVR42mNgwA0AABwAAcdGaXkAAAAASUVORK5CYII=";

export default defineCase({
  id: "images/natural-aspect",
  feature: "image.naturalAspect",
  title: "A picture given one side keeps its proportions",
  word: "Insert → Pictures, resized with Lock aspect ratio (wp:extent cx/cy)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="natural-aspect" title="Images">
      <Paragraph id="before">A paragraph above the pictures.</Paragraph>
      <Image id="wide-width" src={pngWide} alt="Wide, by width" width={160} />
      <Image id="wide-height" src={pngWide} alt="Wide, by height" height={80} />
      <Image id="tall-width" src={pngTall} alt="Tall, by width" width={60} />
      <Paragraph id="after">A paragraph below them.</Paragraph>
    </Document>
  ),

  regions: [],

  expect: {
    ooxml: (a, is) => {
      const xml = a.documentXml ?? "";

      is.includes(xml, '<wp:extent cx="2028825" cy="1019175"/>', "a wide picture given a side is 2:1");
      is.includes(xml, '<wp:extent cx="762000" cy="1524000"/>', "and a tall one is 1:2");
    },

    preview: (b, is) => {
      const pictures = b.charts.filter((f) => f.picture);

      is.equal(pictures.length, 3, "the preview draws all three pictures");
      is.within(pictures[0].w, b.pt(160), "1mm", "the first at the width it was given");
      is.within(pictures[0].h, b.pt(80), "1mm", "and half as high");
      is.within(pictures[1].w, b.pt(160), "1mm", "the second's width follows its height");
      is.within(pictures[2].h, b.pt(120), "1mm", "the tall one is twice as high as wide");
    },

    word: (c, is) => {
      is.within(c.shape(0).width, 160, 1, "Word reads the first at 160pt");
      is.within(c.shape(0).height, 80, 1, "and 80pt high");
    },
  },
});
