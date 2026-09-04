import { Document, Paragraph, Shape, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { COLUMN_MM, withBlocks } from "../_support/style.ts";

/**
 * A drawn rectangle with words on top of it.
 *
 * The one thing in the model that is a shape rather than text: Word draws a
 * real `Rectangle`, and the paragraphs inside sit on its fill rather than
 * beside it. What separates it from the filled table cell `tables/cell-borders`
 * measures is that the box is a size the document decided — a shape does not
 * grow with its text, and that is the whole reason to reach for one.
 *
 * **It is packed as VML, and that is a decision rather than an accident.** Word
 * reads a `w:pict` holding a `v:rect` back as a genuine shape, with the width,
 * height, fill and text the file gives it — measured below. It would read a
 * DrawingML `wps:wsp` the same way, and Word writes that form itself. The
 * difference is entirely on the other side: docx-preview renders VML and has
 * no reading of `wps:wsp` at all, so a DrawingML shape draws nothing on
 * screen. Packing the form both engines read is what keeps the preview showing
 * the document; the alternative was to pack DrawingML and then build the box
 * in the preview, which is a second renderer.
 *
 * One repair is still needed and the parity tier is what justifies it.
 * docx-preview nests the `<foreignObject>` carrying the words *inside* the
 * `<rect>`, where SVG paints no children — measured in Chrome, the words
 * report a bounding box of 0 by 0 there and 92 by 17 when moved out beside it.
 * So a shape drew as a filled block with nothing on it, which is the worst
 * kind of preview bug because it looks deliberate. `liftShapeText` moves the
 * node; both elements come from the file, both are drawn by Word, and only
 * their nesting was wrong.
 */
const WIDTH_PT = 300;
const HEIGHT_PT = 60;
const PADDING_PT = 10;

export default defineCase({
  id: "shapes/block-with-text",
  feature: "shape.blockWithText",
  title: "A drawn rectangle with words on top of it",
  word: "Insert → Shapes → Rectangle, with text (w:pict / v:rect / v:textbox)",
  claim: "supported",

  style: withBlocks({
    /** A dark callout, sized by the case, with its words centred on it. */
    callout: {
      fill: "1F2933",
      color: "FFFFFF",
      weight: "bold",
      fontSizePt: 14,
      paddingPt: PADDING_PT,
      align: "center",
    },
    /** The same box, ruled, to prove a stroke reaches both engines. */
    outlined: {
      fill: "EEF2FF",
      color: "1F2933",
      border: "2F5FBD",
      borderWidthPt: 1.5,
      paddingPt: PADDING_PT,
    },
  }),

  document: template(
    <Document id="block-with-text" title="Shapes">
      <Paragraph id="before">A paragraph above the shape.</Paragraph>

      <Shape id="callout" variant="callout" width={WIDTH_PT} height={HEIGHT_PT}>
        <Paragraph id="callout-line">Paid in full</Paragraph>
      </Shape>

      <Shape id="outlined" variant="outlined" height={HEIGHT_PT}>
        <Paragraph id="outlined-line">Ruled, and the width of the text column</Paragraph>
      </Shape>

      <Paragraph id="after">A paragraph below it.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "callout", anchor: "Paid in full" },
    { id: "outlined", anchor: "Ruled, and the width" },
  ],

  expect: {
    ooxml: (a, is) => {
      const xml = a.documentXml ?? "";

      is.includes(xml, "<w:pict>", "the shape is packed as a picture element");
      is.includes(xml, "<v:rect", "holding a VML rectangle");
      is.includes(xml, "<v:textbox", "with a text box on it");
      is.includes(xml, "<w:txbxContent>", "holding the document's own paragraphs");
      // The form Word writes and docx-preview cannot read. Packing it would
      // draw in Word and vanish on screen.
      is.excludes(xml, "wps:wsp", "and not as a DrawingML shape, which the preview cannot draw");

      is.includes(xml, `width:${WIDTH_PT}pt;height:${HEIGHT_PT}pt`, "the sized shape carries its geometry");
      is.includes(xml, 'fillcolor="#1F2933"', "and the theme's fill");
      is.includes(xml, 'strokecolor="#2F5FBD"', "the ruled one carries its stroke");
      is.includes(xml, 'stroked="f"', "and the unruled one says it has none");
      is.includes(xml, "Paid in full", "the words are in the file");

      // The room inside, said in the one place both engines read it. VML
      // states it as an inset on the text box and docx-preview never reads
      // that attribute, so room stated only there is room Word leaves and the
      // screen does not.
      is.includes(xml, 'inset="0pt,10pt,0pt,10pt"', "the text box insets only the top and bottom");
      is.includes(xml, '<w:ind w:left="200" w:right="200"/>', "and the sides are an indent instead");
    },

    preview: (b, is) => {
      is.equal(b.shapes.length, 2, "the preview draws both shapes");
      is.within(b.shape(0).w, b.pt(WIDTH_PT), "1mm", "the first at the width the file gives it");
      is.within(b.shape(0).h, b.pt(HEIGHT_PT), "1mm", "and the height");
      is.within(b.shape(1).w, b.mm(COLUMN_MM), "1mm", "the second fills the text column");
      is.equal(b.shape(0).fill, "#1F2933", "the first draws the theme's fill");
      is.equal(b.shape(1).stroke, "#2F5FBD", "and the second its stroke");

      // The assertion `liftShapeText` exists for. Nested inside the rect these
      // are 0 by 0, and a filled block with invisible words looks deliberate.
      is.greater(b.shapeTextWidth("Paid in full"), 0, "the words on the first shape are drawn");
      is.greater(b.shapeTextWidth("Ruled, and the width"), 0, "and on the second");

      // And they are set in from the fill's edge rather than jammed against
      // it. This is the assertion the indent exists for: with the room stated
      // only as the text box's own inset, the words drew hard against the box
      // on screen while Word set them in, and the preview was lying.
      is.within(
        b.shape(1).lines[0].x - b.shape(1).x,
        b.pt(PADDING_PT),
        "1mm",
        "the words are set in from the edge of the fill",
      );
    },

    word: (c, is) => {
      is.equal(c.shapes.length, 2, "Word reads two shapes");
      is.equal(c.shape(0).type, "autoShape", "as real shapes rather than pictures");
      is.within(c.shape(0).width, WIDTH_PT, 1, "the first is 300pt wide");
      is.within(c.shape(0).height, HEIGHT_PT, 1, "and 60pt deep");
      is.within(c.shape(1).width, c.mm(COLUMN_MM), 1, "the second fills the text column");
      is.equal(c.shape(0).fill, "1F2933", "Word fills the first as the theme says");
      is.equal(c.shape(0).hasText, true, "and reads words on it");
      is.includes(c.shape(0).text, "Paid in full", "which are the document's own");
      is.includes(c.shape(1).text, "Ruled, and the width", "on both of them");
    },

    /**
     * Where the shapes and their words land, on screen against in Word. The
     * three tiers above agree that a shape exists and says the right thing;
     * only this one says the preview drew it in the same place at the same
     * size.
     */
    parity: (p, is) => {
      is.within(p.previewShapeWidth(0), p.wordShapeWidth(0), "1mm", "the first shape is as wide on screen as in Word");
      is.within(p.previewShapeHeight(0), p.wordShapeHeight(0), "1mm", "and as deep");
      is.within(p.previewShapeWidth(1), p.wordShapeWidth(1), "1mm", "and so is the second");
    },
  },
});
