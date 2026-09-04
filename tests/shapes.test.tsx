import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { documentXml } from "./docx.ts";
import type { DocumentModel, DocumentStyle, ShapeNode } from "docxcelerate";
import { cleanMinimalDocumentStyle } from "docxcelerate";
import { Document, Paragraph, Section, Shape, template } from "docxcelerate/template";

/**
 * A drawn box with words on it.
 *
 * The thing worth testing here is not that a rectangle appears — it is *which*
 * rectangle. Word reads two forms of shape and docx-preview reads one of them,
 * so the packed form is a decision the preview depends on, and a change that
 * quietly moved to the other form would draw correctly in Word and vanish on
 * screen. That is the failure these cases exist to catch, and it is why they
 * assert the element names rather than only the geometry.
 */

const style: DocumentStyle = {
  ...cleanMinimalDocumentStyle,
  blocks: {
    ...cleanMinimalDocumentStyle.blocks,
    callout: { fill: "1F2933", color: "FFFFFF", paddingPt: 10 },
    ruled: { fill: "EEF2FF", border: "2F5FBD", borderWidthPt: 1.5 },
    deep: { fill: "1F2933", heightPt: 90 },
  },
};

function build(body: () => unknown) {
  const Body = body as () => never;

  return buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Section id="body" title="Body">
          <Body />
        </Section>
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
}

function shapeOf(doc: DocumentModel): ShapeNode {
  const section = doc.nodes[0];

  if (section?.kind !== "section") {
    throw new Error("expected the body section");
  }

  const shape = section.children[0];

  if (shape?.kind !== "shape") {
    throw new Error(`expected a shape, got ${shape?.kind ?? "nothing"}`);
  }

  return shape;
}

/** The packed XML, with the style the blocks above are declared in. */
function xmlOf(doc: DocumentModel): Promise<string> {
  return documentXml({ ...doc, style });
}

// ---------------------------------------------------------------------------
// The shape a shape carries
// ---------------------------------------------------------------------------

test("a shape given text becomes the paragraph it would have had to be written as", async () => {
  const doc = await build(() => <Shape id="s" variant="callout">Paid in full</Shape>);
  const shape = shapeOf(doc);

  assertEquals(shape.children.length, 1);
  assertEquals(shape.children[0]?.kind, "paragraph");
});

test("a shape given paragraphs keeps them apart, because two lines are two lines", async () => {
  const doc = await build(() => (
    <Shape id="s" variant="callout">
      <Paragraph id="a">Paid in full</Paragraph>
      <Paragraph id="b">Thank you</Paragraph>
    </Shape>
  ));

  assertEquals(shapeOf(doc).children.length, 2);
});

test("the size is the node's, because a shape is a box you decided the size of", async () => {
  const doc = await build(() => <Shape id="s" width={300} height={60}>On it</Shape>);

  assertEquals(shapeOf(doc).width, 300);
  assertEquals(shapeOf(doc).height, 60);
});

// ---------------------------------------------------------------------------
// What it packs as
// ---------------------------------------------------------------------------

test("a shape packs as VML, which is the form both Word and the preview read", async () => {
  // The decision this whole feature rests on. Word reads DrawingML and VML
  // alike; docx-preview renders VML and has no reading of `wps:wsp` at all, so
  // the other form draws in Word and shows nothing on screen.
  const doc = await build(() => <Shape id="s" variant="callout">Paid in full</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, "<w:pict>");
  assertStringIncludes(xml, "<v:rect");
  assertStringIncludes(xml, "<v:textbox");
  assertStringIncludes(xml, "<w:txbxContent>");
  assertEquals(xml.includes("wps:wsp"), false);
});

test("the picture sits inside a run, which is the only place a run's reader looks", async () => {
  // `w:pict` beside a run rather than inside one is repaired silently by Word
  // and not at all by docx-preview, which walks a run's children looking for
  // one. The shape then draws in Word and vanishes on screen.
  const doc = await build(() => <Shape id="s">On it</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, "<w:r><w:pict>");
});

test("the geometry is written in points, in the attribute both engines read", async () => {
  const doc = await build(() => <Shape id="s" width={300} height={60}>On it</Shape>);

  assertStringIncludes(await xmlOf(doc), 'style="width:300pt;height:60pt"');
});

test("a shape with no width fills the text column", async () => {
  // 210mm less two 25.4mm margins is 159.2mm, which is 451.28pt.
  const doc = await build(() => <Shape id="s" height={40}>On it</Shape>);

  assertStringIncludes(await xmlOf(doc), "width:451.28pt");
});

test("a block that states a depth is a shape's height, because it is already saying that", async () => {
  const doc = await build(() => <Shape id="s" variant="deep">On it</Shape>);

  assertStringIncludes(await xmlOf(doc), "height:90pt");
});

test("the theme's fill is the shape's, and its padding the room inside", async () => {
  const doc = await build(() => <Shape id="s" variant="callout">Paid in full</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, 'fillcolor="#1F2933"');
  assertStringIncludes(xml, 'inset="10pt,10pt,10pt,10pt"');
});

test("a ruled shape states its stroke twice, because the two engines read different halves", async () => {
  // Word reads `strokecolor` on the shape; docx-preview's `parseStroke` reads
  // `color` off a `v:stroke` child and never looks at the attribute. Both are
  // built from the one block, so they cannot come to disagree.
  const doc = await build(() => <Shape id="s" variant="ruled">Ruled</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, 'strokecolor="#2F5FBD"');
  assertStringIncludes(xml, '<v:stroke color="#2F5FBD" weight="1.5pt"');
});

test("a shape with no border says so, rather than leaving Word to pick one", async () => {
  const doc = await build(() => <Shape id="s" variant="callout">Paid in full</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, 'stroked="f"');
  assertEquals(xml.includes("<v:stroke "), false);
});

test("the words are the document's own, run properties and all", async () => {
  const doc = await build(() => <Shape id="s" variant="callout">Paid in full</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, "Paid in full");
  assertStringIncludes(xml, '<w:color w:val="FFFFFF"/>');
});

test("an empty shape still gets a paragraph, because Word will not take a bare text box", async () => {
  const doc = await build(() => <Shape id="s" variant="callout" height={20} />);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, "<w:txbxContent><w:p");
});
