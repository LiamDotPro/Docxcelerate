import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { documentXml, paragraphOf } from "./docx.ts";
import type { DocumentModel, DocumentStyle, ShapeNode } from "docxcelerate";
import { cleanMinimalDocumentStyle } from "docxcelerate";
import { Document, Paragraph, Section, Shape, template, Text } from "docxcelerate/template";

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
    signLine: { tabStopsMm: [{ at: 60, align: "right", leader: "dot" }] },
    inset: {
      indentMm: 10,
      indentRightMm: 10,
      border: "2F5FBD",
      borderSides: ["left"],
      borderWidthPt: 2,
      paddingPt: 6,
      fill: "EEF2FF",
    },
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

test("the words centre on the box drawn, not on the depth the theme falls back to", async () => {
  // The stamp that found this: a 56pt shape over a block style saying 44pt.
  // Exact leading the height of the box is what sits one line in the middle of
  // it, so leading struck from the variant centred the words in a box 12pt
  // shorter than the one drawn and left them riding high.
  const doc = await build(() => <Shape id="s" variant="deep" height={56}>On it</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, "height:56pt");
  assertStringIncludes(xml, 'w:line="1120" w:lineRule="exact"');
});

test("the theme's fill is the shape's, and its padding the room inside", async () => {
  const doc = await build(() => <Shape id="s" variant="callout">Paid in full</Shape>);
  const xml = await xmlOf(doc);

  assertStringIncludes(xml, 'fillcolor="#1F2933"');
  // Top and bottom only. The sides are an indent instead, because
  // docx-preview turns a `v:textbox` into a bare `foreignObject` and never
  // reads `inset` — room stated only there is room Word leaves and the screen
  // does not, which is the preview lying about the document.
  assertStringIncludes(xml, 'inset="0pt,10pt,0pt,10pt"');
  assertStringIncludes(xml, '<w:ind w:left="200" w:right="200"/>');
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

test("a paragraph in a shape keeps its own alignment", async () => {
  const xml = await xmlOf(
    await build(() => (
      <Shape id="s" variant="callout">
        <Paragraph align="right">Ranged right</Paragraph>
      </Shape>
    )),
  );

  assertStringIncludes(paragraphOf(xml, "Ranged right"), '<w:jc w:val="right"/>');
});

test("a paragraph's own indents add to the room inside the box", async () => {
  const xml = await xmlOf(
    await build(() => (
      <Shape id="s" variant="callout">
        <Paragraph variant="inset">Quoted</Paragraph>
      </Shape>
    )),
  );
  const found = paragraphOf(xml, "Quoted");

  assertStringIncludes(found, '<w:ind w:left="767" w:right="767"/>');
  assertStringIncludes(found, 'w:fill="EEF2FF"');
  assertStringIncludes(found, "<w:pBdr>");
});

test("a tab and a mark in a shape's paragraph survive, with the stops its variant names", async () => {
  const xml = await xmlOf(
    await build(() => (
      <Shape id="s" variant="callout">
        <Paragraph variant="signLine">{"Quoted\t"}<Text bold>in full</Text></Paragraph>
      </Shape>
    )),
  );
  const found = paragraphOf(xml, "Quoted");

  assertStringIncludes(found, "<w:tab/>");
  assertStringIncludes(found, '<w:tab w:val="right" w:pos="3402" w:leader="dot"/>');
  assertStringIncludes(found, "<w:b/>");
});
