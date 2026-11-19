import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDeriverRegistry,
  EchoAiClient,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentModel, ParagraphNode, RuntimeState } from "docxcelerate";
import { settleDocxPreview } from "docxcelerate/preview";
import {
  Cell,
  createPublishData,
  Document,
  PageBreak,
  Paragraph,
  Row,
  Table,
  template,
  useState,
} from "docxcelerate/template";
import { JSDOM } from "jsdom";
import { documentXml, paragraphOf, partXml } from "./docx.ts";

/**
 * A paragraph that is an item of a list.
 *
 * What is held here is the counting, because that is the part an author cannot
 * do and the part a mistake hides in: which items are one list, where a list
 * ends, and that a second one starts again. Each is a fact about which
 * numbering a paragraph names, so each is read out of the packed file.
 *
 * What the marker looks like on the page, and where the words stand against
 * it, is the conformance suite's to say, under `conformance/cases/lists/`.
 */

function build(node: unknown): Promise<DocumentModel> {
  return buildDocument(template<Record<string, never>>(node as never), {}, {
    branchMode: "decide",
    dynamicMode: "placeholder",
  });
}

/** Which numbering a paragraph counts in, and how deep, or null when it is no item. */
function numberingOf(xml: string, anchor: string): { numId: string; level: string } | null {
  const found = /<w:numPr><w:ilvl w:val="(\d+)"\/><w:numId w:val="(\d+)"\/><\/w:numPr>/.exec(
    paragraphOf(xml, anchor),
  );

  return found === null ? null : { numId: found[2], level: found[1] };
}

/** The space a paragraph leaves under itself, in twips. */
function spaceAfter(xml: string, anchor: string): string | undefined {
  return /<w:spacing [^>]*w:after="(\d+)"/.exec(paragraphOf(xml, anchor))?.[1];
}

/** The format of a numbering's first level, followed from the `w:num` to its abstract one. */
function formatOf(numbering: string, numId: string): string | undefined {
  const abstract = new RegExp(
    `<w:num w:numId="${numId}"[^>]*><w:abstractNumId w:val="(\\d+)"`,
  ).exec(numbering)?.[1];

  return new RegExp(
    `<w:abstractNum [^>]*w:abstractNumId="${abstract}"[\\s\\S]*?<w:numFmt w:val="(\\w+)"`,
  ).exec(numbering)?.[1];
}

test("a list prop puts the paragraph's place in the model, and nothing else changes", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" list="bullet">An item</Paragraph>
      <Paragraph id="b" list="number" listLevel={1}>A step under one</Paragraph>
      <Paragraph id="c">Not an item</Paragraph>
    </Document>,
  );
  const [a, b, c] = doc.nodes as ParagraphNode[];

  // A first-level item says no level: absent is what "not under anything" is.
  assertEquals(a.list, { kind: "bullet" });
  assertEquals(b.list, { kind: "number", level: 1 });
  assertEquals("list" in c, false);
});

test("items standing together are one list, and a numbered list after a break starts its own", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="s1" list="number">First step</Paragraph>
      <Paragraph id="s1a" list="bullet" listLevel={1}>A point under it</Paragraph>
      <Paragraph id="s2" list="number">Second step</Paragraph>
      <Paragraph id="between">Something that is not an item</Paragraph>
      <Paragraph id="t1" list="number">Another list</Paragraph>
    </Document>,
  );
  const xml = await documentXml(doc);
  const numbering = await partXml(doc, "word/numbering.xml");
  const first = numberingOf(xml, "First step");
  const second = numberingOf(xml, "Second step");
  const point = numberingOf(xml, "A point under it");
  const again = numberingOf(xml, "Another list");

  assertEquals(formatOf(numbering, first?.numId ?? ""), "decimal");
  assertEquals(formatOf(numbering, point?.numId ?? ""), "bullet");
  assertEquals(point?.level, "1");
  // The bulleted point between them does not stop the steps counting.
  assertEquals(second?.numId, first?.numId);
  // A paragraph that is no item does, and the next list is a numbering of its own.
  assertEquals(numberingOf(xml, "Something that is not"), null);
  assertEquals(again !== null && again.numId !== first?.numId, true);
  assertEquals(formatOf(numbering, again?.numId ?? ""), "decimal");
});

test("each numbered list has an abstract numbering of its own, for the preview's sake", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" list="number">One list</Paragraph>
      <Paragraph id="x">Between</Paragraph>
      <Paragraph id="b" list="number">Another</Paragraph>
    </Document>,
  );
  const xml = await documentXml(doc);
  const numbering = await partXml(doc, "word/numbering.xml");
  const abstractOf = (anchor: string) =>
    new RegExp(`<w:num w:numId="${numberingOf(xml, anchor)?.numId}"[^>]*><w:abstractNumId w:val="(\\d+)"`)
      .exec(numbering)?.[1];

  // docx-preview keeps one `w:num` per abstract numbering, so two lists over
  // one would leave the first drawn with no numbers.
  assertEquals(abstractOf("One list") !== undefined, true);
  assertEquals(abstractOf("One list") !== abstractOf("Another"), true);
});

test("items stand together, and the space a paragraph leaves goes after the list", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" list="bullet">First item</Paragraph>
        <Paragraph id="b" list="bullet">Last item</Paragraph>
        <Paragraph id="c">After the list</Paragraph>
      </Document>,
    ),
  );

  assertEquals(spaceAfter(xml, "First item"), "0");
  assertEquals(spaceAfter(xml, "Last item"), spaceAfter(xml, "After the list"));
  assertEquals(spaceAfter(xml, "Last item") !== "0", true);
});

test("an item says its hang on the paragraph, where the preview can read it", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" list="bullet">An item</Paragraph>
        <Paragraph id="b" list="bullet" listLevel={1}>A point under it</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(paragraphOf(xml, "An item"), '<w:ind w:left="360" w:hanging="360"/>');
  assertStringIncludes(paragraphOf(xml, "A point under it"), '<w:ind w:left="720" w:hanging="360"/>');
});

test("a document with no list in it is packed with the numbering it always had", async () => {
  const plain = await build(
    <Document id="d" title="D">
      <Paragraph id="a">No list here</Paragraph>
    </Document>,
  );
  const listed = await build(
    <Document id="d" title="D">
      <Paragraph id="a" list="bullet">A list here</Paragraph>
    </Document>,
  );
  const count = async (doc: DocumentModel) =>
    [...(await partXml(doc, "word/numbering.xml")).matchAll(/<w:abstractNum /g)].length;

  // One is the packing library's own, written whether or not anything uses it.
  assertEquals(await count(plain), 1);
  assertEquals(await count(listed), 2);
});

test("an item in a cell, and an item straight after a page break, are still items", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{}]}>
          <Row id="r">
            <Cell id="c">
              <Paragraph id="in-cell" list="bullet">In a cell</Paragraph>
            </Cell>
          </Row>
        </Table>
        <PageBreak id="break" />
        <Paragraph id="after-break" list="number">After a break</Paragraph>
      </Document>,
    ),
  );

  assertEquals(numberingOf(xml, "In a cell")?.level, "0");
  assertEquals(numberingOf(xml, "After a break")?.level, "0");
  // The break is carried by a paragraph style, and so is being an item. Word
  // takes one style per paragraph, so this is where the two could collide.
  assertEquals([...paragraphOf(xml, "After a break").matchAll(/<w:pStyle /g)].length, 1);
});

test("a generated paragraph can be an item, the same as a written one", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph
        id="next"
        list="bullet"
        generalPrompt="Say what happens next."
        placeholder="What happens next."
      />
    </Document>,
  );

  assertEquals((doc.nodes[0] as ParagraphNode).mode, "dynamic");
  assertEquals((doc.nodes[0] as ParagraphNode).list, { kind: "bullet" });
});

test("steps written from a request's own list are one list once an engine has settled them", async () => {
  interface Steps {
    steps: Array<{ text: string }>;
  }
  const StepList = () => {
    const [steps] = useState((data: Steps) => data.steps);

    return steps.map((step) => <Paragraph id="step" list="number">{step.text}</Paragraph>);
  };
  const steps = template<Steps>(
    <Document id="d" title="D">
      <Paragraph id="intro">What to do:</Paragraph>
      <StepList />
    </Document>,
  );
  const data: Steps = { steps: [{ text: "Fill it in" }, { text: "Sign it" }, { text: "Post it" }] };

  // Published, the loop is one body an engine walks: how many steps there are
  // belongs to a request nobody has made.
  const published = await buildDocument(steps, createPublishData() as Steps, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data as unknown as Record<string, unknown>),
    aiClient: new EchoAiClient(),
  };
  const xml = await documentXml(
    await resolveDocument(published, state, { derivers: createDeriverRegistry() }),
  );

  // Settled, its passes stand next to one another — so they count as one list.
  const ids = ["Fill it in", "Sign it", "Post it"].map((text) => numberingOf(xml, text)?.numId);
  assertEquals(ids[0] !== undefined, true);
  assertEquals(ids, [ids[0], ids[0], ids[0]]);
});

test("the preview lets an item's marker tab to its hang, and leaves other paragraphs alone", () => {
  // What docx-preview draws for an item: the class its marker hangs on, and
  // the hang written inline from the indent the file states on the paragraph.
  const dom = new JSDOM(
    '<div id="page">' +
      '<p class="docx_listparagraph docx-num-2-0" style="text-indent: -18.00pt; margin-inline-start: 18.00pt;"><span>An item</span></p>' +
      '<p style="margin-bottom: 10pt;"><span>Not an item</span></p>' +
      "</div>",
  );
  const page = dom.window.document.getElementById("page") as Element;

  settleDocxPreview(page);

  const [item, plain] = [...page.querySelectorAll("p")].map((p) => p.getAttribute("style") ?? "");
  assertStringIncludes(item, "white-space: pre-wrap; tab-size: 18.00pt;");
  assertEquals(plain, "margin-bottom: 10pt;");
});
