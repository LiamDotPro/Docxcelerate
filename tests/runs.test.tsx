import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDeriverRegistry,
  EchoAiClient,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import type { DocumentModel, ParagraphNode, RuntimeState, TableCellNode } from "docxcelerate";
import {
  Cell,
  createPublishData,
  Document,
  Paragraph,
  Row,
  Table,
  template,
  Text,
  useState,
} from "docxcelerate/template";
import { documentXml, partXml, textOf } from "./docx.ts";

/**
 * Some of a paragraph's words set differently from the rest.
 *
 * Two things are held here that the packed file cannot show by itself. One is
 * that a paragraph nobody marked is unchanged — no `runs` in the model, one
 * run in the file — because every document written before runs existed has to
 * go on publishing what it did. The other is that runs survive the trip an
 * engine takes them on: a published run still has a token in it, and settling
 * the token must leave the runs spelling the text.
 *
 * What Word and the preview make of the result is the conformance suite's to
 * say, under `conformance/cases/runs/`.
 */

function build(node: unknown, data: Record<string, unknown> = {}): Promise<DocumentModel> {
  return buildDocument(template<Record<string, unknown>>(node as never), data, {
    branchMode: "decide",
    dynamicMode: "placeholder",
  });
}

function paragraph(doc: DocumentModel, index = 0): ParagraphNode {
  return doc.nodes[index] as ParagraphNode;
}

/**
 * The runs of the paragraph holding this anchor, each as the words it prints
 * and whether it is bold. By anchor, because the document's title is a
 * paragraph too and would otherwise be run zero of every test.
 */
function runsOf(xml: string, anchor: string): Array<[string, boolean]> {
  const found = [...xml.matchAll(/<w:p>[\s\S]*?<\/w:p>/g)]
    .map((match) => match[0])
    .find((candidate) => candidate.includes(anchor)) ?? "";

  return [...found.matchAll(/<w:r>[\s\S]*?<\/w:r>/g)].map((match) => [
    [...match[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((text) => text[1]).join(""),
    match[0].includes("<w:b/>"),
  ]);
}

test("a marked word is a run of its own, and the words around it are left alone", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="terms">
        Payment is due within <Text bold>thirty days</Text> of the date above.
      </Paragraph>
    </Document>,
  );

  assertEquals(paragraph(doc).text, "Payment is due within thirty days of the date above.");
  assertEquals(paragraph(doc).runs, [
    { text: "Payment is due within " },
    { text: "thirty days", bold: true },
    { text: " of the date above." },
  ]);
  assertEquals(
    runsOf(await documentXml(doc), "Payment is due"),
    [["Payment is due within ", false], ["thirty days", true], [" of the date above.", false]],
  );
});

test("a paragraph nobody marked has no runs, in the model or in the file", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="plain">Dear {"Ada"}, nothing here is set differently.</Paragraph>
    </Document>,
  );

  assertEquals("runs" in paragraph(doc), false);
  assertEquals(runsOf(await documentXml(doc), "Dear Ada").length, 1);
});

test("one <Text> inside another takes both settings, and neighbours set alike are one run", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        <Text bold>Both </Text>
        <Text bold>
          halves, <Text italic>and this</Text>
        </Text>
      </Paragraph>
    </Document>,
  );

  assertEquals(paragraph(doc).runs, [
    { text: "Both halves, ", bold: true },
    { text: "and this", bold: true, italic: true },
  ]);
});

test("a link is packed as a hyperlink through an external relationship", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="pay">
        You can <Text href="https://example.com/pay">pay online</Text> before the due date.
      </Paragraph>
    </Document>,
  );

  assertStringIncludes(await documentXml(doc), ">pay online</w:t></w:r></w:hyperlink>");
  // The half Word refuses the file without: an address is only an address
  // when the relationship says it is outside the package. The id joining the
  // two is new on every pack, so it is the conformance probe that follows it.
  assertStringIncludes(
    await partXml(doc, "word/_rels/document.xml.rels"),
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/pay" TargetMode="External"',
  );
});

test("an address a reader should not be sent to prints as its words alone", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        Do not <Text href="javascript:alert(1)">follow this</Text>.
      </Paragraph>
    </Document>,
  );
  const xml = await documentXml(doc);

  assertEquals(xml.includes("<w:hyperlink"), false);
  assertStringIncludes(xml, "follow this");
});

test("a line end in the text is a break in the file, not whitespace", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="address">{"Flat 2\n14 Example Road\r\nLeeds LS1 4AB"}</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(
    xml,
    '>Flat 2</w:t><w:br/><w:t xml:space="preserve">14 Example Road</w:t><w:br/><w:t xml:space="preserve">Leeds LS1 4AB</w:t>',
  );
});

test("a published run keeps its token, and settling it leaves the runs spelling the text", async () => {
  const Greeting: Paragraph = () => {
    const [state] = useState((data: { name: string; payUrl: string }) => data);

    return (
      <Paragraph id="greeting">
        Dear <Text bold>{state.name}</Text>, you can <Text href={state.payUrl}>pay online</Text>.
      </Paragraph>
    );
  };
  const greeting = template<{ name: string; payUrl: string }>(
    <Document id="d" title="D">
      <Greeting />
    </Document>,
  );
  const data = { name: "Ada Lovelace", payUrl: "https://example.com/pay/42" };

  const published = await buildDocument(greeting, createPublishData() as typeof data, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });

  assertEquals(paragraph(published).runs?.[1], { text: "{{data.name}}", bold: true });
  assertEquals(paragraph(published).runs?.[3].href, "{{data.payUrl}}");

  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: new EchoAiClient(),
  };
  const settled = paragraph(await resolveDocument(published, state, { derivers: createDeriverRegistry() }));
  const built = paragraph(await buildDocument(greeting, data, { branchMode: "decide", dynamicMode: "placeholder" }));

  // What an engine writes is what the author previewed.
  assertEquals(settled.text, "Dear Ada Lovelace, you can pay online.");
  assertEquals(settled.runs, built.runs);
  assertEquals(settled.runs?.map((run) => run.text).join(""), settled.text);
  assertEquals(settled.runs?.[3].href, "https://example.com/pay/42");
});

test("runs that no longer spell the text are not printed", async () => {
  // What an engine on a resolver older than runs hands back: the text settled,
  // the runs still holding the token.
  const stale: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "d",
    title: "D",
    nodes: [{
      id: "p",
      kind: "paragraph",
      mode: "static",
      text: "Dear Ada,",
      runs: [{ text: "Dear " }, { text: "{{data.name}}", bold: true }, { text: "," }],
    }],
  };
  const xml = await documentXml(stale);

  assertStringIncludes(xml, ">Dear Ada,</w:t>");
  assertEquals(xml.includes("{{data.name}}"), false);
});

test("a cell's words can be marked without the cell becoming a list of nodes", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Table id="t" columns={[{}]}>
        <Row id="r">
          <Cell id="c">
            Total <Text bold>due</Text>
          </Cell>
        </Row>
      </Table>
    </Document>,
  );
  const table = doc.nodes[0] as { children: Array<{ children: TableCellNode[] }> };
  const words = table.children[0].children[0].children[0] as ParagraphNode;

  assertEquals(words.runs, [{ text: "Total " }, { text: "due", bold: true }]);
  assertEquals(runsOf(await documentXml(doc), "Total"), [["Total ", false], ["due", true]]);
});

test("a <Text> standing beside the nodes says where it belongs", async () => {
  await assertRejects(
    () =>
      build(
        <Document id="d" title="D">
          <Text bold>On its own</Text>
        </Document>,
      ),
    Error,
    "on its own it is a <Paragraph>",
  );
});

const SUP = '<w:vertAlign w:val="superscript"/>';
const SUB = '<w:vertAlign w:val="subscript"/>';

function count(xml: string, fragment: string): number {
  return xml.split(fragment).length - 1;
}

test("a raised word is a run of its own", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        84 m<Text script="superscript">2</Text>.
      </Paragraph>
    </Document>,
  );
  const xml = await documentXml(doc);

  assertEquals(paragraph(doc).runs, [
    { text: "84 m" },
    { text: "2", script: "superscript" },
    { text: "." },
  ]);
  assertStringIncludes(xml, `${SUP}</w:rPr><w:t xml:space="preserve">2</w:t>`);
  assertEquals(count(xml, "w:vertAlign"), 1);
});

test("a lowered word is a run of its own", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        CO<Text script="subscript">2</Text>
      </Paragraph>
    </Document>,
  );

  assertEquals(paragraph(doc).runs, [{ text: "CO" }, { text: "2", script: "subscript" }]);
  assertStringIncludes(await documentXml(doc), `${SUB}</w:rPr><w:t xml:space="preserve">2</w:t>`);
});

test("nested <Text> takes both marks, and the inner script wins", async () => {
  const bold = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        <Text bold>1<Text script="superscript">st</Text></Text>
      </Paragraph>
    </Document>,
  );
  assertEquals(paragraph(bold).runs, [
    { text: "1", bold: true },
    { text: "st", bold: true, script: "superscript" },
  ]);

  const both = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        <Text script="superscript">a<Text script="subscript">b</Text></Text>
      </Paragraph>
    </Document>,
  );
  const xml = await documentXml(both);
  assertEquals(paragraph(both).runs, [
    { text: "a", script: "superscript" },
    { text: "b", script: "subscript" },
  ]);
  assertEquals(count(xml, SUP), 1);
  assertEquals(count(xml, SUB), 1);
});

test("a paragraph without script has no vertical alignment", async () => {
  const plain = await build(
    <Document id="d" title="D">
      <Paragraph id="p">Due soon.</Paragraph>
    </Document>,
  );
  assertEquals(paragraph(plain).runs, undefined);
  assertEquals((await documentXml(plain)).includes("w:vertAlign"), false);

  const bold = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        Pay <Text bold>due</Text>
      </Paragraph>
    </Document>,
  );
  assertEquals(paragraph(bold).runs?.[1], { text: "due", bold: true });
  assertEquals((await documentXml(bold)).includes("w:vertAlign"), false);
});

test("a raised word in a cell", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Table id="t" columns={[{}]}>
        <Row id="r">
          <Cell id="c">
            Area m<Text script="superscript">2</Text>
          </Cell>
        </Row>
      </Table>
    </Document>,
  );
  const table = doc.nodes[0] as { children: Array<{ children: TableCellNode[] }> };
  const words = table.children[0].children[0].children[0] as ParagraphNode;

  assertEquals(words.runs, [{ text: "Area m" }, { text: "2", script: "superscript" }]);
  assertStringIncludes(await documentXml(doc), SUP);
});

test("a raised link is the footnote-style marker", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        statute<Text script="superscript" href="https://example.com/act">[1]</Text>.
      </Paragraph>
    </Document>,
  );

  assertEquals(
    /<w:hyperlink[^>]*>[\s\S]*?<w:vertAlign w:val="superscript"\/>[\s\S]*?>\[1\]<\/w:t><\/w:r><\/w:hyperlink>/
      .test(await documentXml(doc)),
    true,
  );
});

test("a stored model with an unknown script prints plain", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        84 m<Text script="superscript">2</Text>.
      </Paragraph>
    </Document>,
  );
  paragraph(doc).runs![1].script = "middle" as never;
  const xml = await documentXml(doc);

  assertEquals(xml.includes("w:vertAlign"), false);
  assertStringIncludes(textOf(xml), "84 m2.");
});
