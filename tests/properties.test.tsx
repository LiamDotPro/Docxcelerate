import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDocumentProjectArtifact,
  defineDocumentProject,
  type DocumentModel,
  type DocumentStyle,
  InMemoryDataProvider,
  resolveDocument,
  type RuntimeState,
} from "docxcelerate";
import { Document, Paragraph, template } from "docxcelerate/template";
import { partXml } from "./docx.ts";

const style: DocumentStyle = {
  preset: "test",
  page: { size: "A4", orientation: "portrait", margins: { topMm: 20, rightMm: 20, bottomMm: 20, leftMm: 20 } },
  typography: { bodyFont: "Aptos", headingFont: "Aptos", bodySizePt: 11, bodyLineHeight: 1.4, color: "111827" },
  paragraph: { spacingAfterPt: 10 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 18 },
  sectionHeading: { fontSizePt: 12, weight: "bold", spacingBeforePt: 16, spacingAfterPt: 7 },
};

type Data = { invoice: { number: string } };

function notIn(xml: string, text: string): void {
  assertEquals(xml.includes(text), false);
}

function stateOf(data: Record<string, unknown>): RuntimeState {
  return {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider({ ...data }),
    aiClient: {
      generateParagraph: () => "",
      generateImage: () => ({ path: "x.png", alt: "x" }),
      generateGraph: () => ({ data: { series: [{ values: [1] }] }, caption: "x" }),
    },
  };
}

test("no properties writes the title and neither an author nor Un-named", async () => {
  const t = template(<Document id="d" title="Welcome"><Paragraph id="p">Hi</Paragraph></Document>);
  const doc = await buildDocument(t, {});
  assertEquals("properties" in doc, false);
  const core = await partXml(doc, "docProps/core.xml");
  assertStringIncludes(core, "<dc:title>Welcome</dc:title>");
  for (const absent of ["Un-named", "<dc:creator>", "<cp:lastModifiedBy>"]) notIn(core, absent);
  notIn(await partXml(doc, "docProps/custom.xml"), "<property");
});

test("static properties reach core.xml and custom.xml", async () => {
  const properties = {
    author: "Acme Legal",
    subject: "Renewal",
    keywords: "contract; renewal",
    description: "Draft for review",
    custom: { matter: "M-1042" },
  };
  const t = template(
    <Document id="d" title="Welcome" properties={properties}><Paragraph id="p">Hi</Paragraph></Document>,
  );
  const doc = await buildDocument(t, {});
  assertEquals(doc.properties, properties);
  const core = await partXml(doc, "docProps/core.xml");
  for (
    const part of [
      "<dc:creator>Acme Legal</dc:creator>",
      "<cp:lastModifiedBy>Acme Legal</cp:lastModifiedBy>",
      "<dc:subject>Renewal</dc:subject>",
      "<cp:keywords>contract; renewal</cp:keywords>",
      "<dc:description>Draft for review</dc:description>",
    ]
  ) assertStringIncludes(core, part);
  assertStringIncludes(
    await partXml(doc, "docProps/custom.xml"),
    'pid="2" name="matter"><vt:lpwstr>M-1042</vt:lpwstr></property>',
  );
});

const invoice = template<Data>(
  <Document
    id="d"
    title="Invoice"
    properties={(data: Data) => ({ custom: { invoiceNumber: data.invoice.number } })}
  >
    <Paragraph id="p">Hi</Paragraph>
  </Document>,
);

test("a function of the data builds, publishes a token, and is settled by the resolver", async () => {
  const built = await buildDocument(invoice, { invoice: { number: "INV-7" } });
  assertEquals(built.properties, { custom: { invoiceNumber: "INV-7" } });

  const project = defineDocumentProject<Data>({
    id: "p",
    name: "p",
    template: invoice,
    previewData: { invoice: { number: "INV-7" } },
  });
  const { engineDocument } = await createDocumentProjectArtifact(project, { builtAt: "x" });
  assertEquals(engineDocument.properties?.custom?.invoiceNumber, "{{data.invoice.number}}");

  const resolved: DocumentModel = await resolveDocument(
    engineDocument,
    stateOf({ invoice: { number: "INV-9" } }),
  );
  assertEquals(resolved.properties?.custom?.invoiceNumber, "INV-9");
  assertStringIncludes(await partXml(resolved, "docProps/custom.xml"), "<vt:lpwstr>INV-9</vt:lpwstr>");
});

test("custom names and values are escaped, and empty values are left out", async () => {
  const t = template(
    <Document id="d" title="D" properties={{ custom: { "R&D": "A<B", po: "" } }}>
      <Paragraph id="p">Hi</Paragraph>
    </Document>,
  );
  const custom = await partXml(await buildDocument(t, {}), "docProps/custom.xml");
  assertStringIncludes(custom, 'name="R&amp;D"><vt:lpwstr>A&lt;B</vt:lpwstr>');
  notIn(custom, 'name="po"');
});

test("the title is written when it is not printed", async () => {
  const doc = await buildDocument(
    template(<Document id="d" title="Welcome"><Paragraph id="p">Hi</Paragraph></Document>),
    {},
  );
  const hidden: DocumentModel = { ...doc, style: { ...style, showTitle: false } };
  assertStringIncludes(await partXml(hidden, "docProps/core.xml"), "<dc:title>Welcome</dc:title>");
});
