import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { cleanMinimalDocumentStyle } from "docxcelerate";
import { documentXml, entryOf, partNames, partXml } from "./docx.ts";
import { createDocxBlob } from "docxcelerate/docx";
import type { DocumentModel, DocumentStyle, GraphNode } from "docxcelerate";
import { Cell, Document, Graph, Row, Section, Table, template } from "docxcelerate/template";

/**
 * Charts, packed as charts.
 *
 * The thing worth testing is not that something appears where a chart was
 * asked for — a paragraph saying "[bar chart]" would do that. It is that the
 * package holds a real `c:chartSpace`, related from the document and declared
 * in the content types, with every value cached in it and a workbook behind it.
 * That is the difference between a chart a reader can select, restyle and open
 * the data of, and a picture; and each of the four parts is one a change could
 * quietly drop while the other three still looked right.
 *
 * The numbers are asserted from the packed XML rather than from the model,
 * because the model is not what Word reads.
 */

/** The preset's own palette, which every chart here is themed by. */
const basePalette = cleanMinimalDocumentStyle.palette ?? {
  heading: "111827",
  accent: "2F5FBD",
  muted: "6B7280",
  rule: "D1D5DB",
  page: "FFFFFF",
};

const quarters = ["Q1", "Q2", "Q3", "Q4"];
const twoSeries = [
  { label: "2024", values: [12, 18, 9, 22] },
  { label: "2025", values: [16, 14, 21, 27] },
];

function documentOf(nodes: GraphNode[], style?: DocumentStyle): DocumentModel {
  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: "charts",
    title: "Charts",
    style,
    nodes,
  };
}

function chart(over: Partial<GraphNode> = {}): GraphNode {
  return {
    id: "revenue",
    kind: "graph",
    mode: "static",
    graphType: "bar",
    data: { categories: quarters, series: twoSeries },
    ...over,
  } as GraphNode;
}

/** The one chart part a document holds. */
function chartXml(doc: DocumentModel): Promise<string> {
  return partXml(doc, "word/charts/chart1.xml");
}

// ---------------------------------------------------------------------------
// The package
// ---------------------------------------------------------------------------

test("a chart is a part of the package, not a picture in the body", async () => {
  const names = await partNames(documentOf([chart()]));

  assertEquals(names.includes("word/charts/chart1.xml"), true);
  assertEquals(names.includes("word/charts/_rels/chart1.xml.rels"), true);
  assertEquals(names.includes("word/embeddings/chart1.xlsx"), true);
});

test("a document with no chart is packed byte for byte as it was", async () => {
  const plain: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "plain",
    title: "Plain",
    nodes: [{ id: "hello", kind: "paragraph", mode: "static", text: "Hello." }],
  };
  const names = await partNames(plain);

  assertEquals(names.some((name) => name.startsWith("word/charts/")), false);
  assertEquals(names.some((name) => name.startsWith("word/embeddings/")), false);
});

test("the chart part declares its content type, and so does its workbook", async () => {
  const types = await partXml(documentOf([chart()]), "[Content_Types].xml");

  assertStringIncludes(
    types,
    `PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"`,
  );
  assertStringIncludes(
    types,
    `PartName="/word/embeddings/chart1.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"`,
  );
});

test("the drawing points at a relationship the document actually declares", async () => {
  const doc = documentOf([chart()]);
  const body = await documentXml(doc);
  const rels = await partXml(doc, "word/_rels/document.xml.rels");
  const id = /<c:chart [^>]*r:id="([^"]+)"/.exec(body)?.[1];

  // The token the renderer wrote must be gone: a relationship id that never
  // became one is a drawing pointing at nothing.
  assertEquals(body.includes("dxclChart_"), false);
  assertEquals(id?.startsWith("rId"), true);
  assertStringIncludes(
    rels,
    `Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"`,
  );
});

test("a chart's relationship does not take an id the packer already used", async () => {
  const doc = documentOf([chart()]);
  const rels = await partXml(doc, "word/_rels/document.xml.rels");
  const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((match) => match[1]);

  assertEquals(ids.length, new Set(ids).size);
});

test("several charts each get their own part, workbook and relationship", async () => {
  const doc = documentOf([
    chart({ id: "one" }),
    chart({ id: "two", graphType: "line" }),
    chart({ id: "three", graphType: "pie" }),
  ]);
  const names = await partNames(doc);
  const rels = await partXml(doc, "word/_rels/document.xml.rels");

  assertEquals(
    ["chart1", "chart2", "chart3"].every((name) =>
      names.includes(`word/charts/${name}.xml`) &&
      names.includes(`word/embeddings/${name}.xlsx`)
    ),
    true,
  );
  assertEquals([...rels.matchAll(/relationships\/chart"/g)].length, 3);
});

test("a chart inside a table cell is still a chart", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Table id="grid" columns={[{ width: "auto" }]}>
          <Row>
            <Cell>
              <Graph
                id="inset"
                data={{ categories: ["a"], series: [{ label: "n", values: [1] }] }}
              />
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );

  assertEquals((await partNames(doc)).includes("word/charts/chart1.xml"), true);
});

test("a chart in a running strip relates from that strip, not from the body", async () => {
  const doc: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "furniture",
    title: "Furniture",
    header: [chart({ id: "strip" })],
    nodes: [{ id: "body", kind: "paragraph", mode: "static", text: "Body." }],
  };
  const names = await partNames(doc);
  const header = names.find((name) => /^word\/header\d+\.xml$/.test(name));

  assertEquals(header !== undefined, true);
  assertEquals(names.includes(`word/_rels/${header?.slice("word/".length)}.rels`), true);
  assertStringIncludes(
    await partXml(doc, `word/_rels/${header?.slice("word/".length)}.rels`),
    "charts/chart1.xml",
  );
});

// ---------------------------------------------------------------------------
// What the chart says
// ---------------------------------------------------------------------------

test("every value is cached in the chart, so Word draws without opening anything", async () => {
  const xml = await chartXml(documentOf([chart()]));

  for (const value of [12, 18, 9, 22, 16, 14, 21, 27]) {
    assertStringIncludes(xml, `<c:v>${value}</c:v>`);
  }
  for (const category of quarters) {
    assertStringIncludes(xml, `<c:v>${category}</c:v>`);
  }
  assertStringIncludes(xml, "<c:v>2024</c:v>");
  assertStringIncludes(xml, "<c:v>2025</c:v>");
});

test("the same numbers are in the workbook, which is what Edit Data opens", async () => {
  const blob = await createDocxBlob(documentOf([chart()]));
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const workbook = entryOf(bytes, "word/embeddings/chart1.xlsx");
  const sheet = new TextDecoder().decode(entryOf(workbook, "xl/worksheets/sheet1.xml"));

  assertStringIncludes(sheet, `<c r="A2" t="inlineStr"><is><t xml:space="preserve">Q1</t></is></c>`);
  assertStringIncludes(sheet, `<c r="B2"><v>12</v></c>`);
  assertStringIncludes(sheet, `<c r="C5"><v>27</v></c>`);
  assertStringIncludes(sheet, `<c r="B1" t="inlineStr"><is><t xml:space="preserve">2024</t></is></c>`);
});

test("the chart's formulas name the cells the workbook writes them in", async () => {
  const xml = await chartXml(documentOf([chart()]));

  assertStringIncludes(xml, "<c:f>Sheet1!$A$2:$A$5</c:f>");
  assertStringIncludes(xml, "<c:f>Sheet1!$B$2:$B$5</c:f>");
  assertStringIncludes(xml, "<c:f>Sheet1!$C$1</c:f>");
});

test("each chart type packs as the plot Word draws for it", async () => {
  const expected: Record<string, string> = {
    bar: `<c:barChart><c:barDir val="col"/>`,
    barHorizontal: `<c:barChart><c:barDir val="bar"/>`,
    line: `<c:lineChart>`,
    area: `<c:areaChart>`,
    pie: `<c:pieChart>`,
    doughnut: `<c:doughnutChart>`,
    scatter: `<c:scatterChart>`,
  };

  for (const [graphType, element] of Object.entries(expected)) {
    const xml = await chartXml(documentOf([chart({ graphType: graphType as GraphNode["graphType"] })]));

    assertStringIncludes(xml, element);
  }
});

test("stacking is a property of the bars, not a chart type of its own", async () => {
  const stacked = await chartXml(documentOf([chart({ stacked: true })]));
  const clustered = await chartXml(documentOf([chart()]));

  assertStringIncludes(stacked, `<c:grouping val="stacked"/>`);
  // Overlap is what closes the gap between a stack's segments. Left at the
  // clustered value a stacked bar draws as a staircase, which is the failure
  // this asserts against.
  assertStringIncludes(stacked, `<c:overlap val="100"/>`);
  assertStringIncludes(clustered, `<c:grouping val="clustered"/>`);
  assertStringIncludes(clustered, `<c:overlap val="-27"/>`);
});

test("a horizontal bar reads in the order its categories were written", async () => {
  const horizontal = await chartXml(documentOf([chart({ graphType: "barHorizontal" })]));
  const standing = await chartXml(documentOf([chart()]));

  // OOXML's default puts the first category at the bottom, which is what Word
  // does when you insert a bar chart from a table — and not what a document
  // that handed over an array means by it. Reversing the scale is half of it;
  // the value axis has to cross at the far end too, or the category labels go
  // over to the right-hand side with it.
  assertStringIncludes(horizontal, `<c:orientation val="maxMin"/>`);
  assertStringIncludes(horizontal, `<c:crosses val="max"/>`);
  assertEquals(standing.includes(`<c:orientation val="maxMin"/>`), false);
  assertStringIncludes(standing, `<c:crosses val="autoZero"/>`);
});

test("a pie has no axes, and a bar has both", async () => {
  const pie = await chartXml(documentOf([chart({ graphType: "pie" })]));
  const bar = await chartXml(documentOf([chart()]));

  assertEquals(pie.includes("<c:catAx>"), false);
  assertEquals(pie.includes("<c:valAx>"), false);
  assertStringIncludes(bar, "<c:catAx>");
  assertStringIncludes(bar, "<c:valAx>");
});

test("a scatter measures along both axes rather than naming one", async () => {
  const xml = await chartXml(documentOf([chart({ graphType: "scatter" })]));

  assertEquals(xml.includes("<c:catAx>"), false);
  assertEquals([...xml.matchAll(/<c:valAx>/g)].length, 2);
  assertStringIncludes(xml, "<c:xVal>");
  assertStringIncludes(xml, "<c:yVal>");
});

test("a gap in a series is a gap, not a zero", async () => {
  const xml = await chartXml(
    documentOf([
      chart({ data: { categories: quarters, series: [{ label: "n", values: [1, null, 3, 4] }] } }),
    ]),
  );
  const cache = /<c:numCache>[\s\S]*?<\/c:numCache>/.exec(xml)?.[0] ?? "";

  // The count still names every row, so Word knows the second reading is
  // missing rather than that the series is three long.
  assertStringIncludes(cache, `<c:ptCount val="4"/>`);
  assertEquals(cache.includes(`<c:pt idx="1">`), false);
  assertStringIncludes(xml, `<c:dispBlanksAs val="gap"/>`);
});

test("a series shorter than its categories is squared up rather than left ragged", async () => {
  const xml = await chartXml(
    documentOf([chart({ data: { categories: quarters, series: [{ label: "n", values: [1, 2] }] } })]),
  );

  assertStringIncludes(xml, `<c:ptCount val="4"/>`);
  assertStringIncludes(xml, "<c:f>Sheet1!$B$2:$B$5</c:f>");
});

test("a chart with no categories counts its values by position", async () => {
  const xml = await chartXml(
    documentOf([chart({ data: { series: [{ label: "n", values: [5, 6, 7] }] } })]),
  );

  assertStringIncludes(xml, "<c:v>1</c:v>");
  assertStringIncludes(xml, "<c:v>3</c:v>");
  assertStringIncludes(xml, `<c:ptCount val="3"/>`);
});

// ---------------------------------------------------------------------------
// What the theme decides
// ---------------------------------------------------------------------------

test("series take the theme's chart palette in order", async () => {
  const style: DocumentStyle = {
    ...cleanMinimalDocumentStyle,
    palette: { ...basePalette, series: ["112233", "445566"] },
  };
  const xml = await chartXml(documentOf([chart()], style));

  assertStringIncludes(xml, `<a:srgbClr val="112233"/>`);
  assertStringIncludes(xml, `<a:srgbClr val="445566"/>`);
});

test("a theme that names no chart palette still draws a chart", async () => {
  const xml = await chartXml(documentOf([chart()]));

  assertStringIncludes(xml, `<a:srgbClr val="2A78D6"/>`);
  assertStringIncludes(xml, `<a:srgbClr val="EB6834"/>`);
});

test("a series that names its own colour outranks the theme", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        data: { categories: ["a"], series: [{ label: "n", values: [1], color: "AA00BB" }] },
      }),
    ]),
  );

  assertStringIncludes(xml, `<a:srgbClr val="AA00BB"/>`);
});

test("a chart is ruled and lettered in the theme's own ink, not the series' colours", async () => {
  const xml = await chartXml(documentOf([chart()]));
  const palette = cleanMinimalDocumentStyle.palette;

  assertStringIncludes(xml, `<a:srgbClr val="${palette?.rule}"/>`);
  assertStringIncludes(xml, `<a:srgbClr val="${palette?.muted}"/>`);
  assertStringIncludes(xml, `<a:latin typeface="${cleanMinimalDocumentStyle.typography.bodyFont}"/>`);
});

// ---------------------------------------------------------------------------
// What the document decides
// ---------------------------------------------------------------------------

test("a chart's title is drawn on the chart, and no title deletes Word's own", async () => {
  const titled = await chartXml(documentOf([chart({ title: "Revenue by quarter" })]));
  const untitled = await chartXml(documentOf([chart()]));

  assertStringIncludes(titled, "<a:t>Revenue by quarter</a:t>");
  assertStringIncludes(titled, `<c:autoTitleDeleted val="0"/>`);
  assertStringIncludes(untitled, `<c:autoTitleDeleted val="1"/>`);
});

test("a key appears for two series and not for one, and a pie keeps its own", async () => {
  const many = await chartXml(documentOf([chart()]));
  const one = await chartXml(
    documentOf([chart({ data: { categories: quarters, series: [twoSeries[0]] } })]),
  );
  const pie = await chartXml(
    documentOf([
      chart({ graphType: "pie", data: { categories: quarters, series: [twoSeries[0]] } }),
    ]),
  );

  assertStringIncludes(many, `<c:legendPos val="b"/>`);
  assertEquals(one.includes("<c:legend>"), false);
  assertStringIncludes(pie, `<c:legendPos val="b"/>`);
});

test("a document that asks for no key gets none, whatever it is plotting", async () => {
  const xml = await chartXml(documentOf([chart({ legend: "none" })]));

  assertEquals(xml.includes("<c:legend>"), false);
});

test("a number format reaches the value axis and stops being linked to the source", async () => {
  const xml = await chartXml(documentOf([chart({ numberFormat: "#,##0" })]));

  assertStringIncludes(xml, `<c:numFmt formatCode="#,##0" sourceLinked="0"/>`);
});

test("axis titles are printed, and the vertical one is turned on its side", async () => {
  const xml = await chartXml(
    documentOf([chart({ categoryAxisTitle: "Quarter", valueAxisTitle: "Revenue" })]),
  );

  assertStringIncludes(xml, "<a:t>Quarter</a:t>");
  assertStringIncludes(xml, "<a:t>Revenue</a:t>");
  assertStringIncludes(xml, `rot="-5400000"`);
});

test("a pie's labels show the share, and a bar's show the figure", async () => {
  const pie = await chartXml(
    documentOf([
      chart({
        graphType: "pie",
        dataLabels: true,
        data: { categories: quarters, series: [twoSeries[0]] },
      }),
    ]),
  );
  const bar = await chartXml(documentOf([chart({ dataLabels: true })]));

  assertStringIncludes(pie, `<c:showPercent val="1"/>`);
  assertStringIncludes(bar, `<c:showVal val="1"/>`);
  assertStringIncludes(bar, `<c:showPercent val="0"/>`);
});

test("a label with an ampersand in it is still a label", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        title: "Profit & loss",
        data: { categories: ["R&D"], series: [{ label: "Spend & save", values: [1] }] },
      }),
    ]),
  );

  assertStringIncludes(xml, "<a:t>Profit &amp; loss</a:t>");
  assertStringIncludes(xml, "<c:v>R&amp;D</c:v>");
  assertStringIncludes(xml, "<c:v>Spend &amp; save</c:v>");
});

// ---------------------------------------------------------------------------
// The frame the chart is drawn in
// ---------------------------------------------------------------------------

test("a chart with no size given fills the text column", async () => {
  const body = await documentXml(documentOf([chart()]));
  const extent = /<wp:extent cx="(\d+)" cy="(\d+)"\/>/.exec(body);
  // A4 less 25.4mm margins either side, in points, in EMU.
  const columnPt = (210 - 25.4 * 2) * (72 / 25.4);

  assertEquals(Number(extent?.[1]), Math.round(columnPt * 12700));
  assertEquals(Number(extent?.[2]), Math.round(columnPt * (7 / 12) * 12700));
});

test("a chart that states its size is drawn at it", async () => {
  const body = await documentXml(documentOf([chart({ width: 300, height: 150 })]));

  assertStringIncludes(body, `<wp:extent cx="${300 * 12700}" cy="${150 * 12700}"/>`);
});

test("a chart says what it is for a reader who cannot see it", async () => {
  const body = await documentXml(documentOf([chart({ title: "Revenue by quarter" })]));

  assertStringIncludes(body, `descr="Revenue by quarter"`);
});

test("a caption is a paragraph under the chart, not text drawn into it", async () => {
  const body = await documentXml(documentOf([chart({ caption: "Figure 1: revenue." })]));

  assertStringIncludes(body, "Figure 1: revenue.");
  const chartXml = await partXml(documentOf([chart({ caption: "Figure 1: revenue." })]), "word/charts/chart1.xml");
  assertEquals(chartXml.includes("Figure 1: revenue."), false);
});

test("a chart with nothing to plot stands in for itself rather than packing an empty frame", async () => {
  const doc = documentOf([
    {
      id: "trend",
      kind: "graph",
      mode: "dynamic",
      graphType: "line",
      placeholder: "Balance over the term",
    },
  ]);

  assertEquals((await partNames(doc)).some((name) => name.startsWith("word/charts/")), false);
  assertStringIncludes(await documentXml(doc), "[line chart: Balance over the term]");
});

test("a chart whose series came back empty is a chart with nothing to plot", async () => {
  const doc = documentOf([chart({ data: { categories: [], series: [] }, placeholder: "Nothing yet" })]);

  assertEquals((await partNames(doc)).some((name) => name.startsWith("word/charts/")), false);
});

// ---------------------------------------------------------------------------
// Through the template
// ---------------------------------------------------------------------------

test("a chart written as JSX carries every drawing decision into the node", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Section id="body" title="Body" showTitle={false}>
          <Graph
            id="revenue"
            title="Revenue by quarter"
            graphType="line"
            stacked
            legend="right"
            numberFormat="£#,##0"
            categoryAxisTitle="Quarter"
            valueAxisTitle="Revenue"
            dataLabels
            width={320}
            height={200}
            data={{ categories: quarters, series: twoSeries }}
          />
        </Section>
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );

  const section = doc.nodes[0];
  const node = section.kind === "section" ? section.children[0] as GraphNode : undefined;

  assertEquals(node?.kind, "graph");
  assertEquals(node?.title, "Revenue by quarter");
  assertEquals(node?.graphType, "line");
  assertEquals(node?.stacked, true);
  assertEquals(node?.legend, "right");
  assertEquals(node?.numberFormat, "£#,##0");
  assertEquals(node?.categoryAxisTitle, "Quarter");
  assertEquals(node?.valueAxisTitle, "Revenue");
  assertEquals(node?.dataLabels, true);
  assertEquals(node?.width, 320);
  assertEquals(node?.height, 200);
  assertEquals(node?.data?.series[1].values[3], 27);
});
