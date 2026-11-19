import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { cleanMinimalDocumentStyle } from "docxcelerate";
import { documentXml, elementOf, entryOf, partNames, partXml } from "./docx.ts";
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

/** The sheet inside the workbook that chart's "Edit Data" opens. */
async function embeddedSheet(doc: DocumentModel): Promise<string> {
  const blob = await createDocxBlob(doc);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const workbook = entryOf(bytes, "word/embeddings/chart1.xlsx");

  return new TextDecoder().decode(entryOf(workbook, "xl/worksheets/sheet1.xml"));
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
    radar: `<c:radarChart><c:radarStyle val="marker"/>`,
    bubble: `<c:bubbleChart>`,
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

test("a percent stack stacks to a share, and labels its axis as one", async () => {
  const xml = await chartXml(documentOf([chart({ stacked: "percent" })]));

  assertStringIncludes(xml, `<c:grouping val="percentStacked"/>`);
  // The segments still have to close up: a percent stack left at the clustered
  // overlap is the same staircase an ordinary stack would be.
  assertStringIncludes(xml, `<c:overlap val="100"/>`);
  // Unformatted, Word labels the axis 0, 0.2, 0.4 — a scale nobody means by
  // "100% stacked". `sourceLinked="0"` is what stops the numbers' own format
  // taking it back.
  assertStringIncludes(xml, `<c:numFmt formatCode="0%" sourceLinked="0"/>`);
});

test("a document that names its own format outranks the percent default", async () => {
  const xml = await chartXml(
    documentOf([chart({ stacked: "percent", numberFormat: "0.0%" })]),
  );

  assertStringIncludes(xml, `<c:numFmt formatCode="0.0%" sourceLinked="0"/>`);
});

test("a line and an area stack as standard rather than as clustered", async () => {
  const line = await chartXml(documentOf([chart({ graphType: "line" })]));
  const area = await chartXml(documentOf([chart({ graphType: "area", stacked: "percent" })]));

  assertStringIncludes(line, `<c:grouping val="standard"/>`);
  assertStringIncludes(area, `<c:grouping val="percentStacked"/>`);
});

test("a radar rules its spokes as well as its rings", async () => {
  const radar = await chartXml(documentOf([chart({ graphType: "radar" })]));
  const bar = await chartXml(documentOf([chart()]));

  // The web is what makes a ring readable as a ring. A bar's category axis
  // carries no grid of its own — the value axis already rules the plot — so
  // the count is what separates the two.
  assertEquals([...radar.matchAll(/<c:majorGridlines>/g)].length, 2);
  assertEquals([...bar.matchAll(/<c:majorGridlines>/g)].length, 1);
  assertStringIncludes(radar, "<c:catAx>");
  assertStringIncludes(radar, "<c:valAx>");
  // A radar cannot stack: the ring is the scale, and the schema gives its
  // chart group nowhere to say otherwise.
  assertEquals(radar.includes("<c:grouping"), false);
});

test("a bubble carries a third figure, in a column of its own", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        graphType: "bubble",
        data: {
          categories: ["1", "2", "3"],
          series: [{ label: "Sites", values: [4, 9, 6], sizes: [30, 80, 55] }],
        },
      }),
    ]),
  );

  assertEquals(xml.includes("<c:catAx>"), false);
  assertStringIncludes(xml, "<c:xVal>");
  assertStringIncludes(xml, "<c:yVal>");
  assertStringIncludes(xml, "<c:bubbleSize>");
  // The y in B and the size in C, which is what keeps the chart's formulas and
  // the workbook the reader opens naming the same cells.
  assertStringIncludes(xml, `<c:f>Sheet1!$B$2:$B$4</c:f>`);
  assertStringIncludes(xml, `<c:f>Sheet1!$C$2:$C$4</c:f>`);
  assertStringIncludes(xml, `<c:pt idx="1"><c:v>80</c:v></c:pt>`);
});

test("a bubble's sizes reach the workbook Edit Data opens", async () => {
  const doc = documentOf([
    chart({
      graphType: "bubble",
      data: {
        categories: ["1", "2"],
        series: [
          { label: "North", values: [4, 9], sizes: [30, 80] },
          { label: "South", values: [7, 3], sizes: [12, 44] },
        ],
      },
    }),
  ]);
  const sheet = await embeddedSheet(doc);

  // Two columns per series, so the second series starts at D rather than at C.
  assertStringIncludes(sheet, `<c r="B2"><v>4</v></c><c r="C2"><v>30</v></c>`);
  assertStringIncludes(sheet, `<c r="D2"><v>7</v></c><c r="E2"><v>12</v></c>`);
  // And the size column is headed, so a reader who opens the data can tell
  // which run of figures is which.
  assertStringIncludes(sheet, "North (size)");
});

test("a size nobody gave is no bubble, not a bubble of nothing", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        graphType: "bubble",
        data: {
          categories: ["1", "2", "3"],
          series: [{ label: "Sites", values: [4, 9, 6], sizes: [30, null, 55] }],
        },
      }),
    ]),
  );

  const sizes = xml.slice(xml.indexOf("<c:bubbleSize>"));

  assertStringIncludes(sizes, `<c:ptCount val="3"/>`);
  assertEquals(sizes.slice(0, sizes.indexOf("</c:bubbleSize>")).includes(`idx="1"`), false);
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

test("a dashed series is drawn dashed, and a dotted one in round dots", async () => {
  const line = (dash?: "dashed" | "dotted") => ({ label: "x", values: [1, 2, 3, 4], dash });
  const xml = await chartXml(
    documentOf([chart({ graphType: "line", data: { categories: quarters, series: [line(), line("dashed"), line("dotted")] } })]),
  );

  assertStringIncludes(xml, `<a:solidFill><a:srgbClr val="EB6834"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln>`);
  assertStringIncludes(xml, `<a:solidFill><a:srgbClr val="1BAF7A"/></a:solidFill><a:prstDash val="sysDot"/><a:round/></a:ln>`);
  assertEquals(xml.match(/<a:prstDash /g)?.length, 2);
});

test("a series with no dash is solid, exactly as before", async () => {
  const xml = await chartXml(documentOf([chart({ graphType: "line" })]));

  assertEquals(xml.includes("prstDash"), false);
  assertStringIncludes(xml, `<a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="2A78D6"/></a:solidFill><a:round/></a:ln>`);
});

test("scatter and radar take a dash, because they draw their series as lines", async () => {
  for (const graphType of ["scatter", "radar"] as const) {
    const xml = await chartXml(
      documentOf([chart({ graphType, data: { categories: quarters, series: [{ values: [1, 2, 3, 4], dash: "dashed" }] } })]),
    );

    assertStringIncludes(xml, `<a:prstDash val="dash"/>`);
  }
});

test("a bar, an area or a pie has no line to dash", async () => {
  for (const graphType of ["bar", "area", "pie"] as const) {
    const xml = await chartXml(
      documentOf([chart({ graphType, data: { categories: quarters, series: [{ values: [1, 2, 3, 4], dash: "dotted" }] } })]),
    );

    assertEquals(xml.includes("prstDash"), false);
  }
});

test("a dash nobody defined draws solid rather than breaking the file", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        graphType: "line",
        data: {
          categories: quarters,
          series: [
            { values: [1, 2, 3, 4], dash: "dashdot" as never },
            { values: [1, 2, 3, 4], dash: "constructor" as never },
          ],
        },
      }),
    ]),
  );

  assertEquals(xml.includes("prstDash"), false);
  assertEquals(xml.includes("undefined"), false);
});

test("a dash written as JSX reaches the node", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Section id="body" title="Body" showTitle={false}>
          <Graph
            id="revenue"
            graphType="line"
            data={{ categories: quarters, series: [{ label: "Forecast", values: [1, 2, 3, 4], dash: "dashed" }] }}
          />
        </Section>
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
  const section = doc.nodes[0];
  const node = section.kind === "section" ? section.children[0] as GraphNode : undefined;

  assertEquals(node?.data?.series[0].dash, "dashed");
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
  assertStringIncludes(pie, `<c:showCatName val="0"/>`);
  assertEquals(pie.includes("<c:separator>"), false);
  assertEquals(bar.includes("<c:separator>"), false);
});

test("labels are off unless asked for", async () => {
  assertEquals((await chartXml(documentOf([chart()]))).includes("<c:dLbls>"), false);
  assertEquals(
    (await chartXml(documentOf([chart({ dataLabels: false })]))).includes("<c:dLbls>"),
    false,
  );
});

test("a pie can label each slice with its name and share", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        graphType: "pie",
        dataLabels: "categoryAndPercent",
        data: { categories: quarters, series: [twoSeries[0]] },
      }),
    ]),
  );

  assertStringIncludes(xml, `<c:showVal val="0"/>`);
  assertStringIncludes(xml, `<c:showCatName val="1"/>`);
  assertStringIncludes(xml, `<c:showPercent val="1"/>`);
  assertStringIncludes(xml, `<c:showBubbleSize val="0"/><c:separator>, </c:separator></c:dLbls>`);
});

test("a bar can print its category on the bar, with no separator to write", async () => {
  const xml = await chartXml(documentOf([chart({ dataLabels: "category" })]));

  assertStringIncludes(xml, `<c:showVal val="0"/>`);
  assertStringIncludes(xml, `<c:showCatName val="1"/>`);
  assertStringIncludes(xml, `<c:showPercent val="0"/>`);
  assertEquals(xml.includes("<c:separator>"), false);
});

test("a doughnut can label each slice with its name and value", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        graphType: "doughnut",
        dataLabels: "categoryAndValue",
        data: { categories: quarters, series: [twoSeries[0]] },
      }),
    ]),
  );

  assertStringIncludes(xml, `<c:showVal val="1"/>`);
  assertStringIncludes(xml, `<c:showCatName val="1"/>`);
  assertStringIncludes(xml, `<c:showPercent val="0"/>`);
  assertStringIncludes(xml, `<c:separator>, </c:separator>`);
});

test("an explicit value on a pie overrides the share a pie defaults to", async () => {
  const xml = await chartXml(
    documentOf([
      chart({
        graphType: "pie",
        dataLabels: "value",
        data: { categories: quarters, series: [twoSeries[0]] },
      }),
    ]),
  );

  assertStringIncludes(xml, `<c:showVal val="1"/>`);
  assertStringIncludes(xml, `<c:showPercent val="0"/>`);
});

test("label content written as JSX reaches the model", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Section id="body" title="Body" showTitle={false}>
          <Graph
            id="spend"
            graphType="pie"
            dataLabels="categoryAndPercent"
            data={{ categories: quarters, series: [twoSeries[0]] }}
          />
        </Section>
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
  const section = doc.nodes[0];
  const node = section.kind === "section" ? section.children[0] as GraphNode : undefined;

  assertEquals(node?.dataLabels, "categoryAndPercent");
});

test("a percentage label off a pie is an error", async () => {
  for (const [graphType, dataLabels] of [["bar", "percent"], ["line", "categoryAndPercent"]] as const) {
    await assertRejects(
      () =>
        buildDocument(
          template(
            <Document id="doc" title="Doc">
              <Section id="body" title="Body" showTitle={false}>
                <Graph
                  id="r"
                  graphType={graphType}
                  dataLabels={dataLabels}
                  data={{ categories: quarters, series: twoSeries }}
                />
              </Section>
            </Document>,
          ),
          {},
          { branchMode: "decide", dynamicMode: "placeholder" },
        ),
      Error,
      "only a pie or a doughnut",
    );
  }
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

test("the alt outranks the title in the drawing's description, and the title stays on the chart", async () => {
  const doc = documentOf([chart({ title: "Revenue by quarter", alt: "Revenue rose 40% from Q1 to Q4" })]);
  const body = await documentXml(doc);

  assertStringIncludes(body, `descr="Revenue rose 40% from Q1 to Q4"`);
  assertEquals(body.includes(`descr="Revenue by quarter"`), false);
  assertStringIncludes(await chartXml(doc), "Revenue by quarter");
});

test("without an alt the description falls back to the caption, then the id", async () => {
  assertStringIncludes(await documentXml(documentOf([chart({ caption: "Figure 1" })])), `descr="Figure 1"`);
  assertStringIncludes(await documentXml(documentOf([chart()])), `descr="revenue"`);
});

test("an alt is escaped in the attribute", async () => {
  const body = await documentXml(documentOf([chart({ alt: "R&D spend <doubled>" })]));

  assertStringIncludes(body, `descr="R&amp;D spend &lt;doubled&gt;"`);
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
// Trendlines
// ---------------------------------------------------------------------------

const LINEAR =
  '<c:trendline><c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr><c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>';

/** A one-series chart whose series asks for `trendline`. */
function trended(trendline: unknown, over: Partial<GraphNode> = {}): GraphNode {
  return chart({
    graphType: "line",
    data: { categories: quarters, series: [{ label: "A", values: [12, 18, 9, 22], color: "123456", trendline } as never] },
    ...over,
  });
}

test("a linear trendline is packed on a line series, after its points and before its categories", async () => {
  const xml = await chartXml(documentOf([trended("linear")]));

  assertStringIncludes(xml, LINEAR);
  assertEquals(xml.indexOf("</c:marker>") < xml.indexOf("<c:trendline>"), true);
  assertEquals(xml.indexOf("<c:trendline>") < xml.indexOf("<c:cat>"), true);
});

test("a moving average is packed with its period, on a bar", async () => {
  const xml = await chartXml(documentOf([trended({ movingAverage: 3 }, { graphType: "bar" })]));

  assertStringIncludes(xml, '<c:trendlineType val="movingAvg"/><c:period val="3"/>');
  assertEquals(xml.indexOf('<c:invertIfNegative val="0"/>') < xml.indexOf("<c:trendline>"), true);
  assertEquals(xml.indexOf("<c:trendline>") < xml.indexOf("<c:cat>"), true);
});

test("a series with no trendline writes none", async () => {
  assertEquals((await chartXml(documentOf([chart()]))).includes("<c:trendline"), false);
});

test("only the series that asks gets a trendline", async () => {
  const xml = await chartXml(documentOf([chart({
    graphType: "line",
    data: { categories: quarters, series: [twoSeries[0], { ...twoSeries[1], trendline: "linear" }] },
  })]));

  assertEquals([...xml.matchAll(/<c:trendline>/g)].length, 1);
  assertEquals(xml.indexOf("<c:trendline>") > xml.lastIndexOf("<c:ser>"), true);
});

test("a chart Word gives no trendline is written without one", async () => {
  for (
    const over of [
      { graphType: "pie" },
      { graphType: "doughnut" },
      { graphType: "radar" },
      { graphType: "bar", stacked: true },
      { graphType: "line", stacked: "percent" },
    ] as Partial<GraphNode>[]
  ) {
    assertEquals((await chartXml(documentOf([trended("linear", over)]))).includes("<c:trendline"), false);
  }
});

test("a trendline sits before the x values on a scatter and a bubble", async () => {
  const categories = ["1", "2", "3"];
  for (const graphType of ["scatter", "bubble"] as const) {
    const xml = await chartXml(documentOf([chart({
      graphType,
      data: { categories, series: [{ label: "A", values: [3, 1, 2], sizes: [1, 2, 3], trendline: "linear" }] },
    })]));

    assertEquals(xml.indexOf("<c:trendline>") < xml.indexOf("<c:xVal>"), true);
    if (graphType === "bubble") {
      assertStringIncludes(xml, "<c:bubbleSize>");
      assertStringIncludes(xml, '<c:bubble3D val="0"/></c:ser>');
    }
  }
});

test("a moving average's period is rounded, floored at 2, and dropped when it fills the series", async () => {
  const period = (movingAverage: number) => chartXml(documentOf([trended({ movingAverage })]));

  assertStringIncludes(await period(2.6), '<c:period val="3"/>');
  assertStringIncludes(await period(1), '<c:period val="2"/>');
  assertEquals((await period(4)).includes("<c:trendline"), false);
  assertEquals((await period(Number.NaN)).includes("<c:trendline"), false);
});

test("a trendline kind the toolkit does not know writes nothing", async () => {
  assertEquals((await chartXml(documentOf([trended("exp" as never)]))).includes("<c:trendline"), false);
});

test("a trendline adds nothing to the workbook", async () => {
  assertEquals(
    await embeddedSheet(documentOf([trended("linear")])),
    await embeddedSheet(documentOf([trended(undefined)])),
  );
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
            alt="Revenue rose 40%"
            graphType="line"
            stacked
            legend="right"
            numberFormat="£#,##0"
            categoryAxisTitle="Quarter"
            valueAxisTitle="Revenue"
            valueAxisMin={90}
            valueAxisMax={100}
            valueAxisStep={2}
            dataLabels
            dataTable
            axes={false}
            gridlines={false}
            width={320}
            height={200}
            align="right"
            data={{
              categories: quarters,
              series: [
                { ...twoSeries[0], trendline: "linear" },
                { ...twoSeries[1], trendline: { movingAverage: 3 } },
              ],
            }}
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
  assertEquals(node?.alt, "Revenue rose 40%");
  assertEquals(node?.graphType, "line");
  assertEquals(node?.stacked, true);
  assertEquals(node?.legend, "right");
  assertEquals(node?.numberFormat, "£#,##0");
  assertEquals(node?.categoryAxisTitle, "Quarter");
  assertEquals(node?.valueAxisTitle, "Revenue");
  assertEquals(node?.valueAxisMin, 90);
  assertEquals(node?.valueAxisMax, 100);
  assertEquals(node?.valueAxisStep, 2);
  assertEquals(node?.dataLabels, true);
  assertEquals(node?.dataTable, true);
  assertEquals(node?.axes, false);
  assertEquals(node?.gridlines, false);
  assertEquals(node?.width, 320);
  assertEquals(node?.height, 200);
  assertEquals(node?.align, "right");
  assertEquals(node?.data?.series[1].values[3], 27);
  assertEquals(node?.data?.series[0].trendline, "linear");
  assertEquals(node?.data?.series[1].trendline, { movingAverage: 3 });
});

test("a chart that sets neither axes nor gridlines carries no such keys", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Graph id="plain" graphType="line" data={{ categories: quarters, series: twoSeries }} />
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
  const node = doc.nodes[0] as GraphNode;

  assertEquals(node.kind, "graph");
  assertEquals(Object.hasOwn(node, "axes"), false);
  assertEquals(Object.hasOwn(node, "gridlines"), false);
});

// ---------------------------------------------------------------------------
// Axes and gridlines
// ---------------------------------------------------------------------------

const count = (xml: string, tag: string) => [...xml.matchAll(new RegExp(tag, "g"))].length;

test("a chart draws both axes and its grid unless it is told otherwise", async () => {
  const xml = await chartXml(documentOf([chart()]));

  assertEquals(count(xml, '<c:delete val="0"/>'), 2);
  assertEquals(xml.includes('<c:delete val="1"/>'), false);
  assertEquals(count(xml, "<c:majorGridlines>"), 1);
});

test("a chart without axes deletes them rather than dropping them", async () => {
  const xml = await chartXml(documentOf([chart({ axes: false })]));

  assertEquals(count(xml, '<c:delete val="1"/>'), 2);
  assertEquals(xml.includes('<c:delete val="0"/>'), false);
  assertStringIncludes(xml, "<c:catAx>");
  assertStringIncludes(xml, "<c:valAx>");
  assertEquals(count(xml, "<c:majorGridlines>"), 1);
});

test("a chart without gridlines keeps its axes", async () => {
  const xml = await chartXml(documentOf([chart({ gridlines: false })]));

  assertEquals(xml.includes("<c:majorGridlines>"), false);
  assertEquals(count(xml, '<c:delete val="0"/>'), 2);
});

test("a radar without gridlines loses its spokes and its rings both", async () => {
  const xml = await chartXml(documentOf([chart({ graphType: "radar", gridlines: false })]));

  assertEquals(count(xml, "<c:majorGridlines>"), 0);
});

test("a scatter without axes deletes both of its measured axes", async () => {
  const xml = await chartXml(documentOf([chart({ graphType: "scatter", axes: false })]));

  assertEquals(count(xml, '<c:delete val="1"/>'), 2);
});

test("a pie ignores axes and gridlines, having neither", async () => {
  assertEquals(
    await chartXml(documentOf([chart({ graphType: "pie", axes: false, gridlines: false })])),
    await chartXml(documentOf([chart({ graphType: "pie" })])),
  );
});

test("a chart with no dataTable prop carries no dataTable field", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Graph id="g" data={{ categories: quarters, series: twoSeries }} />
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );

  assertEquals("dataTable" in doc.nodes[0], false);
});

// ---------------------------------------------------------------------------
// Data table
// ---------------------------------------------------------------------------

test("a bar asks for a table and gets one, between the axes and the plot area's fill", async () => {
  const xml = await chartXml(documentOf([chart({ dataTable: true })]));

  assertStringIncludes(
    xml,
    '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/>',
  );
  assertStringIncludes(xml, "</c:valAx><c:dTable>");
  assertStringIncludes(xml, '</c:dTable><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>');
  assertEquals(new RegExp(`<c:dTable>[\\s\\S]*<a:srgbClr val="${basePalette.rule}"/>[\\s\\S]*</c:dTable>`).test(xml), true);
  // The rows carry the keys, so the legend the two series would get is dropped.
  assertEquals(xml.includes("<c:legend>"), false);
});

test("an explicit legend is kept beside the table", async () => {
  const xml = await chartXml(documentOf([chart({ dataTable: true, legend: "bottom" })]));

  assertStringIncludes(xml, '<c:legendPos val="b"/>');
  assertStringIncludes(xml, "<c:dTable>");
});

test("every chart Word tables takes the table", async () => {
  for (const graphType of ["bar", "barHorizontal", "line", "area"] as const) {
    assertStringIncludes(await chartXml(documentOf([chart({ graphType, dataTable: true })])), "<c:dTable>");
  }
});

test("the other types ignore the table and keep their defaults", async () => {
  for (const graphType of ["pie", "doughnut", "scatter", "radar", "bubble"] as const) {
    const series = graphType === "bubble" ? { ...twoSeries[0], sizes: [1, 2, 3, 4] } : twoSeries[0];
    const single = graphType === "pie" || graphType === "doughnut";
    const xml = await chartXml(documentOf([
      chart({
        graphType,
        dataTable: true,
        data: { categories: quarters, series: single || graphType === "bubble" ? [series] : twoSeries },
      }),
    ]));

    assertEquals(xml.includes("<c:dTable"), false);
    if (graphType === "pie") {
      assertStringIncludes(xml, '<c:legendPos val="b"/>');
    }
  }
});

test("a chart that does not ask for a table is written as it always was", async () => {
  const off = await chartXml(documentOf([chart()]));

  assertEquals(off.includes("<c:dTable"), false);
  assertEquals(off, await chartXml(documentOf([chart({ dataTable: false })])));
});

// ---------------------------------------------------------------------------
// Where the chart sits
// ---------------------------------------------------------------------------

test("a chart ranged right takes its caption with it", async () => {
  const xml = await documentXml(documentOf([chart({ width: 240, align: "right", caption: "Figure 1: revenue." })]));

  assertStringIncludes(elementOf(xml, "w:p", "<c:chart"), '<w:jc w:val="right"/>');
  assertStringIncludes(elementOf(xml, "w:p", ">Figure 1: revenue.</w:t>"), '<w:jc w:val="right"/>');
});

test("a chart with no align writes none, and its caption stays centred", async () => {
  const xml = await documentXml(documentOf([chart({ caption: "Figure 1: revenue." })]));

  assertEquals(elementOf(xml, "w:p", "<c:chart").includes("<w:jc"), false);
  assertStringIncludes(elementOf(xml, "w:p", ">Figure 1: revenue.</w:t>"), '<w:jc w:val="center"/>');
});

test("a chart's placeholder sits where the chart will", async () => {
  const xml = await documentXml(documentOf([
    { id: "trend", kind: "graph", mode: "dynamic", graphType: "line", placeholder: "Balance over the term", align: "center" },
  ]));

  assertStringIncludes(elementOf(xml, "w:p", "[line chart: Balance over the term]"), '<w:jc w:val="center"/>');
});

test("a chart written without axis bounds carries none, not undefined ones", async () => {
  const doc = await buildDocument(
    template(
      <Document id="doc" title="Doc">
        <Graph id="plain" data={{ categories: quarters, series: twoSeries }} />
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
  const node = doc.nodes[0] as GraphNode;

  assertEquals("valueAxisMin" in node, false);
});

// ---------------------------------------------------------------------------
// Value axis bounds
// ---------------------------------------------------------------------------

test("a chart with no axis bounds leaves the value axis to Word", async () => {
  const xml = await chartXml(documentOf([chart()]));

  for (const absent of ["<c:min ", "<c:max ", "<c:majorUnit"]) {
    assertEquals(xml.includes(absent), false);
  }
  assertStringIncludes(
    xml,
    `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>`,
  );
});

test("axis bounds are written to the value axis, max before min, and not the category axis", async () => {
  const xml = await chartXml(
    documentOf([chart({ valueAxisMin: 90, valueAxisMax: 100, valueAxisStep: 2 })]),
  );

  assertStringIncludes(
    xml,
    `<c:scaling><c:orientation val="minMax"/><c:max val="100"/><c:min val="90"/></c:scaling>`,
  );
  assertStringIncludes(xml, `<c:crossBetween val="between"/><c:majorUnit val="2"/></c:valAx>`);
  assertEquals(xml.split("<c:min ").length, 2);
});

test("zero is a bound, and one end can be given alone", async () => {
  const xml = await chartXml(documentOf([chart({ valueAxisMin: 0 })]));

  assertStringIncludes(xml, `<c:orientation val="minMax"/><c:min val="0"/></c:scaling>`);
  assertEquals(xml.includes("<c:max "), false);
  assertEquals(xml.includes("<c:majorUnit"), false);
});

test("a scatter's bounds set its y axis and leave x alone", async () => {
  const xml = await chartXml(documentOf([chart({ graphType: "scatter", valueAxisMin: 5 })]));
  const axes = xml.split("<c:valAx>");

  assertEquals(axes[1].includes("<c:min "), false);
  assertStringIncludes(axes[2], `<c:min val="5"/>`);
});

test("a pie ignores axis bounds, even invalid ones", async () => {
  const xml = await chartXml(documentOf([
    chart({
      graphType: "pie",
      data: { categories: quarters, series: [twoSeries[0]] },
      valueAxisMin: 5,
      valueAxisStep: -1,
    }),
  ]));

  assertEquals(xml.includes("<c:min "), false);
  assertEquals(xml.includes("<c:majorUnit"), false);
});

test("axis bounds Word has no drawing for stop the build", async () => {
  const message = (over: Partial<GraphNode>, text: string) =>
    assertRejects(() => chartXml(documentOf([chart(over)])), Error, text);

  await message({ valueAxisMin: 100, valueAxisMax: 100 }, "must be below valueAxisMax");
  await message({ valueAxisStep: 0 }, "greater than zero");
  await message({ valueAxisMax: Number.NaN }, "finite number");
});

// ---------------------------------------------------------------------------
// Text size
// ---------------------------------------------------------------------------

function sizedBy(fontSizePt: number): DocumentStyle {
  return {
    ...cleanMinimalDocumentStyle,
    blocks: { ...cleanMinimalDocumentStyle.blocks, chart: { fontSizePt } },
  };
}

const titled = () => chart({ title: "Revenue", categoryAxisTitle: "Quarter", valueAxisTitle: "£k" });

test("a theme with no chart block letters its charts at 9pt, titled at 12pt", async () => {
  const xml = await chartXml(documentOf([chart({ title: "Revenue", valueAxisTitle: "£k" })]));

  assertStringIncludes(xml, '<a:defRPr sz="900" b="0" i="0"');
  assertStringIncludes(xml, '<a:defRPr sz="1200" b="0">');
});

test("a theme's chart block sizes every piece of a chart's text, and its title follows", async () => {
  const xml = await chartXml(documentOf([titled()], sizedBy(14)));

  assertEquals(xml.split('sz="1400" b="0" i="0"').length - 1, 4);
  assertEquals(xml.split('<a:defRPr sz="1700" b="0">').length - 1, 1);
  assertEquals(xml.split('<a:defRPr sz="1400" b="0"><a:solidFill>').length - 1, 2);
  assertEquals(xml.includes('sz="900"') || xml.includes('sz="1200"'), false);
});

test("a fractional size survives as hundredths of a point", async () => {
  const xml = await chartXml(documentOf([titled()], sizedBy(7.5)));

  assertStringIncludes(xml, 'sz="750"');
  assertStringIncludes(xml, 'sz="1050"');
});

test("a size Word would refuse is held to the range it accepts", async () => {
  const huge = await chartXml(documentOf([titled()], sizedBy(5000)));

  assertStringIncludes(huge, 'sz="400000"');
  assertEquals(huge.includes('sz="5000'), false);

  for (const bad of [-2, Number.NaN]) {
    const xml = await chartXml(documentOf([titled()], sizedBy(bad)));

    assertStringIncludes(xml, 'sz="900"');
    assertStringIncludes(xml, 'sz="1200"');
  }
});

test("other blocks do not size a chart", async () => {
  const style: DocumentStyle = {
    ...cleanMinimalDocumentStyle,
    blocks: { ...cleanMinimalDocumentStyle.blocks, chartCaption: { fontSizePt: 20 } },
  };
  const xml = await chartXml(documentOf([chart({ title: "Revenue", caption: "Figure 1" })], style));

  assertStringIncludes(xml, 'sz="900"');
  assertEquals(xml.includes('sz="2000"'), false);
});

// ---------------------------------------------------------------------------
// Colouring single points
// ---------------------------------------------------------------------------

function coloured(graphType: GraphNode["graphType"], pointColors: (string | null)[], cats = quarters) {
  return chartXml(documentOf([
    chart({
      graphType,
      data: { categories: cats, series: [{ label: "n", values: cats.map((_, at) => at + 1), pointColors }] },
    }),
  ]));
}

test("a bar picks out one point and leaves the series colour alone", async () => {
  const xml = await coloured("bar", [null, null, "C0392B"]);

  assertStringIncludes(
    xml,
    '<c:dPt><c:idx val="2"/><c:invertIfNegative val="0"/><c:spPr><a:solidFill><a:srgbClr val="C0392B"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt>',
  );
  assertEquals(xml.split("<c:dPt>").length - 1, 1);
  assertStringIncludes(xml, '<c:spPr><a:solidFill><a:srgbClr val="2A78D6"/>');
  assertEquals(xml.indexOf("<c:dPt>") > xml.indexOf("<c:invertIfNegative") && xml.indexOf("<c:dPt>") < xml.indexOf("<c:cat>"), true);
});

test("a chart with no point colours writes no data points", async () => {
  assertEquals((await chartXml(documentOf([chart()]))).includes("<c:dPt>"), false);
});

test("a pie slice is overridden and the rest keep the palette", async () => {
  const xml = await coloured("pie", [null, "C0392B"]);

  assertEquals(xml.split("<c:dPt>").length - 1, 4);
  assertStringIncludes(xml, '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="C0392B"/>');
  assertStringIncludes(xml, '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="2A78D6"/>');
  assertStringIncludes(xml, '<c:dPt><c:idx val="2"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="1BAF7A"/>');
});

test("a line paints the marker of a point, not the stroke through it", async () => {
  const xml = await coloured("line", [null, "C0392B"]);

  assertStringIncludes(
    xml,
    '<c:dPt><c:idx val="1"/><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="C0392B"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker></c:dPt>',
  );
});

test("an area has no points to paint", async () => {
  assertEquals((await coloured("area", ["C0392B"])).includes("<c:dPt>"), false);
});

test("point colours past the categories are dropped", async () => {
  assertEquals((await coloured("bar", [null, null, "C0392B"], ["a", "b"])).includes("<c:dPt>"), false);
});

test("pointColors reaches the node from JSX", async () => {
  const doc = await buildDocument(
    template(
      <Document id="d" title="t">
        <Graph id="g" graphType="bar" data={{ series: [{ values: [1, 2], pointColors: [null, "C0392B"] }] }} />
      </Document>,
    ),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
  const node = doc.nodes[0] as GraphNode;

  assertEquals(node.data?.series[0].pointColors, [null, "C0392B"]);
});
