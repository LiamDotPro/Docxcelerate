import { test } from "node:test";
import { JSDOM } from "jsdom";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { documentXml, elementOf, partXml } from "./docx.ts";
import type { DocumentModel, DocumentStyle } from "docxcelerate";
import { Cell, Document, Graph, Image, PageNumber, Paragraph, Row, Table, template, useAi } from "docxcelerate/template";

/**
 * The layout properties a designed, many-page document leans on: a picture
 * pinned behind the page, a cell's paragraphs keeping their own spacing, a
 * proof that marks what the engine has still to write, and page numbers that
 * say "PAGE 3 OF 13" in the face of the footer around them.
 *
 * Asserted on the packed XML, because that is what Word is handed.
 */

const style: DocumentStyle = {
  preset: "test",
  page: { size: "A4", orientation: "portrait", margins: { topMm: 16, rightMm: 16, bottomMm: 16, leftMm: 16 } },
  typography: { bodyFont: "Arial", headingFont: "Georgia", bodySizePt: 10, bodyLineHeight: 1.4, color: "1F2933" },
  paragraph: { spacingAfterPt: 6 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 10 },
  sectionHeading: { fontSizePt: 8, weight: "bold", spacingBeforePt: 10, spacingAfterPt: 4 },
  blocks: {
    band: { fill: "0E3B2D", color: "FFFFFF", paddingSidesPt: { top: 7, bottom: 7, left: 0, right: 0 } },
    inset: { paddingSidesPt: { left: 45 } },
    label: { fontSizePt: 6.5, spacingAfterPt: 3 },
    figure: { fontSizePt: 24, spacingBeforePt: 4 },
    quiet: { fontSizePt: 8 },
    dynamicLabel: { fontSizePt: 6.5, color: "5C6B64", spacingAfterPt: 2 },
    standIn: { color: "5C6B64" },
    word: { fontSizePt: 15, standInLabel: false },
    stamp: { fill: "17624A", standIn: "stampPending" },
    stampPending: { fill: "FFFFFF", borderStyle: "dashed", border: "C8D3CE", borderSides: ["top"] },
    segment: { fill: "17624A", fontSizePt: 1, lineHeight: 1, paddingPt: 0 },
    tracked: { fontSizePt: 6.5, weight: "bold", letterSpacingEm: 0.12, color: "5C6B64" },
    letter: { spacingAfterPt: 8 },
  },
};

const PIXEL = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function build(node: unknown, dynamicMode: "placeholder" | "resolve" = "placeholder"): Promise<DocumentModel> {
  return buildDocument(template<Record<string, never>>(node as never), {}, {
    branchMode: "decide",
    dynamicMode,
    aiClient: { generateParagraph: () => "One.\n\nTwo.\n\nThree." },
  }).then((doc) => ({ ...doc, style }));
}

/** A paragraph the engine writes, for the tests that need one. */
const Written = ({ id, variant }: { id: string; variant?: string }) => {
  useAi({ ask: "Write it.", placeholder: "What goes here." });
  return <Paragraph id={id} variant={variant} />;
};

test("a floating picture is anchored from the page corner, behind the text", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">
          <Image id="photo" src={PIXEL} width={100} height={50} float={{ xMm: 0, yMm: 10, behind: true }} />
        </Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(xml, "<wp:anchor");
  assertStringIncludes(xml, 'behindDoc="1"');
  assertStringIncludes(xml, '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset>');
  assertStringIncludes(xml, '<wp:positionV relativeFrom="page"><wp:posOffset>360000</wp:posOffset>');
  assertStringIncludes(xml, "<wp:wrapNone/>");
});

test("a picture that does not float stays in the line", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">
          <Image id="mark" src={PIXEL} width={10} height={10} />
        </Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(xml, "<wp:inline");
  assertEquals(xml.includes("<wp:anchor"), false);
});

test("a paragraph in a cell keeps the space its own variant leaves", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell>
              <Paragraph id="l" variant="label">Value</Paragraph>
              <Paragraph id="f" variant="figure">£412,860</Paragraph>
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(elementOf(xml, "w:p", "Value"), 'w:after="60"');
  assertStringIncludes(elementOf(xml, "w:p", "£412,860"), 'w:before="80"');
});

test("a cell naming one side of its padding keeps the other sides its table set", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" variant="band" columns={[{ width: "auto" }]}>
          <Row>
            <Cell variant="inset">Harbourside</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  const cell = elementOf(xml, "w:tbl", "Harbourside");
  assertStringIncludes(cell, '<w:top w:type="dxa" w:w="140"/>');
  assertStringIncludes(cell, '<w:left w:type="dxa" w:w="900"/>');
});

test("a paragraph variant that names only a size keeps its cell's ink", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" variant="band" columns={[{ width: "auto" }]}>
          <Row>
            <Cell>
              <Paragraph id="q" variant="quiet">On the band</Paragraph>
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(elementOf(xml, "w:p", "On the band"), '<w:color w:val="FFFFFF"/>');
});

test("a proof labels what the engine has still to write, when the theme asks", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Written id="summary" />
      </Document>,
    ),
  );

  assertStringIncludes(xml, "Dynamic: summary");
  assertStringIncludes(elementOf(xml, "w:p", "What goes here."), '<w:color w:val="5C6B64"/>');
});

test("a theme without a dynamic label prints a placeholder as it always has", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Written id="summary" />
    </Document>,
  );
  const { dynamicLabel: _, standIn: __, ...blocks } = style.blocks!;
  const xml = await documentXml({ ...doc, style: { ...style, blocks } });

  assertEquals(xml.includes("Dynamic:"), false);
});

test("a variant can turn the label off, and a cell can name its look while it waits", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell variant="stamp">
              <Written id="decision" variant="word" />
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertEquals(xml.includes("Dynamic: decision"), false);
  const cell = elementOf(xml, "w:tbl", "What goes here.");
  assertStringIncludes(cell, 'w:fill="FFFFFF"');
  assertStringIncludes(cell, 'w:val="dashed"');
});

test("written text with blank lines in it is set as that many paragraphs", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Written id="letter" variant="letter" />
      </Document>,
      "resolve",
    ),
  );

  for (const text of ["One.", "Two.", "Three."]) {
    assertStringIncludes(elementOf(xml, "w:p", text), 'w:after="160"');
  }
});

test("a page number takes its cell's face, around whatever it is separated by", async () => {
  const xml = await partXml(
    await build(
      <Document
        id="d"
        title="D"
        footer={
          <Table id="f" columns={[{ width: "auto" }]}>
            <Row>
              <Cell variant="tracked">
                <PageNumber id="n" format="currentOfTotal" separator=" OF " />
              </Cell>
            </Row>
          </Table>
        }
      >
        <Paragraph id="a">Body.</Paragraph>
      </Document>,
    ),
    "word/footer1.xml",
  );

  assertStringIncludes(xml, '<w:t xml:space="preserve"> OF </w:t>');
  // Every run of the number carries the cell's tracking, not only the words around it.
  assertEquals((xml.match(/<w:spacing w:val="16"\/>/g) ?? []).length, 3);
});

test("a cell's only paragraph is set at the cell's leading, and a bar inside a cell adds no line", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="outer" columns={[{ width: 60 }]}>
          <Row>
            <Cell>
              <Table id="bar" columns={[{ width: 60 }]}>
                <Row height={3} heightRule="exact">
                  <Cell variant="segment"></Cell>
                </Row>
              </Table>
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The segment's paragraph is one point deep, not a line of body text.
  assertStringIncludes(xml, '<w:spacing w:after="0" w:line="20" w:lineRule="exact"/>');
  // And the outer cell ends on the one-point separator, straight after the bar.
  const afterBar = xml.slice(xml.indexOf("</w:tbl>"), xml.lastIndexOf("</w:tbl>"));
  assertStringIncludes(afterBar, 'w:line="20" w:lineRule="exact"');
});


// ---------------------------------------------------------------------------
// Charts drawn to a design
// ---------------------------------------------------------------------------

/** The first chart part a document packs, as XML. */
async function chartXml(doc: DocumentModel): Promise<string> {
  return await partXml(doc, "word/charts/chart1.xml");
}

test("a styled line chart packs its dashes, band, markers, grid and plot as Word reads them", async () => {
  const xml = await chartXml(
    await build(
      <Document id="d" title="D">
        <Graph
          id="g"
          graphType="line"
          legend="none"
          markers={false}
          gridlines={false}
          textSizePt={7}
          width={300}
          height={150}
          plotArea={{ leftPt: 30, topPt: 15, rightPt: 30, bottomPt: 15 }}
          fromEdge
          data={{
            categories: ["a", "b", "c"],
            series: [
              { label: "Range", values: [3, 4, 5], band: [1, 1, 2], color: "DBEEE5" },
              { label: "Benchmark", values: [1, 2, 3], color: "5C6B64", dash: "dashed", widthPt: 1 },
              { label: "Central", values: [2, 3, 4], color: "17624A", endLabel: true },
            ],
          }}
        />
      </Document>,
    ),
  );

  // The band is a stacked area under the lines: an unfilled foot, then the depth.
  assertStringIncludes(xml, '<c:areaChart><c:grouping val="stacked"/>');
  assertStringIncludes(xml, "<c:v>2</c:v>"); // the depth over the foot at the first point: 3 − 1
  assertStringIncludes(xml, '<a:prstDash val="dash"/>');
  assertStringIncludes(xml, '<a:ln w="12700"');
  assertStringIncludes(xml, '<c:symbol val="none"/>');
  assertEquals(xml.includes("<c:majorGridlines>"), false);
  assertStringIncludes(xml, 'sz="700"');
  assertStringIncludes(xml, '<c:x val="0.1000"/><c:y val="0.1000"/><c:w val="0.8000"/><c:h val="0.8000"/>');
  assertStringIncludes(xml, '<c:crossBetween val="midCat"/>');
  assertStringIncludes(xml, '<c:dLbl><c:idx val="2"/>');
  assertStringIncludes(xml, '<c:showSerName val="1"/>');
});

test("a sparkline packs with its axes deleted", async () => {
  const spark = await chartXml(
    await build(
      <Document id="d" title="D">
        <Graph id="s" graphType="line" axes={false} data={{ series: [{ values: [1, 3, 2] }] }} />
      </Document>,
    ),
  );
  assertEquals((spark.match(/<c:delete val="1"\/>/g) ?? []).length, 2);

});

test("the preview reads a styled chart back as the file packed it", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Graph
        id="g"
        graphType="line"
        markers={false}
        width={300}
        height={150}
        plotArea={{ leftPt: 30, topPt: 15, rightPt: 30, bottomPt: 15 }}
        fromEdge
        data={{
          categories: ["a", "b", "c"],
          series: [
            { label: "Range", values: [3, 4, 5], band: [1, 1, 2], color: "DBEEE5" },
            { label: "Benchmark", values: [1, 2, 3], dash: "dashed" },
          ],
        }}
      />
    </Document>,
  );
  const { createDocxBlob } = await import("docxcelerate/docx");
  const { readPackedCharts } = await import("docxcelerate/preview");
  const [chart] = await readPackedCharts(new Uint8Array(await (await createDocxBlob(doc)).arrayBuffer()));

  assertEquals(chart!.markers, false);
  assertEquals(chart!.fromEdge, true);
  assertEquals(chart!.stacked, false);
  assertEquals(chart!.plotArea, { x: 0.1, y: 0.1, w: 0.8, h: 0.8 });
  assertEquals(chart!.series[0]!.band, [1, 1, 2]);
  assertEquals(chart!.series[0]!.values, [3, 4, 5]);
  assertEquals(chart!.series[1]!.dash, "dashed");
});

test("a table in a cell shares its auto columns from the cell, not the page", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="outer" columns={[{ width: 60 }, { width: "auto" }]}>
          <Row>
            <Cell>
              <Table id="inner" columns={[{ width: 20 }, { width: "auto" }]}>
                <Row>
                  <Cell>a</Cell>
                  <Cell>b</Cell>
                </Row>
              </Table>
            </Cell>
            <Cell>c</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // 60mm less the cell's default 8pt each side is 2948 twips; 20mm of it is 1134.
  const inner = xml.slice(xml.indexOf("<w:tbl>", xml.indexOf("<w:tbl>") + 1));
  const grid = [...inner.slice(0, inner.indexOf("</w:tblGrid>")).matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => Number(m[1]));
  assertEquals(grid[0], 1134);
  assertEquals(grid[0]! + grid[1]! <= 3402 - 320, true);
});

test("written text opens on a drop cap; a placeholder and the later paragraphs do not", async () => {
  const capStyle: DocumentStyle = {
    ...style,
    blocks: {
      ...style.blocks,
      opening: { lineHeight: 1.5, dropCap: { lines: 2, variant: "cap" } },
      cap: { fontSizePt: 36, weight: "bold", color: "17624A" },
    },
  };
  const written = await documentXml({
    ...(await build(
      <Document id="d" title="D">
        <Written id="letter" variant="opening" />
      </Document>,
      "resolve",
    )),
    style: capStyle,
  });

  assertStringIncludes(written, 'w:dropCap="drop"');
  assertStringIncludes(written, 'w:lines="2"');
  assertEquals((written.match(/w:dropCap=/g) ?? []).length, 1);
  assertEquals(/<w:t[^>]*>O<\/w:t>/.test(written), true);
  assertEquals(/<w:t[^>]*>ne\.<\/w:t>/.test(written), true);

  const proof = await documentXml({
    ...(await build(
      <Document id="d" title="D">
        <Written id="letter" variant="opening" />
      </Document>,
    )),
    style: capStyle,
  });
  assertEquals(proof.includes("w:dropCap"), false);
});

test("a numbered list can count in bare figures, set in a variant of their own", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p1" variant="point" list="number" listStyle="decimalPlain">One</Paragraph>
      <Paragraph id="p2" variant="point" list="number" listStyle="decimalPlain">Two</Paragraph>
    </Document>,
  );
  const numbering = await partXml({
    ...doc,
    style: { ...style, blocks: { ...style.blocks, point: { marker: "figure" }, figure: { font: "Georgia", color: "17624A" } } },
  }, "word/numbering.xml");

  assertStringIncludes(numbering, '<w:lvlText w:val="%1"/>');
  assertStringIncludes(numbering, 'w:ascii="Georgia"');
  assertStringIncludes(numbering, '<w:color w:val="17624A"/>');
});

test("a paragraph border keeps its words off the rule by the padding of that side", async () => {
  const xml = await documentXml({
    ...(await build(
      <Document id="d" title="D">
        <Paragraph id="p" variant="ruled">Point</Paragraph>
      </Document>,
    )),
    style: { ...style, blocks: { ...style.blocks, ruled: { border: "DFE5E2", borderSides: ["bottom"], paddingSidesPt: { bottom: 8 } } } },
  });

  assertStringIncludes(elementOf(xml, "w:p", "Point"), 'w:space="8"');
});

test("the preview adds a paragraph's space after to the next one's space before, as Word does", async () => {
  const { settleDocxPreview } = await import("docxcelerate/preview");
  const { document } = new JSDOM(
    "<article><p style='margin-bottom: 7pt'>Label</p><p style='margin-top: 8pt'>Point</p></article>",
  ).window;
  settleDocxPreview(document.body);
  assertEquals((document.querySelectorAll("p")[1] as HTMLElement).style.marginTop, "15pt");
});
