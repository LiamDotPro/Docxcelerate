import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDeriverRegistry,
  EchoAiClient,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import { documentXml, partNames, partXml } from "./docx.ts";
import type { DocumentModel, DocumentStyle } from "docxcelerate";
import {
  Cell,
  Document,
  Image,
  PageBreak,
  PageNumber,
  Paragraph,
  Row,
  Table,
  template,
  createPublishData,
} from "docxcelerate/template";

/**
 * What Word is actually handed.
 *
 * A `.docx` is a zip of XML, so the way to know a fill or a border survived
 * packing is to open the file and look. Without this, a style could silently
 * mean one thing where it is authored and nothing at all in the format the
 * framework exists to produce — which is how `radiusPt` came to be a property
 * that did nothing here.
 */

const style: DocumentStyle = {
  preset: "test",
  page: {
    size: "A4",
    orientation: "portrait",
    margins: { topMm: 16, rightMm: 16, bottomMm: 16, leftMm: 16 },
  },
  typography: {
    bodyFont: "Aptos",
    headingFont: "Aptos",
    bodySizePt: 10,
    bodyLineHeight: 1.4,
    color: "1C2340",
  },
  palette: {
    heading: "2C3D8F",
    accent: "2C3D8F",
    muted: "5A6482",
    rule: "D9DDEB",
    page: "FFFFFF",
  },
  paragraph: { spacingAfterPt: 6 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 10 },
  sectionHeading: { fontSizePt: 8, weight: "bold", spacingBeforePt: 10, spacingAfterPt: 4 },
  blocks: {
    band: {
      fill: "F4F6FD",
      bleed: true,
      border: "E3E7F5",
      borderSides: ["bottom"],
      paddingPt: 10,
    },
    totalRow: {
      fill: "1E2A66",
      color: "FFFFFF",
      weight: "bold",
      fontSizePt: 11,
    },
    badge: {
      fill: "FBF0DC",
      border: "E5C78A",
      color: "8A5A06",
      paddingPt: 5,
      fontSizePt: 7,
      weight: "bold",
      transform: "uppercase",
      letterSpacingEm: 0.1,
    },
    /** A money column: its own face, and set tighter than prose. */
    money: {
      font: "Consolas",
      lineHeight: 1.2,
    },
    /** The muted note under a description, tighter still. */
    note: {
      color: "5A6482",
      fontSizePt: 8,
      lineHeight: 1.1,
    },
    /** Every other body row, tinted — the renderer finds this by name. */
    rowAlt: {
      fill: "F7F8FD",
    },
    /** A block that draws no edges, and says so. */
    noRules: {
      borderSides: [],
    },
    /** A block that sits its contents on the line of the tallest cell. */
    centred: {
      valign: "center",
    },
    /** Prose held to a measure narrower than the text column. */
    measured: {
      maxWidthMm: 100,
    },
    /** A strip of colour that says its own depth. */
    hairline: {
      fill: "2C3D8F",
      heightPt: 2.25,
      spacingAfterPt: 0,
    },
    /** A bar whose last words stop short of the edge it runs to. */
    edgeBar: {
      fill: "1E2A66",
      paddingPt: 6,
      paddingSidesPt: { right: 46 },
    },
    /** Named on a cell inside a `band` table: says leading, nothing else. */
    tightCell: {
      lineHeight: 1.2,
    },
  },
};

/** A one-pixel PNG, so a picture test has bytes to embed. */
const PIXEL = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function build(node: unknown): Promise<DocumentModel> {
  return buildDocument(
    template<Record<string, never>>(node as never),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  ).then((doc) => ({ ...doc, style }));
}

// ---------------------------------------------------------------------------
// A block means the same thing in both outputs
// ---------------------------------------------------------------------------

test("a block's fill is shading in the Word file, not just colour on screen", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(xml, 'w:fill="F4F6FD"');
});

test("a block's border is drawn, on the edges it named", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(xml, "<w:pBdr>");
  assertStringIncludes(xml, 'w:color="E3E7F5"');
  assertStringIncludes(xml, "<w:bottom");
  // Only the bottom was asked for, so no other edge is drawn.
  assertEquals(/<w:pBdr>(?:(?!<\/w:pBdr>).)*<w:top /s.test(xml), false);
});

test("a bleeding block reaches past the margin, by indenting negatively", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
    ),
  );

  // 16mm of margin, so the block starts 16mm to the left of it.
  assertStringIncludes(xml, 'w:left="-907"');
});

test("letter spacing survives, converted from ems to what Word counts in", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="badge">Awaiting</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // 0.1em of a 7pt face is 0.7pt, and Word counts in twentieths of a point.
  assertStringIncludes(xml, '<w:spacing w:val="14"/>');
});

test("a cell's variant sets its text, not only its background", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="badge">Awaiting</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(xml, 'w:fill="FBF0DC"');
  assertStringIncludes(xml, 'w:val="8A5A06"');
  assertStringIncludes(xml, "<w:caps/>");
  assertStringIncludes(xml, "<w:tcBorders>");
  assertStringIncludes(xml, "<w:tcMar>");
});

test("a block can name its own face, so a money column lines up", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: "auto" }]}>
          <Row>
            <Cell id="desc">Discovery workshop</Cell>
            <Cell id="amount" variant="money">£1,520.00</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // Proportional digits do not line up under one another, so the figure needs
  // a face whose digits are all one width — and the prose beside it must not
  // be dragged into it.
  assertStringIncludes(xml, '<w:rFonts w:ascii="Consolas"');
  const amountAt = xml.indexOf("1,520.00");
  const descAt = xml.indexOf("Discovery workshop");
  assertEquals(xml.slice(descAt - 400, descAt).includes("Consolas"), false);
  assertEquals(xml.slice(amountAt - 400, amountAt).includes("Consolas"), true);
});

test("a block can set its own leading, and a paragraph's beats its cell's", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="money">
              <Paragraph id="p1">Discovery workshop</Paragraph>
              <Paragraph id="p2" variant="note">Sprint 14, two consultants</Paragraph>
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // Stated in points, and exactly: 1.2 of a 10pt face is 12pt (240 twips) for
  // the cell, 1.1 of the note's 8pt is 8.8pt (176) inside it, and the body's
  // 1.4 of 10pt (280) is neither. A multiple would have meant one thing to
  // Word and another to a browser; a number of points means one thing.
  assertStringIncludes(xml, '<w:spacing w:after="0" w:line="240" w:lineRule="exact"/>');
  assertStringIncludes(xml, '<w:spacing w:after="0" w:line="176" w:lineRule="exact"/>');
  assertEquals(xml.includes('w:line="280"'), false);
});

test("a block that names no leading is still set on the body's", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Ordinary prose.</Paragraph>
      </Document>,
    ),
  );

  // 1.4 of the body's 10pt is 14pt — unchanged by a block having gained the
  // option to differ, and still said as a measurement rather than a ratio.
  assertStringIncludes(xml, '<w:spacing w:after="120" w:line="280" w:lineRule="exact"/>');
});

test("a picture inside a paragraph shares its line rather than taking one", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="credit">
          <Image id="mark" src={PIXEL} alt="" width={8} height={8} />
          Generated with Docxcelerate
        </Paragraph>
      </Document>,
    ),
  );

  // One `w:p`, holding both the drawing and the words. Given a paragraph of
  // its own the mark becomes a picture with a caption under it, and a one-line
  // footer bar three lines deep.
  const paragraphs = xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];
  const credit = paragraphs.filter((p) => p.includes("Generated with Docxcelerate"));
  assertEquals(credit.length, 1);
  assertEquals(credit[0].includes("<w:drawing>"), true);
  // And the mark leads the line, as it was written.
  assertEquals(
    credit[0].indexOf("<w:drawing>") < credit[0].indexOf("Generated with Docxcelerate"),
    true,
  );
});

test("a paragraph with no picture is the single run it always was", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="p">Just words.</Paragraph>
      </Document>,
    ),
  );

  const paragraphs = (xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [])
    .filter((p) => p.includes("Just words."));
  assertEquals(paragraphs.length, 1);
  assertEquals((paragraphs[0].match(/<w:r>/g) ?? []).length, 1);
});

test("a page number in a cell takes the column, not the body's own mind", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: 18, align: "right" }]}>
          <Row>
            <Cell id="left">Registered in England</Cell>
            <Cell id="right">
              <PageNumber id="n" />
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The field paragraph is the cell's, so it takes the column's alignment and
  // the cell's zero spacing — not the 6pt gap that belongs between blocks of
  // prose, which is the rest of why a footer bar came out three lines deep.
  const fieldPara = (xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [])
    .find((p) => p.includes("PAGE"));
  assertEquals(fieldPara !== undefined, true);
  assertStringIncludes(fieldPara ?? "", '<w:jc w:val="right"/>');
  assertStringIncludes(fieldPara ?? "", 'w:after="0"');
});

test("a bleeding table reaches the paper's edge, not the margin", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="bar" variant="band" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c">Registered in England</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // 16mm of margin, so the table starts 16mm left of it and its columns fill
  // the whole 210mm sheet rather than the 178mm text column.
  assertStringIncludes(xml, '<w:tblInd w:type="dxa" w:w="-907"/>');
  assertStringIncludes(xml, '<w:tblW w:type="dxa" w:w="11906"/>');
});

test("a table that does not bleed still stands on the text column", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c">Inside the margins.</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertEquals(xml.includes("<w:tblInd"), false);
  assertStringIncludes(xml, '<w:tblW w:type="dxa" w:w="10092"/>');
});

test("naming only a first header does not cost the first page its footer", async () => {
  const doc = await build(
    <Document
      id="d"
      title="D"
      header={<Paragraph id="h">Running header</Paragraph>}
      firstHeader={false}
      footer={<Paragraph id="f">Running footer</Paragraph>}
    >
      <Paragraph id="a">Body.</Paragraph>
    </Document>,
  );
  const parts = await partNames(doc);
  const firstFooter = await Promise.all(
    parts.filter((name) => /footer\d+\.xml$/.test(name)).map((name) => partXml(doc, name)),
  );

  // `w:titlePg` is one switch for both strips, so turning it on for the header
  // makes Word take page one's footer from a `first` part too. Absent means
  // "like every other page" — so that part has to carry the running footer,
  // not be missing and leave page one bare.
  assertEquals(firstFooter.filter((xml) => xml.includes("Running footer")).length, 2);
});

test("the theme's zebra is drawn by the renderer, not chosen by the document", async () => {
  const rows = ["One", "Two", "Three", "Four", "Five"];
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row header>
            <Cell id="h">Description</Cell>
          </Row>
          {rows.map((text) => (
            <Row>
              <Cell>{text}</Cell>
            </Row>
          ))}
        </Table>
      </Document>,
    ),
  );

  // Body rows 2 and 4 are tinted, 1, 3 and 5 are not, and the header keeps the
  // accent. Nothing in the document said which row is which — a variant chosen
  // by a map index would be baked once at build time and repeat on publish.
  const cells = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? [];
  const tinted = rows.map((text) =>
    (cells.find((cell) => cell.includes(`>${text}<`)) ?? "").includes('w:fill="F7F8FD"')
  );
  assertEquals(tinted, [false, true, false, true, false]);
  assertStringIncludes(xml, 'w:fill="2C3D8F"');
});

test("a table with no header row is not a list, so it is not striped", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: 50 }, { width: 34 }]}>
          <Row>
            <Cell></Cell>
            <Cell>Subtotal</Cell>
            <Cell>18,650.00</Cell>
          </Row>
          <Row>
            <Cell></Cell>
            <Cell>VAT (20%)</Cell>
            <Cell>3,730.00</Cell>
          </Row>
          <Row>
            <Cell></Cell>
            <Cell variant="totalRow">Total due</Cell>
            <Cell variant="totalRow">22,380.00</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // A totals block is three rows and a spacer column, not a list anybody
  // scans down. Striping it tinted whichever cells landed on the second row —
  // including the empty spacer, which came out as a grey band of nothing
  // beside the figures rather than under them.
  //
  // A zebra is a reading aid for a column of like rows, and what says a table
  // is one of those is its header. Without one there is nothing to help read.
  assertEquals(xml.includes('w:fill="F7F8FD"'), false);
  // The row that asked for a fill still has it.
  assertStringIncludes(xml, 'w:fill="1E2A66"');
});

test("a row that draws its own ground draws no rule under the rest of it", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: 50 }, { width: 34 }]}>
          <Row>
            <Cell></Cell>
            <Cell variant="totalRow">Total due</Cell>
            <Cell variant="totalRow">22,380.00</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The row hairline is a property of the row, not of each cell. Deciding it
  // per cell from that cell's own fill drew it under the columns that
  // happened to be plain and not under the ones that were not — on the
  // invoice, three rules to the left of the totals panel, each stopping where
  // the panel began. A rule across part of a row is not a rule.
  //
  // Once any of the row is filled the row is already set apart, which is what
  // the rule was for.
  const cells = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? [];
  const ruled = cells.filter((cell) => cell.includes("w:bottom") && cell.includes('w:val="single"'));

  assertEquals(ruled.length, 0);
  assertStringIncludes(xml, 'w:fill="1E2A66"');
});

test("a document decorates the same whether it was built here or by an engine", async () => {
  const tree = template(
    <Document id="d" title="D">
      <Table id="t" columns={[{ width: "auto" }, { width: 34 }]}>
        <Row header>
          <Cell>Description</Cell>
          <Cell>Amount</Cell>
        </Row>
        <Row>
          <Cell>One</Cell>
          <Cell>1.00</Cell>
        </Row>
        <Row>
          <Cell>Two</Cell>
          <Cell>2.00</Cell>
        </Row>
      </Table>
    </Document>,
  );

  // Striping and the row rule are decided when the file is packed, from
  // `header` on the row. That is only sound if the flag reaches the packer on
  // both roads to it: the build that produces a preview, and the published
  // document an engine resolves against data nobody had at build time.
  //
  // If it travelled on one and not the other, a person would be sent a file
  // decorated differently from the one they approved — which is the single
  // thing a preview exists not to do.
  const here = await buildDocument(tree, {});
  const published = await buildDocument(tree, createPublishData(), {
    branchMode: "publish",
    deriverMode: "preserve",
  });
  const sent = await resolveDocument(published, {
    ctx: {},
    derived: {},
    dataProvider: new InMemoryDataProvider({}),
    aiClient: new EchoAiClient(),
  }, { derivers: createDeriverRegistry() });

  here.style = style;
  sent.style = style;

  assertEquals(decorationOf(await documentXml(here)), decorationOf(await documentXml(sent)));
  // And it is the decoration the theme asked for, not two matching blanks.
  assertEquals(decorationOf(await documentXml(here)), [
    { fill: "2C3D8F", rule: false },
    { fill: "2C3D8F", rule: false },
    { fill: null, rule: true },
    { fill: null, rule: true },
    { fill: "F7F8FD", rule: false },
    { fill: "F7F8FD", rule: false },
  ]);
});

/** Every cell's fill and whether a rule was drawn under it, in table order. */
function decorationOf(xml: string): Array<{ fill: string | null; rule: boolean }> {
  return (xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []).map((cell) => ({
    fill: cell.match(/w:fill="([0-9A-Fa-f]{6})"/)?.[1] ?? null,
    rule: /<w:bottom [^>]*w:val="single"/.test(cell),
  }));
}

test("a block that names no edges turns the row hairline off with them", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="noRules">No rules.</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The row hairline is the default for an unfilled body row. Naming no edges
  // is a decision, so it does not step in — the alternative today was a border
  // in the page colour, which is a lie written into the file.
  assertEquals(xml.includes("<w:tcBorders>"), false);
  assertEquals(xml.includes('w:color="D9DDEB"'), false);
});

test("a block can sit its contents against the height of the cell", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="centred">Centred.</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(xml, '<w:vAlign w:val="center"/>');
});

test("a heading style carries its tracking into styles.xml", async () => {
  const styles = await partXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Body.</Paragraph>
      </Document>,
    ),
    "word/styles.xml",
  );

  // The test theme tracks neither, so nothing is written — the property has to
  // be absent rather than zero, or every document gains a spacing it never
  // asked for.
  assertEquals(/w:styleId="Heading1"[\s\S]*?<w:spacing w:val=/.test(styles), false);
});

test("a measure narrows the block from the right, leaving the margin alone", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="measured">Held to a measure.</Paragraph>
        <Paragraph id="b">Full width.</Paragraph>
      </Document>,
    ),
  );

  // A 178mm text column held to 100mm gives back 78mm on the right, and
  // nothing on the left — everything else still stands where it stood.
  assertStringIncludes(xml, 'w:right="4422"');
  assertEquals(xml.includes('w:left="4422"'), false);
});

test("a picture's variant draws a box around it, because a run cannot have one", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Image id="code" variant="badge" src={PIXEL} alt="Scan" width={108} height={108} />
      </Document>,
    ),
  );

  // A single-cell table at the picture's size plus its padding: the box holds
  // the same shape whether the picture has arrived or is still a label
  // standing in for one.
  assertStringIncludes(xml, "<w:tbl>");
  assertStringIncludes(xml, 'w:fill="FBF0DC"');
  assertStringIncludes(xml, "<w:tcMar>");
  assertStringIncludes(xml, "<w:drawing>");
  // 108pt plus 5pt of padding either side, in twentieths of a point.
  assertStringIncludes(xml, 'w:w="2360"');
});

test("a data URI carrying a parameter is still a picture", async () => {
  const svg = "data:image/svg+xml;utf8," +
    encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"></svg>');
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Image id="code" src={svg} fallbackSrc={PIXEL} alt="Code" width={20} height={20} />
      </Document>,
    ),
  );

  // `image/svg+xml;utf8` is a media type with a parameter on it. Read whole it
  // matches nothing, and a perfectly good picture came back as the note that
  // says one is missing.
  assertStringIncludes(xml, "<w:drawing>");
  assertEquals(xml.includes("[image: Code]"), false);
});

test("a cell's variant says what differs, not what everything is", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" variant="band" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="tightCell">Issue date</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // `tightCell` names a leading and nothing else, so the band's tint, border
  // and padding still reach it. Taking the narrower naming as a replacement
  // left a cell untinted inside a tinted band unless it restated the tint —
  // and the one that forgot is how this was found.
  assertStringIncludes(xml, 'w:fill="F4F6FD"');
  assertStringIncludes(xml, 'w:color="E3E7F5"');
  assertStringIncludes(xml, '<w:spacing w:after="0" w:line="240" w:lineRule="exact"/>');
});

test("a strip says how deep it is rather than shrinking a font to get there", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="rule" variant="hairline" />
      </Document>,
    ),
  );

  // 2.25pt of navy, stated. The alternative was a 1pt face at a fifth of a
  // line, which reaches the same depth and says nothing about what is drawn.
  assertStringIncludes(xml, 'w:fill="2C3D8F"');
  assertStringIncludes(xml, '<w:spacing w:after="0" w:line="45" w:lineRule="exact"/>');
});

test("a block can leave more room on one side than the others", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="edgeBar">1 / 2</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // A bar that runs to the paper's edge still wants its last words to stop
  // short of it. 6pt on three sides, 46 on the right — and no spacer column,
  // which is what holding the gap used to take.
  assertStringIncludes(xml, '<w:top w:type="dxa" w:w="120"/>');
  assertStringIncludes(xml, '<w:right w:type="dxa" w:w="920"/>');
});

test("a variant the theme has never heard of packs as an ordinary block", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="nonesuch">Text.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(xml.includes("<w:pBdr>"), false);
  assertStringIncludes(xml, "Text.");
});

// ---------------------------------------------------------------------------
// The furniture
// ---------------------------------------------------------------------------

test("a page break is Word's own break, carried by the paragraph after it", async () => {
  const broken = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">One.</Paragraph>
        <PageBreak id="turn" />
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
  );
  const unbroken = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">One.</Paragraph>
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
  );

  // A break-carrying style on the next paragraph, not a paragraph of its own —
  // the standalone form leaves an empty line at the foot of the outgoing page.
  assertStringIncludes(broken, '<w:pStyle w:val="PageBreakBefore"/>');
  assertEquals(broken.includes('w:type="page"'), false);
  assertEquals(unbroken.includes("PageBreakBefore"), false);
  assertEquals(unbroken.includes('w:type="page"'), false);
});

test("the break rides a style, which is the only form both engines read", async () => {
  const styles = await partXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">One.</Paragraph>
        <PageBreak id="turn" />
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
    "word/styles.xml",
  );

  // Word honours `w:pageBreakBefore` written straight onto a paragraph;
  // docx-preview reads it only off the paragraph's style. On the style is the
  // one form that turns the page in both, so the property has to actually
  // reach styles.xml — `docx` types its style options without it.
  assertStringIncludes(styles, '<w:style w:type="paragraph" w:styleId="PageBreakBefore">');
  assertStringIncludes(styles, "<w:pageBreakBefore/>");
  // The heading variant still inherits everything that makes a heading.
  assertStringIncludes(styles, '<w:basedOn w:val="Heading1"/>');
  // And supplying styles has not cost the document its defaults.
  assertStringIncludes(styles, "<w:docDefaults>");
  assertStringIncludes(styles, 'w:styleId="Heading1"');
});

test("a page number is a field Word recounts, not a digit written in", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D" footer={<PageNumber id="n" />}>
        <Paragraph id="a">One.</Paragraph>
      </Document>,
    ),
  );

  // The number lives in the footer part rather than the body, which is itself
  // the point — it is furniture, and Word fills it in per page.
  assertEquals(xml.includes("PAGE"), false);
});

test("a title the document prints itself is not printed again", async () => {
  const withTitle = await documentXml(
    await build(
      <Document id="d" title="Invoice">
        <Paragraph id="a">One.</Paragraph>
      </Document>,
    ),
  );
  const withoutTitle = await documentXml(
    await buildDocument(
      template<Record<string, never>>(
        (
          <Document id="d" title="Invoice">
            <Paragraph id="a">One.</Paragraph>
          </Document>
        ) as never,
      ),
      {},
      { branchMode: "decide", dynamicMode: "placeholder" },
    ).then((doc) => ({ ...doc, style: { ...style, showTitle: false } })),
  );

  assertStringIncludes(withTitle, "Invoice");
  assertEquals(withoutTitle.includes("Invoice"), false);
});
// ---------------------------------------------------------------------------
// What a table is drawn with
// ---------------------------------------------------------------------------

test("a table draws the theme's rules, not Word's default grid", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row header>
            <Cell id="h">Description</Cell>
          </Row>
          <Row>
            <Cell id="c">A line.</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The library's own default is a black box around every cell. Left alone it
  // means every document ships a grid nobody asked for and no preview shows.
  assertEquals(/<w:tblBorders>(?:(?!<\/w:tblBorders>).)*w:val="single"/s.test(xml), false);
  // What separates one row from the next is the palette's rule, under the
  // body cells only — a heading has its fill to set it apart.
  assertStringIncludes(xml, 'w:color="D9DDEB"');
  assertEquals(xml.match(/w:color="D9DDEB"/g)?.length, 1);
});

test("a header row the theme names keeps its own fill, not the accent", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: "auto" }]}>
          <Row header variant="totalRow">
            <Cell id="empty"></Cell>
            <Cell id="total">Total due</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // Every cell in the row, including the empty one beside the words: a bar in
  // two colours is what happens when the default shows through the gap.
  assertEquals(xml.match(/w:fill="1E2A66"/g)?.length, 2);
  assertEquals(xml.includes('w:fill="2C3D8F"'), false);
});

test("a heading row is set in small tracked capitals, as the screen sets it", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row header>
            <Cell id="h">Description</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(xml, "<w:caps/>");
  assertStringIncludes(xml, 'w:val="FFFFFF"');
  // 0.72 of a 10pt body, in half-points.
  assertStringIncludes(xml, '<w:sz w:val="14"/>');
});

test("a picture in a cell is not laid out as a paragraph of prose", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c">
              <Image id="mark" src={PIXEL} alt="" width={8} height={8} />
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The body's line and space-after belong to prose. On a picture they drop a
  // letterhead's mark below the name beside it and make a one-line footer bar
  // two lines deep.
  assertEquals(
    /<w:spacing w:after="120"[^>]*\/>(?:(?!<\/w:p>).)*<w:drawing>/s.test(xml),
    false,
  );
  assertStringIncludes(xml, "<w:drawing>");
});
