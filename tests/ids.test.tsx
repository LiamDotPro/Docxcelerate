import { test } from "node:test";
import { assertEquals } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import type { DocumentNode } from "docxcelerate";
import {
  Document,
  Graph,
  Image,
  Paragraph,
  Section,
  TableOfContents,
  template,
  useState,
} from "docxcelerate/template";

/**
 * The names nodes take when nobody writes one.
 *
 * An id is an address. An engine targets a node by it, two builds of the same
 * document line up in a diff by it, and a rename is a breaking change. That is
 * exactly why an id nobody wrote has to be worth having: if the automatic ones
 * are junk, everybody writes them by hand to avoid them, and the option to leave
 * them out is no option at all.
 *
 * The rule these all check is that a name comes from something that already says
 * what the node is — its heading, or the component that yielded it — and never
 * from where it happens to sit.
 */
interface Data {
  name: string;
  lines: Array<{ label: string }>;
}

const data: Data = { name: "Avery", lines: [{ label: "One" }, { label: "Two" }] };

function idsOf(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => [
    node.id,
    ...(node.kind === "section" || node.kind === "repeat" ? idsOf(node.children) : []),
  ]);
}

async function ids(body: unknown): Promise<string[]> {
  const built = await buildDocument(
    template<Data>(<Document title="Doc">{body as never}</Document>),
    data,
  );

  return idsOf(built.nodes);
}

// ---------------------------------------------------------------------------
// Where a name comes from
// ---------------------------------------------------------------------------

test("a section is named after its heading", async () => {
  assertEquals(await ids(<Section title="Fees and funding" />), ["fees-and-funding"]);
});

test("a node is named after the component that yielded it", async () => {
  const Greeting = () => <Paragraph>Dear Avery,</Paragraph>;

  assertEquals(await ids(<Greeting />), ["greeting"]);
});

test("a component written in camel case becomes a hyphenated name", async () => {
  const SignOff = () => <Paragraph>Kind regards,</Paragraph>;

  assertEquals(await ids(<SignOff />), ["sign-off"]);
});

test("an initialism keeps its word boundary", async () => {
  const VATSummary = () => <Paragraph>VAT.</Paragraph>;

  assertEquals(await ids(<VATSummary />), ["vat-summary"]);
});

test("an element written straight into the template falls back to its kind", async () => {
  assertEquals(await ids(<Paragraph>Plain.</Paragraph>), ["paragraph"]);
});

test("an explicit id always wins", async () => {
  const Greeting = () => <Paragraph id="hello">Dear Avery,</Paragraph>;

  assertEquals(await ids(<Greeting />), ["hello"]);
});

test("a heading wins over the component that yielded it", async () => {
  // The heading is the more specific of the two: two sections from one
  // component are told apart by what they are headed, not by what made them.
  const Block = () => <Section title="Payment details" />;

  assertEquals(await ids(<Block />), ["payment-details"]);
});

test("punctuation in a heading does not reach the address", async () => {
  assertEquals(await ids(<Section title="Terms &amp; notes (2026)" />), ["terms-notes-2026"]);
});

test("a heading of nothing but punctuation still leaves a usable name", async () => {
  assertEquals(await ids(<Section title="—" />), ["section"]);
});

// ---------------------------------------------------------------------------
// Names stay put
// ---------------------------------------------------------------------------

test("inserting a node above another leaves the other's name alone", async () => {
  const Greeting = () => <Paragraph>Dear Avery,</Paragraph>;
  const SignOff = () => <Paragraph>Kind regards,</Paragraph>;
  const Notice = () => <Paragraph>Please note.</Paragraph>;

  const before = await ids(
    <Section title="Body">
      <Greeting />
      <SignOff />
    </Section>,
  );
  const after = await ids(
    <Section title="Body">
      <Greeting />
      <Notice />
      <SignOff />
    </Section>,
  );

  // This is the property the old positional names could not hold: the sign-off
  // is addressed the same way whether or not something was added above it.
  assertEquals(before, ["body", "greeting", "sign-off"]);
  assertEquals(after, ["body", "greeting", "notice", "sign-off"]);
});

test("reordering nodes does not rename them", async () => {
  const Greeting = () => <Paragraph>Dear Avery,</Paragraph>;
  const SignOff = () => <Paragraph>Kind regards,</Paragraph>;

  const forwards = await ids(<Section title="Body"><Greeting /><SignOff /></Section>);
  const backwards = await ids(<Section title="Body"><SignOff /><Greeting /></Section>);

  assertEquals(new Set(forwards), new Set(backwards));
});

// ---------------------------------------------------------------------------
// Two of the same thing
// ---------------------------------------------------------------------------

test("a second node with the same derived name is numbered rather than refused", async () => {
  const Greeting = () => <Paragraph>Dear Avery,</Paragraph>;

  assertEquals(
    await ids(<Section title="Body"><Greeting /><Greeting /></Section>),
    ["body", "greeting", "greeting-2"],
  );
});

test("a third is numbered again", async () => {
  const Line = () => <Paragraph>A line.</Paragraph>;

  assertEquals(
    await ids(<Section title="Body"><Line /><Line /><Line /></Section>),
    ["body", "line", "line-2", "line-3"],
  );
});

test("two explicit ids that collide are still refused", async () => {
  // Numbering is for names nobody chose. Two nodes given the same name by hand
  // is a typo, and quietly renaming one would hide it.
  let message = "";

  try {
    await ids(
      <Section title="Body">
        <Paragraph id="same">One.</Paragraph>
        <Paragraph id="same">Two.</Paragraph>
      </Section>,
    );
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertEquals(message.includes('Two nodes claim the id "same"'), true);
});

// ---------------------------------------------------------------------------
// Every kind gets a usable name
// ---------------------------------------------------------------------------

test("an image with no id is named after its kind", async () => {
  assertEquals(await ids(<Image src="logo.png" alt="A logo" />), ["image"]);
});

test("a graph with no id is named after its kind", async () => {
  assertEquals(await ids(<Graph data={{ series: [{ values: [1] }] }} />), ["graph"]);
});

test("a table of contents with no id is named after its heading", async () => {
  assertEquals(await ids(<TableOfContents title="Contents" />), ["contents"]);
});

test("a loop names its passes from the node inside it", async () => {
  const Lines = () => {
    const [lines] = useState((input: Data) => input.lines);

    return lines.map((line) => <Paragraph>{line.label}</Paragraph>);
  };

  assertEquals(await ids(<Lines />), ["lines-0", "lines-1"]);
});

// ---------------------------------------------------------------------------
// The document itself
// ---------------------------------------------------------------------------

test("a document is named after its title", () => {
  const tree = template<Data>(<Document title="Tenancy Renewal" />);

  assertEquals(tree.id, "tenancy-renewal");
});

test("a document that says its own id keeps it", () => {
  const tree = template<Data>(<Document id="renewal-2026" title="Tenancy Renewal" />);

  assertEquals(tree.id, "renewal-2026");
});
