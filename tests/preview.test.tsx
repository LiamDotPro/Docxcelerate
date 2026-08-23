import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import {
  buildProjectFinalDocument,
  buildProjectPreviewDocument,
  defineDocumentProject,
  derive,
} from "docxcelerate";
import type { DocumentNode, ParagraphNode } from "docxcelerate";
import { Document, Paragraph, Section, template, useAi, useState } from "docxcelerate/template";

/**
 * Why a preview is quick, and what it gives up to stay that way.
 *
 * A preview is rebuilt every time a file is saved. That is the whole point of
 * it: a document is written by looking at it, and a person who has to wait to
 * see what they typed stops looking. So the two things that cost real time —
 * asking a model to write a node, and running a deriver that renders a code or
 * calls a service — stand in rather than run, and are resolved for real only
 * when a document actually is.
 *
 * The cost of getting this wrong is not a slow test. It is a preview somebody
 * closes, and a document nobody proofreads.
 */
interface OrderData {
  reference: string;
  amount: number;
}

const order: OrderData = { reference: "INV-2026-0142", amount: 22380 };

/** How many times each deriver actually ran, so the preview can be held to it. */
const ran = { total: 0, qr: 0 };

const derivers = [
  {
    // Cheap by construction: a number the page should show for real.
    name: "money",
    run: ([amount]: unknown[]) => {
      ran.total += 1;
      return `£${Number(amount ?? 0).toLocaleString("en-GB")}`;
    },
  },
  {
    // Costly: this stands for rendering a scan-to-pay code.
    name: "paymentQr",
    run: () => {
      ran.qr += 1;
      return "data:image/png;base64,REAL";
    },
    placeholder: "[scan-to-pay code]",
  },
];

const Summary = () => {
  const [data] = useState((input: OrderData) => input);

  useAi({
    ask: "One sentence describing the order.",
    placeholder: "A one-line summary of the order.",
    from: { reference: data.reference },
  });

  return <Paragraph id="summary" />;
};

const Total = () => (
  <Paragraph id="total" derivers={[derive("money", { output: "total", inputs: [22380] })]}>
    Total due {"{{derived.total}}"}.
  </Paragraph>
);

const Qr = () => (
  <Paragraph id="qr" derivers={[derive("paymentQr", { output: "code", inputs: [] })]}>
    {"{{derived.code}}"}
  </Paragraph>
);

const project = defineDocumentProject<OrderData>({
  id: "order",
  name: "Order",
  template: template<OrderData>(
    <Document id="order" title="Order">
      <Section id="body" title="Body">
        <Summary />
        <Total />
        <Qr />
      </Section>
    </Document>,
  ),
  previewData: order,
  derivers,
});

function textOf(nodes: DocumentNode[], id: string): string {
  const section = nodes[0];

  if (section?.kind !== "section") {
    throw new Error("expected the body section");
  }

  const node = section.children.find((child) => child.id === id);

  return (node as ParagraphNode | undefined)?.text ?? "";
}

function reset(): void {
  ran.total = 0;
  ran.qr = 0;
}

// ---------------------------------------------------------------------------
// What a preview refuses to wait for
// ---------------------------------------------------------------------------

test("a preview shows the placeholder rather than calling a model", async () => {
  reset();
  const preview = await buildProjectPreviewDocument(project);

  assertEquals(textOf(preview.nodes, "summary"), "A one-line summary of the order.");
});

test("a preview needs no AI client at all, so it cannot be blocked by one", async () => {
  reset();
  // No aiClient is given anywhere. A preview that needed one would throw here,
  // and a document could not be previewed without credentials for a model.
  const preview = await buildProjectPreviewDocument(project);

  assertEquals(preview.nodes.length, 1);
});

test("a costly deriver stands in rather than running", async () => {
  reset();
  await buildProjectPreviewDocument(project);

  assertEquals(ran.qr, 0);
});

test("what the costly deriver declared is what the preview shows", async () => {
  reset();
  const preview = await buildProjectPreviewDocument(project);

  assertEquals(textOf(preview.nodes, "qr"), "[scan-to-pay code]");
});

test("a cheap deriver still runs, so the figures on the page are the real ones", async () => {
  reset();
  const preview = await buildProjectPreviewDocument(project);

  assertEquals(ran.total, 1);
  assertStringIncludes(textOf(preview.nodes, "total"), "£22,380");
});

test("the same preview twice produces the same document", async () => {
  reset();
  const first = await buildProjectPreviewDocument(project);
  const second = await buildProjectPreviewDocument(project);

  // A preview that reshuffled itself between saves is one nobody can proofread.
  assertEquals(first.nodes, second.nodes);
});

// ---------------------------------------------------------------------------
// What a real document does instead
// ---------------------------------------------------------------------------

test("writing a document for real runs the deriver a preview stood in for", async () => {
  reset();
  await buildProjectFinalDocument(project, {
    aiClient: { generateParagraph: () => "Written." },
  });

  assertEquals(ran.qr, 1);
});

test("the real deriver result reaches the document, not the stand-in", async () => {
  reset();
  const written = await buildProjectFinalDocument(project, {
    aiClient: { generateParagraph: () => "Written." },
  });

  assertEquals(textOf(written.nodes, "qr"), "data:image/png;base64,REAL");
});

test("writing a document for real calls the model the preview skipped", async () => {
  reset();
  let asked = 0;
  const written = await buildProjectFinalDocument(project, {
    aiClient: {
      generateParagraph: () => {
        asked += 1;
        return "The order covers August delivery work.";
      },
    },
  });

  assertEquals(asked, 1);
  assertEquals(textOf(written.nodes, "summary"), "The order covers August delivery work.");
});

test("a preview and a written document differ in what a value says, never in whether it is there", async () => {
  reset();
  const preview = await buildProjectPreviewDocument(project);
  const written = await buildProjectFinalDocument(project, {
    aiClient: { generateParagraph: () => "Written." },
  });

  const ids = (doc: typeof preview) => {
    const section = doc.nodes[0];
    return section.kind === "section" ? section.children.map((child) => child.id) : [];
  };

  // The stand-ins change the words, not the shape. A preview that dropped a
  // node would be a preview of a different document.
  assertEquals(ids(preview), ids(written));
});
