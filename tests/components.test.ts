import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import {
  buildDocument,
  defineDocument,
  EchoAiClient,
  graph,
  image,
  doc,
  type Node,
  type NodeComponent,
  paragraph,
  type ParagraphNode,
  section,
} from "docxcelerate";

interface LetterData {
  recipientName: string;
  households: number;
}

const Greeting: NodeComponent<LetterData, ParagraphNode> = (
  data,
): Node<ParagraphNode, LetterData> => ({
  kind: "paragraph",
  mode: "static",
  id: "greeting",
  render() {
    return `Hello ${data.recipientName},`;
  },
});

test("component API builds a static letter document tree", async () => {
  const letterTemplate = doc<LetterData>(
    {
      id: "welcome",
      title: "Welcome Letter",
    },
    [
      section(
        {
          id: "intro",
          title: "Introduction",
        },
        [
          Greeting,
          paragraph({
            id: "introText",
            render: () => "A simple static paragraph.",
          }),
        ],
      ),
    ],
  );
  const legacyLetterTemplate = defineDocument<LetterData>({
    id: "legacy",
    title: "Legacy Shape",
    nodes: [
      section({
        id: "intro",
        title: "Introduction",
        nodes: [Greeting],
      }),
    ],
  });
  assertEquals(legacyLetterTemplate.nodes.length, 1);

  const builtLetter = await buildDocument(letterTemplate, {
    recipientName: "Avery",
    households: 2,
  });

  assertEquals(builtLetter.id, "welcome");
  assertEquals(builtLetter.nodes.length, 1);

  const introSection = builtLetter.nodes[0];
  assertEquals(introSection.kind, "section");

  if (introSection.kind !== "section") {
    throw new Error("Expected section");
  }

  const greeting = introSection.children[0];
  assertEquals(greeting.kind, "paragraph");

  if (greeting.kind === "paragraph") {
    assertEquals(greeting.mode, "static");
    assertEquals(greeting.text, "Hello Avery,");
  }
});

test("dynamic paragraph components resolve through an AI client", async () => {
  const DynamicSummary = paragraph<LetterData>({
    id: "summary",
    generalPrompt(data, availableTokens) {
      return `Write a paragraph for ${data.households} households using ${availableTokens} tokens.`;
    },
    negativePrompt() {
      return "Do not invent facts.";
    },
  });

  const letterTemplate = defineDocument<LetterData>({
    id: "dynamic",
    title: "Dynamic Letter",
    nodes: [DynamicSummary],
  });

  const letter = await buildDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { availableTokens: 500, aiClient: new EchoAiClient() },
  );

  const paragraphNode = letter.nodes[0];
  assertEquals(paragraphNode.kind, "paragraph");

  if (paragraphNode.kind === "paragraph") {
    assertEquals(paragraphNode.mode, "dynamic");
    assertEquals(paragraphNode.prompts?.[0].kind, "general");
    assertEquals(paragraphNode.text?.includes("2 households"), true);
  }
});

test("dynamic paragraph components can render placeholders", async () => {
  const DynamicSummary = paragraph<LetterData>({
    id: "summary",
    placeholder(data) {
      return `Placeholder for ${data.recipientName}.`;
    },
    generalPrompt(data) {
      return `Write a paragraph for ${data.recipientName}.`;
    },
  });

  const letterTemplate = defineDocument<LetterData>({
    id: "preview",
    title: "Preview Letter",
    nodes: [DynamicSummary],
  });

  const letter = await buildDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { dynamicMode: "placeholder" },
  );

  const paragraphNode = letter.nodes[0];
  assertEquals(paragraphNode.kind, "paragraph");

  if (paragraphNode.kind === "paragraph") {
    assertEquals(paragraphNode.mode, "dynamic");
    assertEquals(paragraphNode.text, "Placeholder for Avery.");
    assertEquals(paragraphNode.prompts, undefined);
  }
});

test("dynamic paragraph components require an AI client", async () => {
  const letterTemplate = defineDocument<LetterData>({
    id: "dynamic",
    title: "Dynamic Letter",
    nodes: [
      paragraph({
        id: "summary",
        generalPrompt: () => "Write a summary.",
      }),
    ],
  });

  await assertRejects(
    () =>
      buildDocument(letterTemplate, {
        recipientName: "Avery",
        households: 2,
      }),
    Error,
    "requires an aiClient",
  );
});

test("image and graph node types resolve from component helpers", async () => {
  const letterTemplate = defineDocument<LetterData>({
    id: "visuals",
    title: "Visual Letter",
    nodes: [
      image({
        id: "signature",
        src: "assets/signature.png",
        alt: (data) => `Signature for ${data.recipientName}.`,
      }),
      graph({
        id: "households",
        graphType: "bar",
        data: (data) => ({
          labels: ["Households"],
          series: [{ label: data.recipientName, values: [data.households] }],
        }),
        caption: (data) => `Household count for ${data.recipientName}.`,
      }),
    ],
  });

  const letter = await buildDocument(letterTemplate, {
    recipientName: "Avery",
    households: 2,
  });

  const imageNode = letter.nodes[0];
  const graphNode = letter.nodes[1];

  assertEquals(imageNode.kind, "image");
  if (imageNode.kind === "image") {
    assertEquals(imageNode.mode, "static");
    assertEquals(imageNode.path, "assets/signature.png");
    assertEquals(imageNode.alt, "Signature for Avery.");
  }

  assertEquals(graphNode.kind, "graph");
  if (graphNode.kind === "graph") {
    assertEquals(graphNode.mode, "static");
    assertEquals(graphNode.graphType, "bar");
    assertEquals(graphNode.caption, "Household count for Avery.");
  }
});

test("mode is inferred from the options, not declared", async () => {
  const letterTemplate = defineDocument<LetterData>({
    id: "inference",
    title: "Inference",
    nodes: [
      // Local-resolution member present -> static.
      paragraph({ id: "p-static", render: () => "Text." }),
      image({ id: "i-static", src: "assets/logo.png" }),
      graph({ id: "g-static", data: () => ({ values: [1] }) }),
      // Prompts instead -> dynamic, without anyone saying so.
      paragraph({ id: "p-dynamic", generalPrompt: () => "Write it." }),
      image({ id: "i-dynamic", generalPrompt: () => "Draw it." }),
      graph({ id: "g-dynamic", generalPrompt: () => "Plot it." }),
    ],
  });

  const built = await buildDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { dynamicMode: "placeholder" },
  );

  assertEquals(
    built.nodes.map((node) => `${node.id}:${"mode" in node ? node.mode : ""}`),
    [
      "p-static:static",
      "i-static:static",
      "g-static:static",
      "p-dynamic:dynamic",
      "i-dynamic:dynamic",
      "g-dynamic:dynamic",
    ],
  );
});

test("dynamic graph components can render placeholders", async () => {
  const letterTemplate = defineDocument<LetterData>({
    id: "dynamic-graph",
    title: "Dynamic Graph Letter",
    nodes: [
      graph({
        id: "trend",
        placeholder: (data) => `Preview graph for ${data.recipientName}.`,
        generalPrompt: (data) => `Build a graph for ${data.households} households.`,
      }),
    ],
  });

  const letter = await buildDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { dynamicMode: "placeholder" },
  );
  const graphNode = letter.nodes[0];

  assertEquals(graphNode.kind, "graph");
  if (graphNode.kind === "graph") {
    assertEquals(graphNode.mode, "dynamic");
    assertEquals(graphNode.placeholder, "Preview graph for Avery.");
    assertEquals(graphNode.prompts, undefined);
  }
});
