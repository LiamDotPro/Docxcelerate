import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import {
  buildLetterDocument,
  defineLetter,
  dynamicGraph,
  dynamicParagraph,
  EchoAiClient,
  letter,
  type Node,
  type NodeComponent,
  section,
  staticGraph,
  staticImage,
  staticParagraph,
} from "docxcelerate";
import type { ParagraphNode } from "../src/domain/types.ts";

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
  const letterTemplate = letter<LetterData>(
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
          staticParagraph({
            id: "introText",
            render: () => "A simple static paragraph.",
          }),
        ],
      ),
    ],
  );
  const legacyLetterTemplate = defineLetter<LetterData>({
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

  const builtLetter = await buildLetterDocument(letterTemplate, {
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
  const DynamicSummary = dynamicParagraph<LetterData>({
    id: "summary",
    generalPrompt(data, availableTokens) {
      return `Write a paragraph for ${data.households} households using ${availableTokens} tokens.`;
    },
    negativePrompt() {
      return "Do not invent facts.";
    },
  });

  const letterTemplate = defineLetter<LetterData>({
    id: "dynamic",
    title: "Dynamic Letter",
    nodes: [DynamicSummary],
  });

  const letter = await buildLetterDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { availableTokens: 500, aiClient: new EchoAiClient() },
  );

  const paragraph = letter.nodes[0];
  assertEquals(paragraph.kind, "paragraph");

  if (paragraph.kind === "paragraph") {
    assertEquals(paragraph.mode, "dynamic");
    assertEquals(paragraph.prompts?.[0].kind, "general");
    assertEquals(paragraph.text?.includes("2 households"), true);
  }
});

test("dynamic paragraph components can render placeholders", async () => {
  const DynamicSummary = dynamicParagraph<LetterData>({
    id: "summary",
    placeholder(data) {
      return `Placeholder for ${data.recipientName}.`;
    },
    generalPrompt(data) {
      return `Write a paragraph for ${data.recipientName}.`;
    },
  });

  const letterTemplate = defineLetter<LetterData>({
    id: "preview",
    title: "Preview Letter",
    nodes: [DynamicSummary],
  });

  const letter = await buildLetterDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { dynamicMode: "placeholder" },
  );

  const paragraph = letter.nodes[0];
  assertEquals(paragraph.kind, "paragraph");

  if (paragraph.kind === "paragraph") {
    assertEquals(paragraph.mode, "dynamic");
    assertEquals(paragraph.text, "Placeholder for Avery.");
    assertEquals(paragraph.prompts, undefined);
  }
});

test("dynamic paragraph components require an AI client", async () => {
  const letterTemplate = defineLetter<LetterData>({
    id: "dynamic",
    title: "Dynamic Letter",
    nodes: [
      dynamicParagraph({
        id: "summary",
        generalPrompt: () => "Write a summary.",
      }),
    ],
  });

  await assertRejects(
    () =>
      buildLetterDocument(letterTemplate, {
        recipientName: "Avery",
        households: 2,
      }),
    Error,
    "requires an aiClient",
  );
});

test("image and graph node types resolve from component helpers", async () => {
  const letterTemplate = defineLetter<LetterData>({
    id: "visuals",
    title: "Visual Letter",
    nodes: [
      staticImage({
        id: "signature",
        src: "assets/signature.png",
        alt: (data) => `Signature for ${data.recipientName}.`,
      }),
      staticGraph({
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

  const letter = await buildLetterDocument(letterTemplate, {
    recipientName: "Avery",
    households: 2,
  });

  const image = letter.nodes[0];
  const graph = letter.nodes[1];

  assertEquals(image.kind, "image");
  if (image.kind === "image") {
    assertEquals(image.mode, "static");
    assertEquals(image.path, "assets/signature.png");
    assertEquals(image.alt, "Signature for Avery.");
  }

  assertEquals(graph.kind, "graph");
  if (graph.kind === "graph") {
    assertEquals(graph.mode, "static");
    assertEquals(graph.graphType, "bar");
    assertEquals(graph.caption, "Household count for Avery.");
  }
});

test("dynamic graph components can render placeholders", async () => {
  const letterTemplate = defineLetter<LetterData>({
    id: "dynamic-graph",
    title: "Dynamic Graph Letter",
    nodes: [
      dynamicGraph({
        id: "trend",
        placeholder: (data) => `Preview graph for ${data.recipientName}.`,
        generalPrompt: (data) => `Build a graph for ${data.households} households.`,
      }),
    ],
  });

  const letter = await buildLetterDocument(
    letterTemplate,
    { recipientName: "Avery", households: 2 },
    { dynamicMode: "placeholder" },
  );
  const graph = letter.nodes[0];

  assertEquals(graph.kind, "graph");
  if (graph.kind === "graph") {
    assertEquals(graph.mode, "dynamic");
    assertEquals(graph.placeholder, "Preview graph for Avery.");
    assertEquals(graph.prompts, undefined);
  }
});
