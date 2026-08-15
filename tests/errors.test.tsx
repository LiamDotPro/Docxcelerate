import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import {
  createPublishData,
  Document,
  Graph,
  Image,
  Paragraph,
  Repeat,
  Section,
  TableOfContents,
  template,
  useSetPrompts,
  useState,
} from "docxcelerate/template";

/**
 * The failures this framework promises to make loud.
 *
 * Each one exists because the quiet version produces a document that is wrong
 * rather than a build that stops, and a wrong document is only discovered by
 * whoever receives it.
 */
interface Data {
  name: string;
  items: number[];
}

const data: Data = { name: "Avery", items: [1, 2] };

function build(children: unknown, options = {}) {
  return buildDocument(
    template<Data>(<Document id="d" title="D">{children as never}</Document>),
    data,
    options,
  );
}

function assertBuildRejects(children: unknown, message: string, options = {}) {
  return assertRejects(() => build(children, options), Error, message);
}

test("a template must be a single Document element", () => {
  const cases: Array<[unknown, string]> = [
    [<Paragraph id="p">Text.</Paragraph>, "a <paragraph> element"],
    [<Section id="s" title="S" />, "a <section> element"],
    ["just a string", "string"],
    [null, "null"],
    [[<Document id="a" title="A" />, <Document id="b" title="B" />], "object"],
  ];

  for (const [value, expected] of cases) {
    let message = "";
    try {
      template<Data>(value);
    } catch (error) {
      message = error instanceof Error ? error.message : String(error);
    }

    assertEquals([expected, message.includes(expected)], [expected, true]);
  }
});

test("a Document needs both an id and a title", () => {
  let message = "";
  try {
    template<Data>(<Document id="" title="" />);
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertEquals(message.includes("needs both an id and a title"), true);
});

test("a Document cannot be nested inside another element", async () => {
  await assertBuildRejects(
    <Section id="s" title="S">
      <Document id="inner" title="Inner" />
    </Section>,
    "can only be the root of a template",
  );
});

test("text where an element belongs says where it was found", async () => {
  await assertBuildRejects(
    <Section id="s" title="S">{"loose text"}</Section>,
    "Text only lives inside a <Paragraph>",
  );
});

test("an element inside a paragraph is rejected rather than flattened", async () => {
  await assertBuildRejects(
    <Paragraph id="p">
      {"before "}
      {<Paragraph id="inner">inner</Paragraph> as never}
    </Paragraph>,
    "put elements beside it, not inside it",
  );
});

test("a document element called directly explains itself", () => {
  let message = "";
  try {
    (Paragraph as unknown as (props: unknown) => unknown)({ id: "p" });
  } catch (error) {
    message = error instanceof Error ? error.message : String(error);
  }

  assertEquals(message.includes("cannot be called directly"), true);
});

test("an intrinsic tag has no meaning in a document", async () => {
  // <div> would silently disappear if JSX accepted it.
  await assertBuildRejects(
    <Section id="s" title="S">{{ kind: "not an element" } as never}</Section>,
    "not a document element",
  );
});

test("an image with neither a source nor a prompt says nothing describes it", async () => {
  await assertBuildRejects(
    <Image id="orphan" />,
    "neither a src nor any prompt",
  );
});

test("a graph with neither data nor a prompt says nothing describes it", async () => {
  await assertBuildRejects(
    <Graph id="orphan" />,
    "neither data nor any prompt",
  );
});

test("a node claiming both its content and a prompt is a contradiction", async () => {
  await assertBuildRejects(
    <Paragraph id="both" generalPrompt="Write it.">Already written.</Paragraph>,
    "supplies both its content and a prompt",
  );
  await assertBuildRejects(
    <Image id="both" src="a.png" generalPrompt="Draw it." />,
    "supplies both its content and a prompt",
  );
  await assertBuildRejects(
    <Graph id="both" data={{ values: [1] }} generalPrompt="Plot it." />,
    "supplies both its content and a prompt",
  );
});

test("a dynamic node without the AI client it needs names the method", async () => {
  const Prompted: Paragraph = () => {
    useSetPrompts({ generalPrompt: "Write it." });
    return <Paragraph id="p" />;
  };

  await assertBuildRejects(<Prompted />, "requires an aiClient");
  await assertBuildRejects(
    <Image id="i" generalPrompt="Draw it." />,
    "requires an aiClient.generateImage",
  );
  await assertBuildRejects(
    <Graph id="g" generalPrompt="Plot it." />,
    "requires an aiClient.generateGraph",
  );
});

test("a repeat over something that is not a collection says what it found", async () => {
  await assertBuildRejects(
    <Repeat over="name">
      <Paragraph>x</Paragraph>
    </Repeat>,
    "instead of a collection",
  );
  await assertBuildRejects(
    <Repeat over="missing">
      <Paragraph>x</Paragraph>
    </Repeat>,
    "found nothing instead of a collection",
  );
});

test("two explicit ids that collide name both positions", async () => {
  await assertBuildRejects(
    [
      <Paragraph id="same">One.</Paragraph>,
      <Section id="wrapper" title="W">
        <Paragraph id="same">Two.</Paragraph>
      </Section>,
    ],
    'Two nodes claim the id "same"',
  );
});

test("derived ids do not collide with each other", async () => {
  // Position supplies the id, so a tree of anonymous nodes still addresses
  // each one uniquely.
  const built = await build(
    <Section id="s" title="S">
      <Paragraph>One.</Paragraph>
      <Paragraph>Two.</Paragraph>
      <Section id="inner" title="I">
        <Paragraph>Three.</Paragraph>
        <TableOfContents />
      </Section>
    </Section>,
  );

  const collect = (nodes: typeof built.nodes): string[] =>
    nodes.flatMap((node) =>
      node.kind === "section" ? [node.id, ...collect(node.children)] : [node.id]
    );
  const ids = collect(built.nodes);

  assertEquals(ids.length, 6);
  assertEquals(new Set(ids).size, 6);
});

test("publishing refuses to read a value as a number, and says what to do", async () => {
  const Total: Paragraph = () => {
    const [state] = useState((input: Data) => ({
      doubled: (input.items as unknown as number) * 2,
    }));

    return <Paragraph id="total">{String(state.doubled)}</Paragraph>;
  };

  await assertRejects(
    () =>
      buildDocument(
        template<Data>(<Document id="d" title="D"><Total /></Document>),
        createPublishData() as Data,
        { branchMode: "publish", deriverMode: "preserve" },
      ),
    Error,
    "use a deriver",
  );
});

test("publishing refuses to iterate, and names the Repeat to write", async () => {
  const List: Section = () => {
    const [state] = useState((input: Data) => ({ items: input.items }));

    return (
      <Section id="list" title="List">
        {state.items.map((item, index) => <Paragraph id={`i-${index}`}>{item}</Paragraph>)}
      </Section>
    );
  };

  await assertRejects(
    () =>
      buildDocument(
        template<Data>(<Document id="d" title="D"><List /></Document>),
        createPublishData() as Data,
        { branchMode: "publish", deriverMode: "preserve" },
      ),
    Error,
    '<Repeat over="items">',
  );
});

test("publishing lets a value be interpolated as text", async () => {
  const Greeting: Paragraph = () => {
    const [state] = useState((input: Data) => ({ name: input.name }));

    return <Paragraph id="greeting">Hello {state.name}.</Paragraph>;
  };

  const built = await buildDocument(
    template<Data>(<Document id="d" title="D"><Greeting /></Document>),
    createPublishData() as Data,
    { branchMode: "publish", deriverMode: "preserve" },
  );

  const node = built.nodes[0];
  assertEquals(node.kind === "paragraph" && node.text, "Hello {{data.name}}.");
});
