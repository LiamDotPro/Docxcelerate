import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import { EchoAiClient } from "docxcelerate";
import {
  branch,
  buildDocument,
  compare,
  createPublishData,
  dataPath,
  Document,
  Graph,
  Image,
  literal,
  Paragraph,
  Repeat,
  Section,
  template,
  truthy,
  useFormat,
  usePlaceholderData,
  useSetPlaceholders,
  useSetPrompts,
  useShared,
  useState,
} from "docxcelerate/template";
import type { DocumentNode, ParagraphNode } from "docxcelerate";

interface TenancyData {
  recipientName: string;
  balanceDue: number;
  city: string;
  visits: Array<{ label: string }>;
}

const tenancy: TenancyData = {
  recipientName: "Avery",
  balanceDue: 240,
  city: "Leeds",
  visits: [{ label: "First" }, { label: "Second" }],
};

function paragraphs(nodes: DocumentNode[]): ParagraphNode[] {
  return nodes.flatMap((node) => {
    if (node.kind === "section" || node.kind === "repeat") {
      return paragraphs(node.children);
    }

    return node.kind === "paragraph" ? [node] : [];
  });
}

test("a component reads state seeded from data and branches on it", async () => {
  const Balance: Paragraph = () => {
    const [state] = useState((data: TenancyData) => ({
      name: data.recipientName,
      settled: data.balanceDue === 0,
    }));

    if (state.settled) {
      return <Paragraph id="settled">Nothing outstanding, {state.name}.</Paragraph>;
    }

    return <Paragraph id="arrears">A balance remains for {state.name}.</Paragraph>;
  };

  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Section id="opening" title="Opening">
          <Balance />
        </Section>
      </Document>,
    ),
    tenancy,
  );

  assertEquals(paragraphs(built.nodes).map((node) => node.id), ["arrears"]);
  assertEquals(paragraphs(built.nodes)[0].text, "A balance remains for Avery.");
});

test("data reaches a child component as props from its parent", async () => {
  const Arrears: Paragraph<{ amount: number }> = ({ amount }) => {
    const { currency } = useFormat("en-GB");

    return <Paragraph id="arrears">You owe {currency(amount)}.</Paragraph>;
  };

  const Opening: Section = () => {
    const [state] = useState((data: TenancyData) => ({ due: data.balanceDue }));

    return (
      <Section id="opening" title="Opening">
        <Arrears amount={state.due} />
      </Section>
    );
  };

  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Opening />
      </Document>,
    ),
    tenancy,
  );

  assertEquals(paragraphs(built.nodes)[0].text, "You owe £240.00.");
});

test("useSetPrompts makes the yielded node dynamic without declaring a mode", async () => {
  const Summary: Paragraph = () => {
    const [state] = useState((data: TenancyData) => ({ city: data.city }));

    useSetPrompts({
      systemPrompt: "You are a tenancy officer.",
      generalPrompt: `Summarise the account for ${state.city}.`,
      negativePrompt: "Do not invent figures.",
    });

    return <Paragraph id="summary" />;
  };

  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Summary />
      </Document>,
    ),
    tenancy,
    { aiClient: new EchoAiClient() },
  );

  const [summary] = paragraphs(built.nodes);
  assertEquals(summary.mode, "dynamic");
  assertEquals(summary.prompts?.map((prompt) => prompt.kind), ["system", "general", "negative"]);
  assertEquals(summary.text?.includes("Leeds"), true);
});

test("an arm that supplies its own text keeps it, prompts set before the branch or not", async () => {
  // Hooks run in call order, so a component sets prompts before it branches.
  // The arm that says what it says is still static.
  const Balance: Paragraph = () => {
    const [state] = useState((data: TenancyData) => ({
      name: data.recipientName,
      settled: data.balanceDue === 0,
    }));

    useSetPrompts({ generalPrompt: `Explain the balance to ${state.name}.` });

    if (state.settled) {
      return <Paragraph id="settled">Nothing outstanding, {state.name}.</Paragraph>;
    }

    return <Paragraph id="arrears" />;
  };

  const build = (balanceDue: number) =>
    buildDocument(
      template<TenancyData>(
        <Document id="tenancy" title="Tenancy">
          <Balance />
        </Document>,
      ),
      { ...tenancy, balanceDue },
      { aiClient: new EchoAiClient() },
    );

  const [cleared] = paragraphs((await build(0)).nodes);
  assertEquals(cleared.mode, "static");
  assertEquals(cleared.text, "Nothing outstanding, Avery.");

  const [owing] = paragraphs((await build(240)).nodes);
  assertEquals(owing.mode, "dynamic");
  assertEquals(owing.prompts?.[0].text, "Explain the balance to Avery.");
});

test("one node claiming to be both written and generated is a contradiction", async () => {
  await assertRejects(
    () =>
      buildDocument(
        template<TenancyData>(
          <Document id="tenancy" title="Tenancy">
            <Paragraph id="both" generalPrompt="Write it.">Already written.</Paragraph>
          </Document>,
        ),
        tenancy,
      ),
    Error,
    "supplies both its content and a prompt",
  );
});

test("placeholders stand in where a document is previewed rather than written", async () => {
  const Summary: Paragraph = () => {
    const placeholder = usePlaceholderData();

    useSetPrompts({ generalPrompt: "Summarise the account." });
    useSetPlaceholders(`A summary for ${placeholder.city()}.`);

    return <Paragraph id="summary" />;
  };

  const build = () =>
    buildDocument(
      template<TenancyData>(
        <Document id="tenancy" title="Tenancy">
          <Summary />
        </Document>,
      ),
      tenancy,
      { dynamicMode: "placeholder" },
    );

  const first = paragraphs((await build()).nodes)[0];
  const second = paragraphs((await build()).nodes)[0];

  assertEquals(first.mode, "dynamic");
  assertEquals(first.prompts, undefined);
  assertEquals(first.text?.startsWith("A summary for "), true);
  // Seeded from where the component sits, so a preview can be proofread.
  assertEquals(first.text, second.text);
});

test("useShared carries a value from one component to the next", async () => {
  const Counted: Paragraph<{ label: string }> = ({ label }) => {
    const [seen, setSeen] = useShared<string[]>("mentioned", []);
    setSeen([...seen, label]);

    return <Paragraph id={`seen-${label}`}>Seen so far: {seen.length}.</Paragraph>;
  };

  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Counted label="a" />
        <Counted label="b" />
        <Counted label="c" />
      </Document>,
    ),
    tenancy,
  );

  assertEquals(
    paragraphs(built.nodes).map((node) => node.text),
    ["Seen so far: 0.", "Seen so far: 1.", "Seen so far: 2."],
  );
});

test("a branch takes one arm when the data is real", async () => {
  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        {branch(
          compare({ type: "ref", ref: dataPath("balanceDue") }, "gt", literal(0)),
          () => <Paragraph id="arrears">Arrears.</Paragraph>,
          () => <Paragraph id="settled">Settled.</Paragraph>,
        )}
      </Document>,
    ),
    tenancy,
  );

  assertEquals(paragraphs(built.nodes).map((node) => node.id), ["arrears"]);
});

test("publishing keeps both arms, each under the condition that selects it", async () => {
  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        {branch(
          truthy(dataPath("hasArrears")),
          () => <Paragraph id="arrears">Arrears.</Paragraph>,
          () => <Paragraph id="settled">Settled.</Paragraph>,
        )}
      </Document>,
    ),
    tenancy,
    { branchMode: "publish", deriverMode: "preserve" },
  );

  assertEquals(paragraphs(built.nodes).map((node) => node.id), ["arrears", "settled"]);
  assertEquals(built.nodes[0].when, { type: "truthy", ref: dataPath("hasArrears") });
  // Inverted into the opposite shape rather than wrapped, so an engine that
  // only knows the original two conditions still understands it.
  assertEquals(built.nodes[1].when, { type: "not", ref: dataPath("hasArrears") });
});

test("nested branches are capped rather than published in silence", async () => {
  const nest = (depth: number, at = "root"): ReturnType<typeof branch> =>
    branch(
      truthy(dataPath(`flag${depth}`)),
      () =>
        depth === 0
          ? <Paragraph id={`${at}-yes`}>Yes.</Paragraph>
          : nest(depth - 1, `${at}-yes`),
      () =>
        depth === 0
          ? <Paragraph id={`${at}-no`}>No.</Paragraph>
          : nest(depth - 1, `${at}-no`),
    );

  await assertRejects(
    () =>
      buildDocument(
        template<TenancyData>(
          <Document id="tenancy" title="Tenancy">{nest(6)}</Document>,
        ),
        tenancy,
        { branchMode: "publish", deriverMode: "preserve", branchLimit: 8 },
      ),
    Error,
    "passed 8 branches",
  );
});

test("a repeat is walked with real data and published as a loop without it", async () => {
  const tree = template<TenancyData>(
    <Document id="tenancy" title="Tenancy">
      <Repeat over="visits" as="visit">
        <Paragraph id="visit">Visit {"{{ctx.visit.label}}"}.</Paragraph>
      </Repeat>
    </Document>,
  );

  const walked = await buildDocument(tree, tenancy);
  assertEquals(
    paragraphs(walked.nodes).map((node) => node.text),
    ["Visit First.", "Visit Second."],
  );
  // Suffixed per pass, the same way the engine suffixes a published loop, so an
  // id means the same thing in a preview and in a written document.
  assertEquals(paragraphs(walked.nodes).map((node) => node.id), ["visit-0", "visit-1"]);

  const published = await buildDocument(tree, tenancy, {
    branchMode: "publish",
    deriverMode: "preserve",
  });

  assertEquals(published.nodes[0].kind, "repeat");
  if (published.nodes[0].kind === "repeat") {
    assertEquals(published.nodes[0].source.path, "visits");
    assertEquals(published.nodes[0].children.length, 1);
  }
});

test("two nodes claiming one id is reported, not resolved by whoever came last", async () => {
  await assertRejects(
    () =>
      buildDocument(
        template<TenancyData>(
          <Document id="tenancy" title="Tenancy">
            <Paragraph id="opening">One.</Paragraph>
            <Paragraph id="opening">Two.</Paragraph>
          </Document>,
        ),
        tenancy,
      ),
    Error,
    'Two nodes claim the id "opening"',
  );
});

test("a node without an id takes one from where it sits", async () => {
  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Section id="opening" title="Opening">
          <Paragraph>First.</Paragraph>
          <Paragraph>Second.</Paragraph>
        </Section>
      </Document>,
    ),
    tenancy,
  );

  const ids = paragraphs(built.nodes).map((node) => node.id);
  assertEquals(ids.length, 2);
  assertEquals(new Set(ids).size, 2);
});

test("images and graphs infer their mode the same way paragraphs do", async () => {
  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Image id="signature" src="assets/signature.png" alt="A signature" />
        <Image id="scene" generalPrompt="Draw the property." />
        <Graph id="totals" graphType="bar" data={{ labels: ["a"], values: [1] }} />
        <Graph id="trend" generalPrompt="Plot the balance." />
      </Document>,
    ),
    tenancy,
    { dynamicMode: "placeholder" },
  );

  assertEquals(
    built.nodes.map((node) => `${node.id}:${"mode" in node ? node.mode : ""}`),
    ["signature:static", "scene:dynamic", "totals:static", "trend:dynamic"],
  );
});

test("publishing interpolates a value it does not know yet", async () => {
  const Greeting: Paragraph = () => {
    const [state] = useState((data: TenancyData) => ({ name: data.recipientName }));

    return <Paragraph id="greeting">Hello {state.name},</Paragraph>;
  };

  const built = await buildDocument(
    template<TenancyData>(
      <Document id="tenancy" title="Tenancy">
        <Greeting />
      </Document>,
    ),
    createPublishData() as TenancyData,
    { branchMode: "publish", deriverMode: "preserve" },
  );

  assertEquals(paragraphs(built.nodes)[0].text, "Hello {{data.recipientName}},");
});

test("computing on a value nobody has yet is refused, with what to do instead", async () => {
  const Total: Paragraph = () => {
    const { currency } = useFormat("en-GB");
    const [state] = useState((data: TenancyData) => ({ due: data.balanceDue }));

    return <Paragraph id="total">You owe {currency(state.due)}.</Paragraph>;
  };

  await assertRejects(
    () =>
      buildDocument(
        template<TenancyData>(
          <Document id="tenancy" title="Tenancy">
            <Total />
          </Document>,
        ),
        createPublishData() as TenancyData,
        { branchMode: "publish", deriverMode: "preserve" },
      ),
    Error,
    "use a deriver",
  );
});

test("iterating a collection nobody has yet points at the loop that survives", async () => {
  const Visits: Section = () => {
    const [state] = useState((data: TenancyData) => ({ visits: data.visits }));

    return (
      <Section id="visits" title="Visits">
        {state.visits.map((visit, index) => (
          <Paragraph id={`visit-${index}`}>{visit.label}</Paragraph>
        ))}
      </Section>
    );
  };

  await assertRejects(
    () =>
      buildDocument(
        template<TenancyData>(
          <Document id="tenancy" title="Tenancy">
            <Visits />
          </Document>,
        ),
        createPublishData() as TenancyData,
        { branchMode: "publish", deriverMode: "preserve" },
      ),
    Error,
    '<Repeat over="visits">',
  );
});

test("hooks called after an await say so instead of landing on another component", async () => {
  const Late: Paragraph = async () => {
    await Promise.resolve();
    useSetPrompts({ generalPrompt: "Too late." });

    return <Paragraph id="late" />;
  };

  await assertRejects(
    () =>
      buildDocument(
        template<TenancyData>(
          <Document id="tenancy" title="Tenancy">
            <Late />
          </Document>,
        ),
        tenancy,
      ),
    Error,
    "after an await",
  );
});
