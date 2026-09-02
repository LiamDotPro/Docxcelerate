import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import { buildDocument, type DocumentModel, type DocumentNode } from "docxcelerate";
import {
  ctxPath,
  Document,
  Paragraph,
  Section,
  template,
  useAvailableTokens,
  deriver,
  useDeriver,
  useFormat,
  usePlaceholderData,
  useSetPlaceholders,
  useSetPrompts,
  useShared,
  useState,
} from "docxcelerate/template";

interface Data {
  name: string;
  amount: number;
  items: string[];
}

const data: Data = { name: "Avery", amount: 1234.5, items: ["one", "two", "three"] };

function wrap(children: unknown) {
  return template<Data>(
    <Document id="d" title="D">{children as never}</Document>,
  );
}

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    if (node.kind === "section" || node.kind === "repeat") return texts(node.children);
    return node.kind === "paragraph" ? [node.text ?? ""] : [];
  });
}

function build(children: unknown, options = {}): Promise<DocumentModel> {
  return buildDocument(wrap(children), data, options);
}

test("the state initializer runs once and receives the build's data", async () => {
  let calls = 0;

  const Probe: Paragraph = () => {
    const [state] = useState((input: Data) => {
      calls += 1;
      return { name: input.name };
    });

    return <Paragraph id="probe">{state.name}</Paragraph>;
  };

  assertEquals(texts((await build(<Probe />)).nodes), ["Avery"]);
  assertEquals(calls, 1);
});

test("a plain value works as an initializer, without being called", async () => {
  const Probe: Paragraph = () => {
    const [state] = useState({ label: "fixed" });

    return <Paragraph id="probe">{state.label}</Paragraph>;
  };

  assertEquals(texts((await build(<Probe />)).nodes), ["fixed"]);
});

test("the state setter is visible to a later read in the same component", async () => {
  const Probe: Paragraph = () => {
    const [, setCount] = useState(1);
    setCount((previous) => previous + 10);
    // Reading through the hook again is not how a component works — the
    // destructured value is a snapshot, as it is in React. What the setter
    // moves is the cell, which is what `useShared` consumers observe.
    const [, setAgain] = useState(99);
    setAgain(0);

    return <Paragraph id="probe">ok</Paragraph>;
  };

  assertEquals(texts((await build(<Probe />)).nodes), ["ok"]);
});

test("sibling components of one type keep separate state", async () => {
  const Counter: Paragraph<{ start: number }> = ({ start }) => {
    const [state] = useState(() => ({ value: start * 2 }));

    return <Paragraph>{String(state.value)}</Paragraph>;
  };

  const built = await build([
    <Counter start={1} />,
    <Counter start={2} />,
    <Counter start={3} />,
  ]);

  assertEquals(texts(built.nodes), ["2", "4", "6"]);
});

test("useShared initialises once and carries forward in document order", async () => {
  let initialised = 0;

  const Tally: Paragraph<{ add: number }> = ({ add }) => {
    const [total, setTotal] = useShared("total", () => {
      initialised += 1;
      return 0;
    });
    setTotal(total + add);

    return <Paragraph>seen {String(total)}</Paragraph>;
  };

  const built = await build([<Tally add={5} />, <Tally add={7} />, <Tally add={1} />]);

  assertEquals(texts(built.nodes), ["seen 0", "seen 5", "seen 12"]);
  assertEquals(initialised, 1);
});

test("useShared is scoped to one build, not to the module", async () => {
  const Tally: Paragraph = () => {
    const [count, setCount] = useShared("count", 0);
    setCount(count + 1);

    return <Paragraph id="tally">{String(count)}</Paragraph>;
  };

  assertEquals(texts((await build(<Tally />)).nodes), ["0"]);
  // A second build starts over; state that leaked between builds would make
  // one document depend on how many were written before it.
  assertEquals(texts((await build(<Tally />)).nodes), ["0"]);
});

test("useAvailableTokens reports the budget the build was given", async () => {
  const Probe: Paragraph = () => <Paragraph id="probe">{String(useAvailableTokens())}</Paragraph>;

  assertEquals(texts((await build(<Probe />)).nodes), ["2000"]);
  assertEquals(texts((await build(<Probe />, { availableTokens: 512 })).nodes), ["512"]);
});

test("useFormat formats against the build's locale", async () => {
  const Probe: Paragraph = () => {
    const { currency, number, list, plural } = useFormat();

    return (
      <Paragraph id="probe">
        {currency(1234.5)} | {number(0.25, { style: "percent" })} |{" "}
        {list(["a", "b", "c"])} | {plural(1, "visit")} | {plural(3, "visit")}
      </Paragraph>
    );
  };

  assertEquals(texts((await build(<Probe />)).nodes), [
    "£1,234.50 | 25% | a, b and c | visit | visits",
  ]);
});

test("a locale passed to useFormat overrides the build's", async () => {
  const Probe: Paragraph = () => {
    const { currency } = useFormat("de-DE");

    return <Paragraph id="probe">{currency(1234.5, "EUR")}</Paragraph>;
  };

  const [text] = texts((await build(<Probe />)).nodes);
  assertEquals(text.includes("1.234,50"), true);
});

test("the build locale reaches useFormat", async () => {
  const Probe: Paragraph = () => {
    const { number } = useFormat();

    return <Paragraph id="probe">{number(1234.5)}</Paragraph>;
  };

  const [text] = texts((await build(<Probe />, { locale: "de-DE" })).nodes);
  assertEquals(text, "1.234,5");
});

const double = deriver({ name: "double", run: (value: number) => value * 2 });

test("useDeriver runs a deriver and hands back what it produced", async () => {
  // A deriver may be async, so the component that awaits one becomes async —
  // and its hooks have to be reached before that await.
  const Probe: Paragraph = async () => {
    const doubled = await useDeriver(double, [21]);

    return <Paragraph id="probe">{`${doubled}`}</Paragraph>;
  };

  assertEquals(texts((await build(<Probe />)).nodes), ["42"]);
});

test("what a deriver produced is also readable as a token", async () => {
  const Probe: Paragraph = async () => {
    await useDeriver(double, [21]);

    return <Paragraph id="probe">{"{{derived.double}}"}</Paragraph>;
  };

  assertEquals(texts((await build(<Probe />)).nodes), ["42"]);
});

test("two calls to one deriver from one component get their own keys", async () => {
  const Probe: Paragraph = async () => {
    await Promise.all([useDeriver(double, [21]), useDeriver(double, [50])]);

    return (
      <Paragraph id="probe">{"{{derived.double}}"} and {"{{derived.double-2}}"}</Paragraph>
    );
  };

  assertEquals(texts((await build(<Probe />)).nodes), ["42 and 100"]);
});

test("placeholder data is stable for a node and differs between nodes", async () => {
  const Probe: Paragraph<{ id: string }> = ({ id }) => {
    const placeholder = usePlaceholderData();

    useSetPrompts({ generalPrompt: "Write it." });
    useSetPlaceholders(`${placeholder.name()} of ${placeholder.city()}`);

    return <Paragraph id={id} />;
  };

  const once = texts((await build([<Probe id="a" />, <Probe id="b" />], {
    dynamicMode: "placeholder",
  })).nodes);
  const twice = texts((await build([<Probe id="a" />, <Probe id="b" />], {
    dynamicMode: "placeholder",
  })).nodes);

  // The same node reads the same every time, so a preview can be proofread.
  assertEquals(once, twice);
  // Two nodes do not read as the same person, which would make a preview
  // look like a bug.
  assertEquals(once[0] === once[1], false);
});

test("placeholder generators produce the kinds of value they claim", async () => {
  const Probe: Paragraph = () => {
    const placeholder = usePlaceholderData();

    useSetPrompts({ generalPrompt: "Write it." });
    useSetPlaceholders(
      [
        placeholder.sentence(4),
        placeholder.paragraph(2),
        placeholder.currency(12.5),
        placeholder.date(0),
        placeholder.pick(["only"]),
      ].join(" ~ "),
    );

    return <Paragraph id="probe" />;
  };

  const [text] = texts((await build(<Probe />, { dynamicMode: "placeholder" })).nodes);
  const [sentence, paragraph, currency, date, picked] = text.split(" ~ ");

  assertEquals(sentence.endsWith("."), true);
  assertEquals(paragraph.split(".").length > 2, true);
  assertEquals(currency, "£12.50");
  assertEquals(date, "15 January 2024");
  assertEquals(picked, "only");
});

test("a hook called outside a build says so", async () => {
  await assertRejects(
    // deno-lint-ignore require-await -- assertRejects needs a promise; the call throws.
    async () => {
      useAvailableTokens();
    },
    Error,
    "outside a document build",
  );

  await assertRejects(
    // deno-lint-ignore require-await -- assertRejects needs a promise; the call throws.
    async () => {
      useSetPrompts({ generalPrompt: "nowhere" });
    },
    Error,
    "outside a component",
  );
});

test("prompts set by a shared hook reach the node the caller returns", async () => {
  // The reason useSetPrompts sets rather than returns: a hook can apply house
  // style to a node it does not own.
  function useHouseVoice() {
    useSetPrompts({ systemPrompt: "House voice." });
  }

  const Probe: Paragraph = () => {
    useHouseVoice();
    useSetPrompts({ generalPrompt: "Say something." });

    return <Paragraph id="probe" />;
  };

  const built = await build(<Probe />, { dynamicMode: "resolve", aiClient: { generateParagraph: (request: { prompt: string }) => request.prompt } });
  const [text] = texts(built.nodes);

  assertEquals(text.includes("SYSTEM: House voice."), true);
  assertEquals(text.includes("GENERAL: Say something."), true);
});

test("a prop overrides what a shared hook set", async () => {
  const Probe: Paragraph = () => {
    useSetPrompts({ systemPrompt: "Hook voice.", generalPrompt: "Say something." });

    return <Paragraph id="probe" systemPrompt="Prop voice." />;
  };

  const built = await build(<Probe />, { aiClient: { generateParagraph: (request: { prompt: string }) => request.prompt } });
  const [text] = texts(built.nodes);

  assertEquals(text.includes("Prop voice."), true);
  assertEquals(text.includes("Hook voice."), false);
});

test("prompts do not leak from one component to the next", async () => {
  const Dynamic: Paragraph = () => {
    useSetPrompts({ generalPrompt: "Generated." });

    return <Paragraph id="dynamic" />;
  };

  const Static: Paragraph = () => <Paragraph id="static">Written.</Paragraph>;

  const built = await build([<Dynamic />, <Static />], { dynamicMode: "placeholder" });
  const [dynamic, statically] = built.nodes;

  assertEquals(dynamic.kind === "paragraph" && dynamic.mode, "dynamic");
  assertEquals(statically.kind === "paragraph" && statically.mode, "static");
  assertEquals(texts(built.nodes)[1], "Written.");
});

test("a component nested inside a section still reaches the build", async () => {
  const Probe: Paragraph = () => {
    const [state] = useState((input: Data) => ({ name: input.name }));

    return <Paragraph id="deep">{state.name}</Paragraph>;
  };

  const Outer: Section = () => (
    <Section id="outer" title="Outer">
      <Section id="inner" title="Inner">
        <Probe />
      </Section>
    </Section>
  );

  assertEquals(texts((await build(<Outer />)).nodes), ["Avery"]);
});

test("ctxPath names a reference the same way a token spells it", () => {
  assertEquals(ctxPath("visit.label"), { scope: "ctx", path: "visit.label" });
});
