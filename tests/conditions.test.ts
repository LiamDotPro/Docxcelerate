import { test } from "node:test";
import { assertEquals } from "./assert.ts";
import {
  type Condition,
  evaluateCondition,
  InMemoryDataProvider,
  invertCondition,
  type RuntimeState,
} from "docxcelerate";

/**
 * Conditions are the half of a published branch that the engine evaluates, so
 * they are tested against the shapes the branch compiler emits rather than
 * against the type alone.
 */
function stateWith(data: Record<string, unknown>, derived: Record<string, unknown> = {}) {
  const state: RuntimeState = {
    ctx: { ...data },
    derived,
    dataProvider: new InMemoryDataProvider({ ...data }),
    aiClient: { generateParagraph: () => "" },
  };

  return state;
}

const data = { balanceDue: 240, name: "Avery", settled: false, zero: 0, blank: "" };
const ref = { scope: "data", path: "balanceDue" } as const;
const literal = (value: string | number | boolean) => ({ type: "literal", value } as const);

test("an absent condition includes the node", async () => {
  assertEquals(await evaluateCondition(undefined, stateWith(data)), true);
});

test("truthy and not read the value, including the falsy ones", async () => {
  const state = stateWith(data);
  const at = (path: string) => ({ scope: "data", path } as const);

  assertEquals(await evaluateCondition({ type: "truthy", ref: at("balanceDue") }, state), true);
  assertEquals(await evaluateCondition({ type: "truthy", ref: at("zero") }, state), false);
  assertEquals(await evaluateCondition({ type: "truthy", ref: at("blank") }, state), false);
  assertEquals(await evaluateCondition({ type: "truthy", ref: at("missing") }, state), false);
  assertEquals(await evaluateCondition({ type: "not", ref: at("zero") }, state), true);
  assertEquals(await evaluateCondition({ type: "not", ref: at("balanceDue") }, state), false);
});

test("comparison operators order numbers", async () => {
  const state = stateWith(data);
  const compare = (operator: string, value: number): Condition => ({
    type: "compare",
    operator: operator as "gt",
    left: { type: "ref", ref },
    right: literal(value),
  });

  assertEquals(await evaluateCondition(compare("gt", 0), state), true);
  assertEquals(await evaluateCondition(compare("gt", 240), state), false);
  assertEquals(await evaluateCondition(compare("gte", 240), state), true);
  assertEquals(await evaluateCondition(compare("lt", 500), state), true);
  assertEquals(await evaluateCondition(compare("lte", 240), state), true);
  assertEquals(await evaluateCondition(compare("eq", 240), state), true);
  assertEquals(await evaluateCondition(compare("ne", 240), state), false);
});

test("equality does not coerce across types", async () => {
  const state = stateWith(data);
  const compare = (value: string | number): Condition => ({
    type: "compare",
    operator: "eq",
    left: { type: "ref", ref },
    right: literal(value),
  });

  assertEquals(await evaluateCondition(compare(240), state), true);
  // "240" == 240 in JavaScript. A document that says a balance is settled
  // because a string looked like a number is the failure this rules out.
  assertEquals(await evaluateCondition(compare("240"), state), false);
});

test("ordering two values of different kinds is false rather than a coerced guess", async () => {
  const state = stateWith(data);
  const condition: Condition = {
    type: "compare",
    operator: "gt",
    left: { type: "ref", ref: { scope: "data", path: "name" } },
    right: literal(0),
  };

  assertEquals(await evaluateCondition(condition, state), false);
});

test("strings order by collation", async () => {
  const state = stateWith(data);
  const compare = (operator: "lt" | "gt"): Condition => ({
    type: "compare",
    operator,
    left: { type: "ref", ref: { scope: "data", path: "name" } },
    right: literal("Beatrix"),
  });

  assertEquals(await evaluateCondition(compare("lt"), state), true);
  assertEquals(await evaluateCondition(compare("gt"), state), false);
});

test("and and or combine their members", async () => {
  const state = stateWith(data);
  const yes: Condition = { type: "truthy", ref };
  const no: Condition = { type: "not", ref };

  assertEquals(await evaluateCondition({ type: "and", conditions: [yes, yes] }, state), true);
  assertEquals(await evaluateCondition({ type: "and", conditions: [yes, no] }, state), false);
  assertEquals(await evaluateCondition({ type: "or", conditions: [no, yes] }, state), true);
  assertEquals(await evaluateCondition({ type: "or", conditions: [no, no] }, state), false);
  // An empty `and` holds vacuously; an empty `or` has nothing to satisfy it.
  assertEquals(await evaluateCondition({ type: "and", conditions: [] }, state), true);
  assertEquals(await evaluateCondition({ type: "or", conditions: [] }, state), false);
});

test("negate flips whatever it wraps", async () => {
  const state = stateWith(data);
  const inner: Condition = { type: "truthy", ref };

  assertEquals(await evaluateCondition({ type: "negate", condition: inner }, state), false);
});

test("conditions read derived and ctx as well as data", async () => {
  const state = stateWith(data, { total: 12 });

  assertEquals(
    await evaluateCondition(
      { type: "truthy", ref: { scope: "derived", path: "total" } },
      state,
    ),
    true,
  );
  assertEquals(
    await evaluateCondition(
      { type: "truthy", ref: { scope: "ctx", path: "name" } },
      state,
    ),
    true,
  );
});

/**
 * Both arms of a published branch are stored, one under a condition and one
 * under its inverse. If inversion disagreed with evaluation, a recipient could
 * match both arms or neither — so the two are checked against each other.
 */
test("every condition and its inverse disagree, whatever the data", async () => {
  const state = stateWith(data, { total: 0 });
  const at = (path: string) => ({ scope: "data", path } as const);
  const conditions: Condition[] = [
    { type: "truthy", ref: at("balanceDue") },
    { type: "truthy", ref: at("zero") },
    { type: "not", ref: at("settled") },
    { type: "compare", operator: "gt", left: { type: "ref", ref }, right: literal(0) },
    { type: "compare", operator: "eq", left: { type: "ref", ref }, right: literal(240) },
    { type: "compare", operator: "lte", left: { type: "ref", ref }, right: literal(1) },
    {
      type: "and",
      conditions: [{ type: "truthy", ref: at("balanceDue") }, { type: "not", ref: at("zero") }],
    },
    {
      type: "or",
      conditions: [{ type: "truthy", ref: at("zero") }, { type: "truthy", ref: at("blank") }],
    },
    { type: "negate", condition: { type: "truthy", ref: at("name") } },
  ];

  for (const condition of conditions) {
    const taken = await evaluateCondition(condition, state);
    const inverse = await evaluateCondition(invertCondition(condition), state);

    assertEquals(
      [JSON.stringify(condition), inverse],
      [JSON.stringify(condition), !taken],
    );
  }
});

test("inverting twice returns the original meaning", async () => {
  const state = stateWith(data);
  const condition: Condition = {
    type: "compare",
    operator: "gte",
    left: { type: "ref", ref },
    right: literal(240),
  };

  assertEquals(
    await evaluateCondition(invertCondition(invertCondition(condition)), state),
    await evaluateCondition(condition, state),
  );
});

test("truthy inverts into the shape an older engine already understands", () => {
  // Rather than wrapping in `negate`, which an engine predating the richer
  // conditions would not know how to read.
  assertEquals(invertCondition({ type: "truthy", ref }), { type: "not", ref });
  assertEquals(invertCondition({ type: "not", ref }), { type: "truthy", ref });
});
