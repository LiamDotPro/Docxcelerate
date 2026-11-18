import assert from "node:assert/strict";
import { test } from "node:test";
import { arrange } from "docxcelerate/template";
import { assertEquals } from "./assert.ts";

const built: string[] = [];
const arrangements = {
  stacked: () => (built.push("stacked"), "stacked"),
  table: () => (built.push("table"), "table"),
};

test("arrange builds only the arrangement asked for", () => {
  built.length = 0;
  assertEquals(arrange("table", arrangements), "table");
  assertEquals(built, ["table"]);
});

test("no arrangement asked for is the first one written", () => {
  assertEquals(arrange(undefined, arrangements), "stacked");
});

test("an arrangement nobody wrote fails, naming the ones there are", () => {
  assert.throws(() => arrange("grid" as never, arrangements), /expected one of stacked, table/);
  // Not reachable through the prototype either.
  assert.throws(() => arrange("toString" as never, arrangements), /"toString"/);
  assert.throws(() => arrange(undefined, {}), /Unknown arrangement/);
});
