import { test } from "node:test";
import { assertEquals } from "./assert.ts";
import { settleDocxPreview } from "docxcelerate/preview";
import { JSDOM } from "jsdom";

/**
 * A double border is drawn across three of the file's widths, as Word draws
 * it, because CSS cannot show two lines in less than 3px. Nothing else is
 * widened.
 */
test("a double border is widened to three widths and a dashed one is left alone", () => {
  const { document } = new JSDOM("<!doctype html><html><body></body></html>").window;
  const body = document.createElement("div");

  body.innerHTML =
    '<p style="border-bottom: 1pt double #1C2340"></p><p style="border-top: 1pt dashed #5A6482"></p>';
  settleDocxPreview(body);

  const [double, dashed] = [...body.querySelectorAll("p")];

  assertEquals(double.style.borderBottomWidth, "3pt");
  assertEquals(double.style.borderBottomStyle, "double");
  assertEquals(dashed.style.borderTopWidth, "1pt");
});
