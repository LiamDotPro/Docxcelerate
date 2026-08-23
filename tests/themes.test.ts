import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { cleanMinimalDocumentStyle, type DocumentModel } from "docxcelerate";
import {
  boldBriefTheme,
  legalSerifTheme,
  slateReportTheme,
  THEME_IDS,
  THEMES,
  themeById,
  themeStyle,
} from "docxcelerate/themes";
import { renderDocumentWebsite } from "docxcelerate/renderer";

/**
 * Themes as data.
 *
 * A theme is only worth having if a document carries it to whatever renders
 * one — otherwise it is a decision made in a project and lost on the way out.
 * So these tests care about two things: the style a theme resolves to, and that
 * a renderer can be seen honouring it.
 */

test("a theme names itself in the style it resolves to", () => {
  for (const theme of THEMES) {
    assertEquals(theme.style.preset, theme.id);
  }
});

test("theme ids are unique, and the default is still the fallback style", () => {
  assertEquals(new Set(THEME_IDS).size, THEME_IDS.length);
  // The renderer's fallback and the Clean Minimal theme are one object, so a
  // change to the theme cannot leave the fallback behind.
  assertEquals(cleanMinimalDocumentStyle, themeById("clean-minimal").style);
});

test("a theme lists every font it asks for", () => {
  assertEquals(slateReportTheme.fonts, ["Calibri"]);
  assertEquals(themeById("clean-minimal").fonts, ["Aptos", "Cambria"]);
});

test("themeById refuses an id nothing carries", () => {
  try {
    themeById("no-such-theme");
    throw new Error("Expected themeById to throw.");
  } catch (error) {
    assertStringIncludes(String(error), `Unknown theme "no-such-theme"`);
  }
});

test("overriding one margin keeps the other three", () => {
  const style = themeStyle(slateReportTheme, { page: { margins: { topMm: 40 } } });

  assertEquals(style.page.margins.topMm, 40);
  assertEquals(style.page.margins.leftMm, slateReportTheme.style.page.margins.leftMm);
  assertEquals(style.page.margins.bottomMm, slateReportTheme.style.page.margins.bottomMm);
  // The style still says which theme it came from, however much was overridden.
  assertEquals(style.preset, "slate-report");
  // And overriding is not mutation: the theme is shared, so a project that
  // edits its own style must not edit everybody else's.
  assertEquals(slateReportTheme.style.page.margins.topMm, 20);
});

test("overriding one group leaves the rest of the style alone", () => {
  const style = themeStyle(boldBriefTheme, { typography: { bodySizePt: 9 } });

  assertEquals(style.typography.bodySizePt, 9);
  assertEquals(style.typography.bodyFont, "Verdana");
  assertEquals(style.title.fontSizePt, boldBriefTheme.style.title.fontSizePt);
});

test("the web renderer sets the page from the document style", () => {
  const page = renderDocumentWebsite(themedDocument(legalSerifTheme.id));

  // Legal Serif is 12pt Times on a wider left margin — all three come through
  // as custom properties rather than as the renderer's own defaults.
  assertStringIncludes(page, "--body-size: 12pt;");
  assertStringIncludes(page, `--body-font: "Times New Roman", serif;`);
  assertStringIncludes(page, "--page-padding: 25.4mm 25.4mm 25.4mm 31.75mm;");
  assertStringIncludes(page, "--heading-transform: uppercase;");
});

test("US Letter and landscape change the page the preview draws", () => {
  const letter = renderDocumentWebsite(themedDocument(boldBriefTheme.id));
  assertStringIncludes(letter, "--page-width: 215.9mm;");
  assertStringIncludes(letter, "--page-height: 279.4mm;");

  const turned: DocumentModel = {
    ...themedDocument(boldBriefTheme.id),
    style: themeStyle(boldBriefTheme, { page: { orientation: "landscape" } }),
  };

  assertStringIncludes(renderDocumentWebsite(turned), "--page-width: 279.4mm;");
  assertStringIncludes(renderDocumentWebsite(turned), "--page-height: 215.9mm;");
});

test("a document with no style still renders, on the fallback", () => {
  const page = renderDocumentWebsite(themedDocument(undefined));

  assertStringIncludes(page, "--page-width: 210mm;");
  assertStringIncludes(page, "--body-size: 11pt;");
});

/** One paragraph, set in whichever theme was asked for. */
function themedDocument(themeId: string | undefined): DocumentModel {
  return {
    schemaVersion: "docxcelerate.letter/v0",
    id: "themed",
    title: "Themed",
    style: themeId ? themeById(themeId).style : undefined,
    nodes: [{ id: "body", kind: "paragraph", mode: "static", text: "Set in a theme." }],
  };
}
