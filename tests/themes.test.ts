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
import { documentXml, partXml } from "./docx.ts";

/**
 * Themes as data.
 *
 * A theme is only worth having if a document carries it to whatever renders
 * one — otherwise it is a decision made in a project and lost on the way out.
 * So these tests care about two things: the style a theme resolves to, and
 * that the packed file can be seen honouring it.
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

test("the packed file is set from the document style", async () => {
  const doc = themedDocument(legalSerifTheme.id);
  const styles = await partXml(doc, "word/styles.xml");
  const body = await documentXml(doc);

  // Legal Serif is 12pt Times on a wider left margin. All three reach the file
  // rather than the packer's own defaults — half-points for the size, twips
  // for the margin.
  assertStringIncludes(styles, 'w:ascii="Times New Roman"');
  assertStringIncludes(styles, '<w:sz w:val="24"/>');
  assertStringIncludes(body, 'w:left="1800"');
});

test("US Letter and landscape change the page that is packed", async () => {
  const letter = await documentXml(themedDocument(boldBriefTheme.id));

  // 215.9mm by 279.4mm, in twips.
  assertStringIncludes(letter, 'w:w="12240"');
  assertStringIncludes(letter, 'w:h="15840"');

  const turned: DocumentModel = {
    ...themedDocument(boldBriefTheme.id),
    style: themeStyle(boldBriefTheme, { page: { orientation: "landscape" } }),
  };
  const sideways = await documentXml(turned);

  assertStringIncludes(sideways, 'w:w="15840"');
  assertStringIncludes(sideways, 'w:h="12240"');
  assertStringIncludes(sideways, 'w:orient="landscape"');
});

test("a document with no style still packs, on the fallback", async () => {
  const body = await documentXml(themedDocument(undefined));

  // A4, and the default theme's 11pt body.
  assertStringIncludes(body, 'w:w="11906"');
  assertStringIncludes(await partXml(themedDocument(undefined), "word/styles.xml"), '<w:sz w:val="22"/>');
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
