import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * The faces Word sets Chinese, Japanese and Korean characters, and complex
 * scripts, in — named in the theme, and left alone by a block's own Latin face.
 *
 * The preview reads `w:eastAsia` into its font stack, so CJK draws in the named
 * face. docx-preview does not read `w:cs`: Arabic and Hebrew draw in the
 * browser's fallback there, a deliberate degradation, so it is not asserted.
 */
export default defineCase({
  id: "text/script-fonts",
  feature: "paragraph.scriptFonts",
  title: "East Asian and complex-script faces from the theme",
  word: "Font → Asian text font / Complex scripts (w:rFonts w:eastAsia, w:cs)",
  claim: "supported",

  style: {
    ...caseStyle,
    typography: { ...caseStyle.typography, eastAsiaFont: "Yu Mincho", complexScriptFont: "Arial" },
    blocks: { ...caseStyle.blocks, money: { font: "Consolas" } },
  },

  document: template(
    <Document id="script-fonts" title="Script fonts">
      <Paragraph id="a">請求書の本文です。</Paragraph>
      <Paragraph id="b" variant="money">合計 ¥12,000</Paragraph>
    </Document>
  ),

  regions: [
    { id: "body", anchor: "請求書の本文" },
    { id: "money", anchor: "合計 ¥12,000" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.docDefaults?.eastAsiaFont, "Yu Mincho", "the document default names the East Asian face");
      is.equal(a.docDefaults?.csFont, "Arial", "and the complex-script face");
      is.equal(a.para("合計 ¥12,000").run.font, "Consolas", "a block's face is the Latin face");
      is.equal(a.para("合計 ¥12,000").run.eastAsiaFont, "Yu Mincho", "and leaves the East Asian face alone");
    },

    preview: (b, is) => {
      is.includes(b.para("請求書の本文").fontFamily, "Yu Mincho", "CJK text falls back to the East Asian face");
    },
  },
});
