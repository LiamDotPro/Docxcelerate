import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A numbered list taken up again after other text, counting from 3.
 *
 * Anything that is not an item ends a list, so the clauses after the
 * schedule would start again at one. `listStart` on the item that opens the
 * second list says where it picks up, and the item after it follows.
 *
 * The preview's 3 lives in the `counter-reset` docx-preview writes from
 * `w:start`, which no probe can read: it can only see that the item is
 * counted by a counter of its own.
 */
export default defineCase({
  id: "lists/list-start",
  feature: "list.start",
  title: "A numbered list picked up again at 3",
  word: "Home → Numbering → Set Numbering Value (w:start)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="list-start" title="Numbered clauses">
      <Paragraph id="supply" list="number">The Supplier shall deliver the Goods.</Paragraph>
      <Paragraph id="pay" list="number">The Buyer shall pay the Price.</Paragraph>
      <Paragraph id="schedule">The Schedule follows.</Paragraph>
      <Paragraph id="risk" list="number" listStart={3}>Risk passes on delivery.</Paragraph>
      <Paragraph id="law" list="number">This agreement is governed by English law.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "supply", anchor: "The Supplier shall" },
    { id: "pay", anchor: "The Buyer shall pay" },
    { id: "schedule", anchor: "The Schedule follows" },
    { id: "risk", anchor: "Risk passes" },
    { id: "law", anchor: "governed by English law" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.list("Risk passes").start, 3, "the list counts from the number it was given");
      is.equal(a.list("governed by English law").numId, a.list("Risk passes").numId, "and the next item follows it");
      is.equal(a.list("The Supplier shall").start, 1, "the first list still counts from one");
      is.notEqual(a.list("Risk passes").numId, a.list("The Supplier shall").numId, "as a list of its own");
    },

    preview: (b, is) => {
      is.includes(b.para("Risk passes").marker, "counter(", "the preview counts it with a counter");
      is.notEqual(b.para("Risk passes").marker, b.para("The Supplier shall").marker, "of its own, not the first list's");
    },

    word: (c, is) => {
      is.less(c.para("Risk passes").firstLineIndent, 0, "Word hangs the first line, where its numeral stands");
    },
  },
});
