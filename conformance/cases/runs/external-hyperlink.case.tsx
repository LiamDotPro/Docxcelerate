import { Document, Paragraph, template, Text } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Words a reader can follow to an address outside the document.
 *
 * `<Text href>` is what says it. Before there was one, a document that wanted
 * to send its reader somewhere printed the address into the sentence and left
 * the reader to copy it out — which is what this case did when it was written
 * red. Word will sometimes turn a printed address into a link by itself as it
 * is typed; it does no such thing to a file it is handed.
 *
 * A link is two things in two parts of the package, and a packer can get
 * either half wrong alone. The element in `document.xml` names a relationship;
 * the relationship holds the address, and has to be marked `External`. Leave
 * that mark off and Word reads the address as a part of the package, finds no
 * such part, and refuses the whole file — so it is asserted on its own line
 * rather than taken on trust from the address being right.
 *
 * **The Word tier is not written, deliberately.** The Word probe has no
 * reading of a hyperlink, and the one this case wants — `Hyperlinks(1).Address`
 * — would be PowerShell written on a machine with no Word to run it against.
 * A probe nobody has run is a measurement nobody took. It goes in on the
 * Windows machine — and until it has, "supported" here means the file and the
 * preview, which is everything this case can see.
 */
const ADDRESS = "https://example.com/pay";

export default defineCase({
  id: "runs/external-hyperlink",
  feature: "run.hyperlink",
  title: "Words that link to an address outside the document",
  word: "Insert → Link → Existing File or Web Page (w:hyperlink, TargetMode=External)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="external-hyperlink" title="External hyperlink">
      <Paragraph id="pay">
        You can <Text href={ADDRESS}>pay online</Text> before the due date.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "sentence", anchor: "You can pay online" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.includes(a.para("You can pay online").text, "before the due date.", "the sentence is packed, whole");

      is.equal(a.link("pay online").text, "pay online", "the two words are the link");
      is.equal(a.link("pay online").target, ADDRESS, "which points at the address");
      is.equal(a.link("pay online").external, true, "through a relationship marked external");
      is.excludes(a.para("You can pay online").text, "https://", "so the address is followed, not printed");
    },

    preview: (b, is) => {
      is.equal(b.para("You can pay online").lineCount, 1, "the sentence is set on its one line");

      is.equal(b.run("pay online").text, "pay online", "the preview draws the two words as a run of their own");
      is.equal(b.run("pay online").href, ADDRESS, "inside a link to the same address");
    },
  },
});
