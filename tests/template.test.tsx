/** @jsxImportSource docxcelerate/template */
import { test } from "node:test";
import { assertEquals } from "./assert.ts";
import { buildLetterDocument, staticParagraph } from "docxcelerate";
import {
  defineLetterComponent,
  defineSectionComponent,
  Letter,
  Section,
  template,
} from "docxcelerate/template";

interface LetterData {
  recipientName: string;
}

const Greeting = staticParagraph<LetterData>({
  id: "greeting",
  render(data) {
    return `Hello ${data.recipientName},`;
  },
});

test("TSX template composes a letter tree from node components", async () => {
  const letterTemplate = template<LetterData>(
    <Letter id="welcome" title="Welcome">
      <Section id="opening" title="Opening">
        <Greeting />
      </Section>
    </Letter>,
  );

  const builtLetter = await buildLetterDocument(letterTemplate, {
    recipientName: "Avery",
  });

  assertEquals(builtLetter.nodes.length, 1);
  const opening = builtLetter.nodes[0];
  assertEquals(opening.kind, "section");

  if (opening.kind !== "section") {
    throw new Error("Expected section");
  }

  const greeting = opening.children[0];
  assertEquals(greeting.kind, "paragraph");

  if (greeting.kind === "paragraph") {
    assertEquals(greeting.text, "Hello Avery,");
  }
});

test("TSX template can extend generic letter and section components", async () => {
  const CaseLetter = defineLetterComponent<LetterData>({
    id: "case-letter",
    title: "Case Letter",
  });
  const OpeningSection = defineSectionComponent<
    LetterData,
    { title?: string }
  >((props) => ({
    id: "opening",
    title: props.title ?? "Opening",
  }));

  const letterTemplate = template<LetterData>(
    <CaseLetter>
      <OpeningSection title="Opening Details">
        <Greeting />
      </OpeningSection>
    </CaseLetter>,
  );

  const builtLetter = await buildLetterDocument(letterTemplate, {
    recipientName: "Avery",
  });
  const opening = builtLetter.nodes[0];

  assertEquals(builtLetter.id, "case-letter");
  assertEquals(builtLetter.title, "Case Letter");
  assertEquals(opening.kind, "section");

  if (opening.kind === "section") {
    assertEquals(opening.title, "Opening Details");
  }
});
