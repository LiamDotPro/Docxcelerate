import { test } from "node:test";
import { assertEquals, assertRejects, assertStringIncludes } from "./assert.ts";
import {
  buildDocument,
  createDocumentProjectArtifact,
  defineDocumentProject,
  type DocumentNode,
  InMemoryDataProvider,
  resolveDocument,
} from "docxcelerate";
import {
  defineMessages,
  Document,
  matchLocale,
  messagesFor,
  Paragraph,
  template,
  useAi,
  useFormat,
  useLocale,
  useMessages,
  useState,
} from "docxcelerate/template";
import { documentXml, partXml } from "./docx.ts";
import { loadRegistryModule } from "./registry_support.ts";

const { PaymentSummary } = await loadRegistryModule<
  typeof import("../registry/components/payment-summary/payment-summary.node.tsx")
>("components/payment-summary/payment-summary.node.tsx");

const messages = defineMessages("en", {
  greeting: (name: string) => `Dear ${name},`,
}, {
  fr: { greeting: (name) => `Chère ${name},` },
  "fr-CA": { greeting: (name) => `Bonjour ${name},` },
});

function texts(nodes: DocumentNode[]): string[] {
  return nodes.flatMap((node) => {
    if (node.kind === "section" || node.kind === "repeat") return texts(node.children);
    return node.kind === "paragraph" ? [node.text ?? ""] : [];
  });
}

const Greeting: Paragraph = () => {
  const t = useMessages(messages);
  return <Paragraph id="greeting">{t.greeting("Avery")}</Paragraph>;
};

const greeting = template(<Document id="d" title="D"><Greeting /></Document>);

test("a locale takes its own messages, then its language's, then the source's", () => {
  assertEquals(matchLocale(["en", "fr", "fr-CA"], "fr-CA"), "fr-CA");
  assertEquals(matchLocale(["en", "fr", "fr-CA"], "FR-be"), "fr");
  assertEquals(matchLocale(["en", "fr-FR"], "fr-CH"), "fr-FR");
  assertEquals(matchLocale(["en", "fr"], "de-DE"), undefined);
  assertEquals(messagesFor(messages, "de-DE").greeting("A"), "Dear A,");
  assertEquals(messagesFor(messages, undefined).greeting("A"), "Dear A,");
});

test("useMessages picks by the build's locale, and the source when it has none", async () => {
  assertEquals(texts((await buildDocument(greeting, {})).nodes), ["Dear Avery,"]);
  assertEquals(texts((await buildDocument(greeting, {}, { locale: "fr-FR" })).nodes), ["Chère Avery,"]);
  assertEquals(texts((await buildDocument(greeting, {}, { locale: "fr-CA" })).nodes), ["Bonjour Avery,"]);
});

test("a message function interpolates a published value into a token", async () => {
  const project = defineDocumentProject({
    id: "p",
    name: "p",
    template: template<{ name: string }>(
      <Document id="d" title="D"><Named /></Document>,
    ),
    previewData: { name: "Avery" },
    locales: ["en-GB", "fr-FR"],
  });

  const artifact = await createDocumentProjectArtifact(project, { builtAt: "x" });
  assertEquals(texts(artifact.engineDocument.nodes), ["Dear {{data.name}},"]);
  assertEquals(texts(artifact.localizedDocuments!["fr-FR"].engineDocument.nodes), ["Chère {{data.name}},"]);
  assertEquals(texts(artifact.localizedDocuments!["fr-FR"].previewDocument.nodes), ["Chère Avery,"]);
  assertEquals(artifact.engineDocument.locale, "en-GB");
  assertEquals(artifact.manifest.locales, ["en-GB", "fr-FR"]);
  assertEquals(artifact.manifest.localizedDocuments, {
    "fr-FR": { previewDocument: "preview.fr-FR.json", engineDocument: "document.fr-FR.json" },
  });
});

function Named() {
  const t = useMessages(messages);
  const [name] = useState((data: { name: string }) => data.name);
  return <Paragraph id="named">{t.greeting(name)}</Paragraph>;
}

test("a project's first locale is its default, and nothing is recorded without one", async () => {
  const plain = await buildDocument(greeting, {});
  assertEquals("locale" in plain, false);

  const LocaleProbe: Paragraph = () => <Paragraph id="l">{useLocale()}</Paragraph>;
  const project = defineDocumentProject({
    id: "p",
    name: "p",
    template: template(<Document id="d" title="D"><LocaleProbe /></Document>),
    previewData: {},
    locales: ["de-DE", "en-GB"],
  });
  const artifact = await createDocumentProjectArtifact(project, { builtAt: "x" });
  assertEquals(texts(artifact.previewDocument.nodes), ["de-DE"]);
  assertEquals(artifact.previewDocument.locale, "de-DE");
});

test("a malformed locale is refused where it enters", async () => {
  await assertRejects(() => buildDocument(greeting, {}, { locale: "not a locale" }), Error, "BCP 47");
});

test("plural follows the locale's rules", async () => {
  const seen: string[] = [];
  const Probe: Paragraph = () => {
    const en = useFormat("en");
    const fr = useFormat("fr");
    const pl = useFormat("pl");
    const forms = { one: "plik", few: "pliki", many: "plików", other: "pliku" };
    seen.push(
      en.plural(0, "file"),
      en.plural(1, "file"),
      fr.plural(0, "fichier", "fichiers"),
      fr.plural(2, "fichier", "fichiers"),
      pl.plural(1, forms),
      pl.plural(3, forms),
      pl.plural(5, forms),
    );
    return <Paragraph id="p">x</Paragraph>;
  };
  await buildDocument(template(<Document id="d" title="D"><Probe /></Document>), {});
  assertEquals(seen, ["files", "file", "fichier", "fichiers", "plik", "pliki", "plików"]);
});

test("Word is told the document's language, and only when it has one", async () => {
  const french = await buildDocument(greeting, {}, { locale: "fr-FR" });
  assertStringIncludes(await partXml(french, "word/styles.xml"), '<w:lang w:val="fr-FR"/>');

  const plain = await buildDocument(greeting, {});
  assertEquals((await partXml(plain, "word/styles.xml")).includes("w:lang"), false);
  // The body is untouched by it.
  assertStringIncludes(await documentXml(french), "Chère Avery,");
});

test("a dynamic node is asked to write in the document's language, at build and in the engine", async () => {
  const asked: string[] = [];
  const aiClient = {
    generateParagraph({ prompt }: { prompt: string }) {
      asked.push(prompt);
      return "Bonjour.";
    },
  };
  const Summary: Paragraph = () => {
    useAi({ ask: "Summarise.", placeholder: "A summary." });
    return <Paragraph id="summary" />;
  };
  const summary = template(<Document id="d" title="D"><Summary /></Document>);

  await buildDocument(summary, {}, { locale: "fr-FR", aiClient });
  await buildDocument(summary, {}, { aiClient });

  const published = await buildDocument(summary, {}, { locale: "fr-FR", dynamicMode: "placeholder" });
  await resolveDocument(
    { ...published, nodes: published.nodes.map((node) => ({ ...node, text: undefined })) },
    { ctx: {}, derived: {}, dataProvider: new InMemoryDataProvider({}), aiClient },
  );

  assertStringIncludes(asked[0], "SYSTEM: Write in French (France) (fr-FR), the language of this document.");
  assertEquals(asked[1].includes("Write in"), false);
  assertStringIncludes(asked[2], "French (France)");
});

test("the registry's payment summary is translated, figures and all", async () => {
  const doc = template(<Document id="d" title="D"><PaymentSummary /></Document>);
  const data = { account: { reference: "RIV-1", balanceDue: 1234.5, dueBy: "2026-09-30", currency: "EUR" } };

  const fr = await buildDocument(doc, data, { locale: "fr-FR" });
  assertEquals(fr.nodes[0].kind === "section" && fr.nodes[0].title, "Votre solde");
  // Intl spaces French and German figures with no-break spaces; compared as plain ones.
  const plain = (text: string) => text.replace(/\s/g, " ");
  assertEquals(plain(texts(fr.nodes)[1]), "Le solde de votre compte s’élève à 1 234,50 €, à régler au plus tard le 30 septembre 2026.");

  const de = await buildDocument(doc, data, { locale: "de-DE" });
  assertEquals(plain(texts(de.nodes)[1]), "Der offene Betrag auf Ihrem Konto beträgt 1.234,50 €, zahlbar bis zum 30. September 2026.");

  const en = await buildDocument(doc, data, { locale: "en-GB" });
  assertEquals(texts(en.nodes)[1], "The balance on your account is €1,234.50, payable by 30 September 2026.");
});
