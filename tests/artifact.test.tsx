/** @jsxImportSource docxcelerate/template */
import { test } from "node:test";
import { assertEquals, assertRejects } from "./assert.ts";
import {
  createDocumentProjectArtifact,
  ctxRef,
  dataRef,
  defineDocumentProject,
  derive,
} from "docxcelerate";
import { Document, Paragraph, Repeat, Section, template } from "docxcelerate/template";

/**
 * What gets published.
 *
 * The artifact carries the derivers a document names, bundled as source, so an
 * engine can run them without the project. A deriver the artifact fails to
 * notice is not a missing entry in a manifest — it is a document that fails at
 * request time, for every recipient, after the build said everything was fine.
 */
interface Data {
  name: string;
  balanceDue: number;
  visits: Array<{ cost: number }>;
}

const previewData: Data = {
  name: "Avery",
  balanceDue: 240,
  visits: [{ cost: 10 }, { cost: 25 }],
};

const derivers = {
  money: ([amount]: unknown[]) => `£${Number(amount ?? 0).toFixed(2)}`,
  shout: ([text]: unknown[]) => String(text ?? "").toUpperCase(),
};

function projectWith(children: unknown) {
  return defineDocumentProject<Data>({
    id: "artifact",
    name: "Artifact",
    template: template<Data>(
      <Document id="artifact" title="Artifact">{children as never}</Document>,
    ),
    previewData,
    derivers,
  });
}

test("a deriver on a top-level node is collected and bundled", async () => {
  const artifact = await createDocumentProjectArtifact(
    projectWith(
      <Paragraph
        id="balance"
        derivers={[derive("money", { output: "label", inputs: [dataRef("balanceDue")] })]}
      >
        {"{{derived.label}}"}
      </Paragraph>,
    ),
  );

  assertEquals(artifact.manifest.deriverNames, ["money"]);
  assertEquals(artifact.derivers?.names, ["money"]);
  assertEquals(artifact.derivers?.source.includes("money"), true);
});

test("a deriver nested in a section is collected", async () => {
  const artifact = await createDocumentProjectArtifact(
    projectWith(
      <Section id="s" title="S">
        <Paragraph
          id="balance"
          derivers={[derive("money", { output: "label", inputs: [dataRef("balanceDue")] })]}
        >
          {"{{derived.label}}"}
        </Paragraph>
      </Section>,
    ),
  );

  assertEquals(artifact.manifest.deriverNames, ["money"]);
});

test("a deriver inside a repeat is collected too", async () => {
  // The loop is published as a loop, so its body is a place derivers live that
  // no section encloses.
  const artifact = await createDocumentProjectArtifact(
    projectWith(
      <Repeat over="visits" as="visit">
        <Paragraph
          id="visit"
          derivers={[derive("money", { output: "cost", inputs: [ctxRef("visit.cost")] })]}
        >
          {"{{derived.cost}}"}
        </Paragraph>
      </Repeat>,
    ),
  );

  assertEquals(artifact.manifest.deriverNames, ["money"]);
  assertEquals(artifact.derivers?.names, ["money"]);
});

test("a deriver the project never registered fails the build, not the document", async () => {
  await assertRejects(
    () =>
      createDocumentProjectArtifact(
        projectWith(
          <Repeat over="visits" as="visit">
            <Paragraph
              id="visit"
              derivers={[derive("notRegistered", { output: "x", inputs: [] })]}
            >
              {"{{derived.x}}"}
            </Paragraph>
          </Repeat>,
        ),
      ),
    Error,
    "notRegistered",
  );
});

test("only the derivers a document names are bundled", async () => {
  const artifact = await createDocumentProjectArtifact(
    projectWith(
      <Paragraph
        id="balance"
        derivers={[derive("money", { output: "label", inputs: [dataRef("balanceDue")] })]}
      >
        {"{{derived.label}}"}
      </Paragraph>,
    ),
  );

  // `shout` is registered by the project but unused, so it stays out of the
  // published bundle rather than travelling to an engine that will not call it.
  assertEquals(artifact.derivers?.names.includes("shout"), false);
});

test("the artifact carries a preview and an engine document that differ", async () => {
  const artifact = await createDocumentProjectArtifact(
    projectWith(<Paragraph id="greeting">Hello {"{{data.name}}"}.</Paragraph>),
  );

  const preview = JSON.stringify(artifact.previewDocument);
  const engine = JSON.stringify(artifact.engineDocument);

  // The preview is resolved against previewData; the engine document is not.
  assertEquals(preview.includes("Hello Avery."), true);
  assertEquals(engine.includes("{{data.name}}"), true);
  assertEquals(engine.includes("Avery"), false);
});

test("the manifest still writes the names an older engine reads", async () => {
  const artifact = await createDocumentProjectArtifact(
    projectWith(<Paragraph id="greeting">Hello.</Paragraph>),
  );

  assertEquals(artifact.manifest.previewDocument, "preview.json");
  assertEquals(artifact.manifest.previewLetter, artifact.manifest.previewDocument);
  assertEquals(artifact.manifest.engineLetter, artifact.manifest.engineDocument);
  assertEquals(artifact.previewLetter, artifact.previewDocument);
});
