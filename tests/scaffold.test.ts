import { test } from "node:test";
import { mkdtemp, readFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { assertEquals } from "./assert.ts";
import {
  generateNodeDefinition,
  officialDocxcelerateApiEndpoint,
  officialDocxcelerateApiServer,
  scaffoldDocumentProject,
  scaffoldWorkspaceProject,
} from "docxcelerate/scaffold";
import { version } from "docxcelerate";

test("scaffold creates a structured document project and node generator updates exports", async () => {
  const documentsDir = await tempDir();
  const scaffold = await scaffoldDocumentProject({
    name: "case review",
    title: "Case Review",
    documentsDir,
  });
  const generatedNode = await generateNodeDefinition({
    projectDir: scaffold.projectDir,
    name: "risk summary",
    type: "paragraph",
  });

  const index = await readTextFile(`${scaffold.projectDir}/nodes/index.ts`);
  const node = await readTextFile(generatedNode.filePath);
  const style = await readTextFile(`${scaffold.projectDir}/document-style.ts`);
  const derivers = await readTextFile(`${scaffold.projectDir}/derivers/index.ts`);
  const projectSource = await readTextFile(scaffold.entrypoint);

  assertEquals(generatedNode.componentName, "RiskSummary");
  assertEquals(index.includes("RiskSummary"), true);
  assertEquals(node.includes("const RiskSummary: Paragraph = () =>"), true);
  // Data reaches a component through state, never through render.
  assertEquals(node.includes("useState((data: DocumentData)"), true);
  // A generated node starts with its content. Prompts are what would make it
  // dynamic, and the generator has no business deciding that for you.
  assertEquals(node.includes(`<Paragraph id="risk-summary">`), true);
  assertEquals(node.includes("useSetPrompts"), false);
  assertEquals(style.includes("cleanMinimalDocumentStyle"), true);
  assertEquals(derivers.includes("DeriverDefinitions"), true);
  assertEquals(projectSource.includes("derivers"), true);
});

test("workspace scaffold creates a project container for documents", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "housing documents",
    parentDir,
  });

  assertEquals(workspace.projectDir.endsWith("/housing-documents"), true);
  assertEquals(workspace.template, "sample");
  assertEquals(
    workspace.files.map((path) => path.slice(workspace.projectDir.length + 1)).sort(),
    [
      ".gitignore",
      "README.md",
      "documents/welcome/derivers/index.ts",
      "documents/welcome/document-style.ts",
      "documents/welcome/document.project.ts",
      "documents/welcome/document.tsx",
      "documents/welcome/nodes/balance-summary.node.tsx",
      "documents/welcome/nodes/greeting.node.tsx",
      "documents/welcome/nodes/index.ts",
      "documents/welcome/preview-data.ts",
      "documents/welcome/types.ts",
      "docxcelerate.config.json",
      "index.html",
      "package.json",
      "preview/main.ts",
      "preview/styles.css",
      "tsconfig.json",
      "vite.config.ts",
    ],
  );
});

test("a workspace is filled in with its own name, version and endpoint", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "housing documents",
    parentDir,
  });
  const packageJson = JSON.parse(await readTextFile(`${workspace.projectDir}/package.json`));
  const config = JSON.parse(
    await readTextFile(`${workspace.projectDir}/docxcelerate.config.json`),
  );
  const indexHtml = await readTextFile(`${workspace.projectDir}/index.html`);
  const readme = await readTextFile(`${workspace.projectDir}/README.md`);

  assertEquals(packageJson.name, "housing-documents");
  // A scaffolded workspace depends on the toolkit that scaffolded it, not on
  // whatever version happened to be current when the template was written.
  assertEquals(packageJson.dependencies.docxcelerate, `^${version}`);
  assertEquals(config.schemaVersion, "docxcelerate.config/v0");
  assertEquals(config.presets.local.upload.endpoint, "");
  assertEquals(indexHtml.includes("<title>Housing Documents Preview</title>"), true);
  assertEquals(readme.startsWith("# Housing Documents\n"), true);
  assertEquals(readme.includes("Docxcelerate API endpoint: not configured"), true);
  assertEquals(
    readme.includes("A sample document is available at `documents/welcome/document.project.ts`."),
    true,
  );
});

test("nothing a scaffold writes is left carrying a placeholder", async () => {
  const parentDir = await tempDir();
  const sample = await scaffoldWorkspaceProject({ name: "sample workspace", parentDir });
  const blank = await scaffoldWorkspaceProject({
    name: "blank workspace",
    parentDir,
    template: "blank",
  });
  const document = await scaffoldDocumentProject({
    name: "tenancy renewal",
    documentsDir: `${parentDir}/documents`,
  });
  const nodes = await Promise.all(
    (["paragraph", "image", "graph"] as const).map((type) =>
      generateNodeDefinition({ projectDir: document.projectDir, name: `${type} block`, type })
    ),
  );

  // A template that grows a placeholder nobody fills is the failure this
  // catches, across every file every entrypoint writes rather than one at a
  // time. `readTemplate` throws on an unfilled one, so reaching a written file
  // that still carries the syntax means something wrote it past the templates.
  const written = [
    ...sample.files,
    ...blank.files,
    ...document.files,
    ...nodes.map((node) => node.filePath),
  ];
  const carrying: string[] = [];

  for (const path of written) {
    if (/__[A-Z0-9_]+__/.test(await readTextFile(path))) {
      carrying.push(path);
    }
  }

  assertEquals(carrying, []);
});

test("a blank workspace says so where a sample one points at its example", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "blank documents",
    parentDir,
    template: "blank",
  });
  const readme = await readTextFile(`${workspace.projectDir}/README.md`);

  assertEquals(workspace.template, "blank");
  assertEquals(await exists(`${workspace.projectDir}/documents/.gitkeep`), true);
  assertEquals(await exists(`${workspace.projectDir}/documents/welcome/document.project.ts`), false);
  assertEquals(
    readme.includes("This workspace starts blank. Create a document with `dxcl document new`."),
    true,
  );
  assertEquals(readme.includes("A sample document is available"), false);
});

test("workspace scaffold can configure the official Docxcelerate API endpoint", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "official documents",
    parentDir,
    apiEndpoint: officialDocxcelerateApiServer,
  });
  const config = JSON.parse(
    await readTextFile(`${workspace.projectDir}/docxcelerate.config.json`),
  );
  const readme = await readTextFile(`${workspace.projectDir}/README.md`);

  assertEquals(workspace.apiEndpoint, officialDocxcelerateApiEndpoint);
  assertEquals(config.presets.local.upload.endpoint, officialDocxcelerateApiEndpoint);
  assertEquals(
    readme.includes(`Docxcelerate API endpoint: ${officialDocxcelerateApiEndpoint}`),
    true,
  );
});

test("document scaffold can create documents inside a workspace", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "case workspace",
    parentDir,
  });
  const document = await scaffoldDocumentProject({
    name: "case review",
    title: "Case Review",
    documentsDir: `${workspace.projectDir}/documents`,
  });

  assertEquals(document.projectDir.endsWith("/case-workspace/documents/case-review"), true);
  assertEquals(await exists(`${workspace.projectDir}/documents/case-review/document.project.ts`), true);
  assertEquals(await exists(`${workspace.projectDir}/documents/case-review/document-style.ts`), true);
});

test("node generator supports image and graph node types", async () => {
  const documentsDir = await tempDir();
  const scaffold = await scaffoldDocumentProject({
    name: "visual review",
    documentsDir,
  });
  const image = await generateNodeDefinition({
    projectDir: scaffold.projectDir,
    name: "signature image",
    type: "image",
  });
  const graph = await generateNodeDefinition({
    projectDir: scaffold.projectDir,
    name: "trend chart",
    type: "graph",
  });

  const imageSource = await readTextFile(image.filePath);
  const graphSource = await readTextFile(graph.filePath);
  const index = await readTextFile(`${scaffold.projectDir}/nodes/index.ts`);

  assertEquals(imageSource.includes("const SignatureImage: Image = () =>"), true);
  assertEquals(imageSource.includes("src="), true);
  assertEquals(graphSource.includes("const TrendChart: Graph = () =>"), true);
  assertEquals(graphSource.includes("graphType=\"bar\""), true);
  assertEquals(index.includes("SignatureImage"), true);
  assertEquals(index.includes("TrendChart"), true);
});

/** Scaffolding joins paths with "/", so keep the temp root in that form too. */
async function tempDir(): Promise<string> {
  const created = await mkdtemp(join(tmpdir(), "dxcl-"));

  return created.replaceAll("\\", "/");
}

function readTextFile(path: string): Promise<string> {
  return readFile(path, "utf8");
}

async function exists(path: string): Promise<boolean> {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}
