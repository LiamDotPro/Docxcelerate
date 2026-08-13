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

test("scaffold creates a structured letter project and node generator updates exports", async () => {
  const lettersDir = await tempDir();
  const scaffold = await scaffoldDocumentProject({
    name: "case review",
    title: "Case Review",
    lettersDir,
  });
  const generatedNode = await generateNodeDefinition({
    projectDir: scaffold.projectDir,
    name: "risk summary",
    type: "paragraph",
    mode: "dynamic",
  });

  const index = await readTextFile(`${scaffold.projectDir}/nodes/index.ts`);
  const node = await readTextFile(generatedNode.filePath);
  const style = await readTextFile(`${scaffold.projectDir}/letter-style.ts`);
  const derivers = await readTextFile(`${scaffold.projectDir}/derivers/index.ts`);
  const projectSource = await readTextFile(scaffold.entrypoint);

  assertEquals(generatedNode.componentName, "RiskSummary");
  assertEquals(index.includes("RiskSummary"), true);
  assertEquals(node.includes("paragraph<LetterData>"), true);
  // Mode is inferred from the options, so the prompt is what makes this node
  // dynamic — assert on that rather than on a helper name.
  assertEquals(node.includes("generalPrompt"), true);
  assertEquals(style.includes("cleanMinimalDocumentStyle"), true);
  assertEquals(derivers.includes("DeriverDefinitions"), true);
  assertEquals(projectSource.includes("derivers"), true);
});

test("workspace scaffold creates a project container for letters", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "housing letters",
    parentDir,
  });
  const packageJson = JSON.parse(
    await readTextFile(`${workspace.projectDir}/package.json`),
  );
  const config = JSON.parse(
    await readTextFile(`${workspace.projectDir}/docxcelerate.config.json`),
  );
  const tsconfig = await readTextFile(`${workspace.projectDir}/tsconfig.json`);
  const gitignore = await readTextFile(`${workspace.projectDir}/.gitignore`);
  const sampleDeriver = await readTextFile(
    `${workspace.projectDir}/letters/welcome/derivers/index.ts`,
  );
  const sampleNode = await readTextFile(
    `${workspace.projectDir}/letters/welcome/nodes/balance-summary.node.ts`,
  );
  const previewMain = await readTextFile(`${workspace.projectDir}/preview/main.ts`);
  const viteConfig = await readTextFile(`${workspace.projectDir}/vite.config.ts`);

  assertEquals(workspace.projectDir.endsWith("/housing-letters"), true);
  assertEquals(workspace.template, "sample");
  assertEquals(packageJson.scripts.dev, "vite --host 127.0.0.1 --port 4507");
  assertEquals(packageJson.scripts["letter:new"], "dxcl letter new");
  assertEquals(packageJson.dependencies.docxcelerate, "^0.1.3");
  assertEquals(packageJson.dependencies.docx, "^9.6.1");
  assertEquals(packageJson.dependencies["docx-preview"], "^0.3.7");
  assertEquals(packageJson.devDependencies.vite, "^8.0.13");
  assertEquals(config.schemaVersion, "docxcelerate.config/v0");
  assertEquals(config.activePreset, "local");
  assertEquals(config.presets.local.upload.endpoint, "");
  assertEquals(config.presets.local.upload.body, "stored-letter");
  assertEquals(tsconfig.includes('"types": ["vite/client"]'), true);
  assertEquals(tsconfig.includes("preview/**/*.ts"), true);
  assertEquals(previewMain.includes("createDocumentProjectArtifact"), true);
  assertEquals(previewMain.includes("createDocxBlob"), true);
  assertEquals(previewMain.includes('import("docxcelerate/docx")'), true);
  assertEquals(previewMain.includes("docxPreview.renderAsync"), true);
  assertEquals(previewMain.includes('"microsoft-office"'), true);
  assertEquals(previewMain.includes('"google-docs"'), true);
  assertEquals(previewMain.includes("view.officeapps.live.com/op/embed.aspx"), true);
  assertEquals(previewMain.includes("docs.google.com/gview"), true);
  assertEquals(previewMain.includes("isPrivatePreviewHost"), true);
  assertEquals(previewMain.includes("renderDocumentPreview"), true);
  assertEquals(previewMain.includes("renderHome"), true);
  assertEquals(previewMain.includes("Build & upload"), true);
  assertEquals(viteConfig.includes("/api/docxcelerate/letters"), true);
  assertEquals(viteConfig.includes("/api/docxcelerate/build"), true);
  assertEquals(viteConfig.includes("/api/docxcelerate/preview-docx"), true);
  assertEquals(viteConfig.includes("PreviewDocxFile"), true);
  assertEquals(viteConfig.includes("safeDocxFileName"), true);
  assertEquals(viteConfig.includes('import { dirname, join } from "node:path";'), true);
  assertEquals(viteConfig.includes("entrypoint?: string;"), true);
  assertEquals(viteConfig.includes("LetterDeriverBundlePayload"), true);
  assertEquals(viteConfig.includes("artifact.derivers"), true);
  assertEquals(viteConfig.includes('const buildDir = preset.build?.outDir ?? "build";'), true);
  assertEquals(
    viteConfig.includes("const outDir = join(letterDirFromArtifact(artifact), buildDir);"),
    true,
  );
  assertEquals(viteConfig.includes("function letterDirFromArtifact"), true);
  assertEquals(viteConfig.includes('return join("letters", slugify(artifact.manifest.id));'), true);
  assertEquals(viteConfig.includes("return artifact.engineLetter;"), true);
  assertEquals(viteConfig.includes("safeFileName"), false);
  assertEquals(await exists(`${workspace.projectDir}/index.html`), true);
  assertEquals(await exists(`${workspace.projectDir}/docxcelerate.config.json`), true);
  assertEquals(await exists(`${workspace.projectDir}/vite.config.ts`), true);
  assertEquals(await exists(`${workspace.projectDir}/preview/main.ts`), true);
  assertEquals(await exists(`${workspace.projectDir}/preview/styles.css`), true);
  assertEquals(await exists(`${workspace.projectDir}/letters/welcome/letter.project.ts`), true);
  assertEquals(sampleDeriver.includes("currencyLabel"), true);
  assertEquals(sampleNode.includes('derive("currencyLabel"'), true);
  assertEquals(tsconfig.includes("letters/**/*.tsx"), true);
  assertEquals(gitignore.includes("letters/**/build/"), true);
  assertEquals(await exists(`${workspace.projectDir}/letters/.gitkeep`), false);
});

test("workspace scaffold can create a blank project", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "blank letters",
    parentDir,
    template: "blank",
  });

  assertEquals(workspace.template, "blank");
  assertEquals(await exists(`${workspace.projectDir}/letters/.gitkeep`), true);
  assertEquals(await exists(`${workspace.projectDir}/letters/welcome/letter.project.ts`), false);
});

test("workspace scaffold can configure the official Docxcelerate API endpoint", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "official letters",
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

test("letter scaffold can create letters inside a workspace", async () => {
  const parentDir = await tempDir();
  const workspace = await scaffoldWorkspaceProject({
    name: "case workspace",
    parentDir,
  });
  const letter = await scaffoldDocumentProject({
    name: "case review",
    title: "Case Review",
    lettersDir: `${workspace.projectDir}/letters`,
  });

  assertEquals(letter.projectDir.endsWith("/case-workspace/letters/case-review"), true);
  assertEquals(await exists(`${workspace.projectDir}/letters/case-review/letter.project.ts`), true);
  assertEquals(await exists(`${workspace.projectDir}/letters/case-review/letter-style.ts`), true);
});

test("node generator supports image and graph node types", async () => {
  const lettersDir = await tempDir();
  const scaffold = await scaffoldDocumentProject({
    name: "visual review",
    lettersDir,
  });
  const image = await generateNodeDefinition({
    projectDir: scaffold.projectDir,
    name: "signature image",
    type: "image",
    mode: "static",
  });
  const graph = await generateNodeDefinition({
    projectDir: scaffold.projectDir,
    name: "trend chart",
    type: "graph",
    mode: "dynamic",
  });

  const imageSource = await readTextFile(image.filePath);
  const graphSource = await readTextFile(graph.filePath);
  const index = await readTextFile(`${scaffold.projectDir}/nodes/index.ts`);

  assertEquals(imageSource.includes("image<LetterData>"), true);
  assertEquals(imageSource.includes("src"), true);
  assertEquals(graphSource.includes("graph<LetterData>"), true);
  assertEquals(graphSource.includes("generalPrompt"), true);
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
