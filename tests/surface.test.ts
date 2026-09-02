import { test } from "node:test";
import { assertEquals } from "./assert.ts";

/**
 * Every value each entrypoint publishes, written down.
 *
 * An export is a promise: once it ships, somebody can import it, and taking it
 * back is a breaking change. Nothing was checking how many promises were being
 * made, so helpers that were only ever called from inside their own module had
 * been exported anyway, and `docxcelerate` and `docxcelerate/document` had
 * grown apart without anybody deciding they should.
 *
 * This is a snapshot, and it is meant to be edited — but deliberately, in the
 * change that adds or removes the export, where a reviewer can see it. A
 * surface that grows by accident grows here first.
 *
 * Types are not listed. They are checked by `npm run jsr:doc`, which fails on
 * an undocumented export, and by the compiler. What breaks a consumer at
 * runtime is a value that stopped being there.
 */

const surface: Record<string, string[]> = {
  ".": [
    "COMPONENTS",
    "COMPONENT_CATEGORIES",
    "DeriverRegistry",
    "EchoAiClient",
    "InMemoryDataProvider",
    "REGISTRY",
    "REGISTRY_THEMES",
    "THEMES",
    "THEME_IDS",
    "boldBriefTheme",
    "buildDocument",
    "buildProjectEngineDocument",
    "buildProjectFinalDocument",
    "buildProjectPreviewDocument",
    "cleanMinimalDocumentStyle",
    "cleanMinimalTheme",
    "collectDocumentDeriverNames",
    "createDefaultDeriverRegistry",
    "createDeriverBundle",
    "createDeriverRegistry",
    "createDeriverRegistryFromBundle",
    "createDocumentProjectArtifact",
    "ctxRef",
    "dataRef",
    "defineDocumentProject",
    "defineTheme",
    "derive",
    "derivedRef",
    "deriver",
    "evaluateCondition",
    "findRegistryEntry",
    "imageSourceOf",
    "invertCondition",
    "isSvg",
    "legalSerifTheme",
    "listDeriverDefinitionNames",
    "literalValue",
    "normalizeDeriverDefinitions",
    "rasterTypeOf",
    "registryEntry",
    "resolveDocument",
    "runDerivers",
    "slateReportTheme",
    "themeById",
    "themeStyle",
    "themesByCategory",
    "version",
    "warmLetterTheme",
  ],
  "./cli": [
    "runCli",
  ],
  "./document": [
    "Cell",
    "DeriverRegistry",
    "Document",
    "Graph",
    "Image",
    "PageBreak",
    "PageNumber",
    "Paragraph",
    "Row",
    "Section",
    "Table",
    "TableOfContents",
    "__and",
    "__compare",
    "__not",
    "__or",
    "__test",
    "and",
    "branch",
    "buildDocument",
    "cleanMinimalDocumentStyle",
    "compare",
    "createDefaultDeriverRegistry",
    "createDeriverRegistry",
    "createElement",
    "createPublishData",
    "ctxPath",
    "ctxRef",
    "dataPath",
    "dataRef",
    "defineDocumentProject",
    "derive",
    "derivedPath",
    "derivedRef",
    "deriver",
    "elementMarker",
    "expr",
    "host",
    "hostKindOf",
    "hostMarker",
    "isPublishValue",
    "isStaticChildren",
    "isTemplateElement",
    "listDeriverDefinitionNames",
    "literal",
    "literalValue",
    "normalizeDeriverDefinitions",
    "or",
    "publishRefOf",
    "refValue",
    "runDerivers",
    "staticChildrenMarker",
    "template",
    "truthy",
    "useAi",
    "useAvailableTokens",
    "useDeriver",
    "useFormat",
    "usePlaceholderData",
    "useSetPlaceholders",
    "useSetPrompts",
    "useShared",
    "useState",
  ],
  "./docx": [
    "createDocxBlob",
    "createDocxDocument",
    "renderDocxBytes",
  ],
  "./registry": [
    "COMPONENTS",
    "COMPONENT_CATEGORIES",
    "REGISTRY",
    "REGISTRY_THEMES",
    "findRegistryEntry",
    "registryEntry",
  ],
  "./registry/install": [
    "findDocumentProjects",
    "installRegistryEntry",
    "registryRoot",
    "resolveInstallOrder",
  ],
  "./scaffold": [
    "generateNodeDefinition",
    "normalizeDocxcelerateApiEndpoint",
    "officialDocxcelerateApiEndpoint",
    "officialDocxcelerateApiServer",
    "scaffoldDocumentProject",
    "scaffoldWorkspaceProject",
  ],
  "./template": [
    "Cell",
    "Document",
    "Graph",
    "Image",
    "PageBreak",
    "PageNumber",
    "Paragraph",
    "Row",
    "Section",
    "Table",
    "TableOfContents",
    "__and",
    "__compare",
    "__not",
    "__or",
    "__test",
    "and",
    "branch",
    "buildDocument",
    "compare",
    "createElement",
    "createPublishData",
    "ctxPath",
    "dataPath",
    "derivedPath",
    "deriver",
    "elementMarker",
    "expr",
    "host",
    "hostKindOf",
    "hostMarker",
    "isPublishValue",
    "isStaticChildren",
    "isTemplateElement",
    "literal",
    "or",
    "publishRefOf",
    "refValue",
    "staticChildrenMarker",
    "template",
    "truthy",
    "useAi",
    "useAvailableTokens",
    "useDeriver",
    "useFormat",
    "usePlaceholderData",
    "useSetPlaceholders",
    "useSetPrompts",
    "useShared",
    "useState",
  ],
  "./themes": [
    "THEMES",
    "THEME_IDS",
    "boldBriefTheme",
    "cleanMinimalTheme",
    "defineTheme",
    "legalSerifTheme",
    "slateReportTheme",
    "themeById",
    "themeStyle",
    "themesByCategory",
    "warmLetterTheme",
  ],
  "./transform": [
    "assertCompiledSources",
    "compiledMarker",
    "docxcelerateEsbuildTransform",
    "docxcelerateTransform",
    "findUncompiledSources",
    "isCompiledSource",
    "transformDocumentSource",
  ],
};

test("every entrypoint publishes exactly what it says it does", async () => {
  const drift: string[] = [];

  for (const [subpath, expected] of Object.entries(surface)) {
    const specifier = subpath === "." ? "docxcelerate" : `docxcelerate${subpath.slice(1)}`;
    const mod = await import(specifier) as Record<string, unknown>;
    const actual = Object.keys(mod).filter((name) => name !== "default").sort();

    for (const name of actual) {
      if (!expected.includes(name)) drift.push(`${subpath} gained ${name}`);
    }

    for (const name of expected) {
      if (!actual.includes(name)) drift.push(`${subpath} lost ${name}`);
    }
  }

  assertEquals(drift, []);
});

test("the model, the components and defineDocumentProject are the same in / and /document", async () => {
  const root = await import("docxcelerate") as Record<string, unknown>;
  const document = await import("docxcelerate/document") as Record<string, unknown>;
  const shared = ["buildDocument", "defineDocumentProject", "cleanMinimalDocumentStyle", "derive"];
  const missing: string[] = [];

  // The two entrypoints are for different jobs and neither contains the other:
  // /document adds the authoring surface, / adds building and resolving. What
  // they must not do is disagree about the part they both carry.
  for (const name of shared) {
    if (root[name] !== document[name]) missing.push(name);
  }

  assertEquals(missing, []);
});
