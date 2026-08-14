import {
  generateNodeDefinition,
  normalizeDocxcelerateApiEndpoint,
  officialDocxcelerateApiEndpoint,
  officialDocxcelerateApiServer,
  scaffoldDocumentProject,
  scaffoldWorkspaceProject,
  type WorkspaceProjectTemplate,
} from "./scaffold.ts";
import { spawn } from "node:child_process";
import { readFile, stat as statPath, writeFile } from "node:fs/promises";
import { dirname, join, relative, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

let inputIterator: AsyncIterator<string> | undefined;
let inputBuffer = "";
let inputEnded = false;

const [rawCommand, ...rawCommandArgs] = process.argv.slice(2);
const { command, commandArgs } = normalizeCommand(rawCommand, rawCommandArgs);

if (!command || command === "help" || command === "--help") {
  printHelp();
  process.exit(0);
}

if (commandArgs.includes("--help") || commandArgs.includes("-h")) {
  printHelp();
  process.exit(0);
}

if (command === "init" || command === "project") {
  const args = parseArgs(commandArgs);
  const guided = args.positionals.length === 0;
  const name = args.positionals[0] ?? await askRequired("Docxcelerate project name");
  const parentDir = args.options.dir;
  const template = await resolveInitTemplate(args, guided);
  const apiEndpoint = await resolveInitApiEndpoint(args, guided);
  const force = args.flags.has("force") ||
    (guided ? await askBoolean("Overwrite existing files if needed?", false) : false);

  const result = await scaffoldWorkspaceProject({ name, parentDir, apiEndpoint, template, force });

  console.log(`Created Docxcelerate project -> ${result.projectDir}`);
  console.log(`Template: ${result.template}`);
  console.log(`API endpoint: ${result.apiEndpoint || "not configured"}`);
  console.log(`Files: ${result.files.length}`);
  await useLocalPackageDependencyWhenAvailable(result.projectDir);
  await runNpmInstall(result.projectDir);
  console.log(`Next: cd ${result.projectDir} && npm run dev`);
  process.exit(0);
}

if (command === "new" || command === "scaffold") {
  const args = parseArgs(commandArgs);
  const guided = args.positionals.length === 0;
  const name = args.positionals[0] ?? await askRequired("Document project name");
  const title = args.options.title ??
    (guided ? await askText("Document title", titleFromName(name)) : undefined);
  const documentsDir = args.options.dir ??
    (guided ? await askText("Documents directory", "documents") : undefined);
  const force = args.flags.has("force") ||
    (guided ? await askBoolean("Overwrite existing files if needed?", false) : false);

  const result = await scaffoldDocumentProject({
    name,
    title,
    documentsDir,
    force,
  });

  console.log(`Created document -> ${result.projectDir}`);
  console.log(`Entrypoint: ${result.entrypoint}`);
  console.log(`Files: ${result.files.length}`);
  process.exit(0);
}

if (command === "node" || command === "generate-node") {
  const args = parseArgs(commandArgs);
  const guided = args.positionals.length < 2;
  const projectDir = args.positionals[0] ??
    await askRequired("Document project directory", "documents/welcome");
  const name = args.positionals[1] ?? await askRequired("Node name");
  const type = args.options.type ??
    (guided
      ? await askChoice("Node type", ["paragraph", "image", "graph"], "paragraph")
      : "paragraph");
  if (type !== "paragraph" && type !== "image" && type !== "graph") {
    fail(`Unsupported node type: ${type}. Expected "paragraph", "image", or "graph".`);
  }
  const mode = args.options.mode ?? args.options.kind ??
    (guided ? await askChoice("Node mode", ["static", "dynamic"], "static") : "static");
  if (mode !== "static" && mode !== "dynamic") {
    fail(`Unsupported node mode: ${mode}. Expected "static" or "dynamic".`);
  }
  const force = args.flags.has("force") ||
    (guided ? await askBoolean("Overwrite existing node file if needed?", false) : false);

  const result = await generateNodeDefinition({
    projectDir,
    name,
    type,
    mode,
    force,
  });

  console.log(`Created ${mode} ${type} node ${result.componentName} -> ${result.filePath}`);
  console.log(`Updated exports -> ${result.exportPath}`);
  process.exit(0);
}

console.error(`Unknown scaffold command: ${command}`);
printHelp();
process.exit(1);

/**
 * Strips the namespace off `dxcl document new` and friends, leaving the bare
 * verb the dispatch below matches on.
 *
 * `letter` is still accepted as a namespace. It was the only spelling before
 * the vocabulary settled on documents, and it costs one array entry to keep
 * every script and habit that already types it working.
 */
function normalizeCommand(
  command: string | undefined,
  args: string[],
): { command: string | undefined; commandArgs: string[] } {
  if (command !== "document" && command !== "letter") {
    return { command, commandArgs: args };
  }

  const [namespaced, ...rest] = args;

  if (!namespaced || namespaced === "help" || namespaced === "--help") {
    return { command: "help", commandArgs: rest };
  }

  return {
    command: namespaced,
    commandArgs: rest,
  };
}

interface ParsedArgs {
  positionals: string[];
  options: Record<string, string>;
  flags: Set<string>;
}

function parseArgs(args: string[]): ParsedArgs {
  const positionals: string[] = [];
  const options: Record<string, string> = {};
  const flags = new Set<string>();

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];

    if (!arg.startsWith("--")) {
      positionals.push(arg);
      continue;
    }

    const withoutPrefix = arg.slice(2);
    const equalsIndex = withoutPrefix.indexOf("=");

    if (equalsIndex !== -1) {
      options[withoutPrefix.slice(0, equalsIndex)] = withoutPrefix.slice(equalsIndex + 1);
      continue;
    }

    const next = args[index + 1];
    if (next && !next.startsWith("--")) {
      options[withoutPrefix] = next;
      index += 1;
      continue;
    }

    flags.add(withoutPrefix);
  }

  return { positionals, options, flags };
}

function fail(message: string): never {
  console.error(message);
  printHelp();
  process.exit(1);
}

async function runNpmInstall(projectDir: string): Promise<void> {
  console.log(`Installing dependencies in ${projectDir}...`);
  const isWindows = process.platform === "win32";
  const code = await new Promise<number>((resolveStatus, rejectStatus) => {
    // Windows needs a shell to run npm.cmd, and passing the command as one
    // string (rather than an args array) avoids Node's DEP0190 warning.
    const child = isWindows
      ? spawn("npm.cmd install", { cwd: projectDir, stdio: "inherit", shell: true })
      : spawn("npm", ["install"], { cwd: projectDir, stdio: "inherit" });

    child.on("error", rejectStatus);
    child.on("close", (exitCode) => resolveStatus(exitCode ?? 1));
  }).catch(() => 1);

  if (code !== 0) {
    exitWithError(`npm install failed in ${projectDir}.`);
  }
}

function exitWithError(message: string): never {
  console.error(message);
  process.exit(1);
}

async function useLocalPackageDependencyWhenAvailable(projectDir: string): Promise<void> {
  const dependency = await localPackageDependency(projectDir);
  if (!dependency) {
    return;
  }

  const packageJsonPath = join(projectDir, "package.json");
  const packageJson = JSON.parse(await readFile(packageJsonPath, "utf8"));
  packageJson.dependencies = packageJson.dependencies ?? {};
  packageJson.dependencies.docxcelerate = dependency;
  await writeFile(packageJsonPath, `${JSON.stringify(packageJson, null, 2)}\n`, "utf8");
  console.log(`Using local docxcelerate package: ${dependency}`);
}

async function localPackageDependency(projectDir: string): Promise<string | undefined> {
  let packageRoot: string;

  try {
    packageRoot = dirname(dirname(dirname(fileURLToPath(import.meta.url))));
  } catch {
    return undefined;
  }

  try {
    const packageJson = JSON.parse(
      await readFile(join(packageRoot, "package.json"), "utf8"),
    ) as { name?: string };

    if (packageJson.name !== "docxcelerate") {
      return undefined;
    }

    // Only a repo checkout carries TypeScript sources; the published package
    // ships compiled output, so this keeps `file:` links out of installs.
    await statPath(join(packageRoot, "src", "document.ts"));
  } catch {
    return undefined;
  }

  const relativePath = relative(resolve(projectDir), resolve(packageRoot)).replaceAll("\\", "/");
  return `file:${relativePath === "" ? "." : relativePath}`;
}

async function resolveInitApiEndpoint(
  args: ParsedArgs,
  guided: boolean,
): Promise<string | undefined> {
  const explicitEndpoint = args.options["api-endpoint"] ?? args.options.endpoint;
  const useOfficial = args.flags.has("official-server");
  const noEndpoint = args.flags.has("no-api-endpoint");
  const selectedCount = [explicitEndpoint !== undefined, useOfficial, noEndpoint]
    .filter(Boolean).length;

  if (selectedCount > 1) {
    fail("Choose only one of --api-endpoint, --official-server, or --no-api-endpoint.");
  }

  if (explicitEndpoint !== undefined) {
    return normalizeDocxcelerateApiEndpoint(explicitEndpoint);
  }

  if (useOfficial) {
    return officialDocxcelerateApiEndpoint;
  }

  if (noEndpoint) {
    return "";
  }

  if (!guided) {
    return undefined;
  }

  console.log(`Official Docxcelerate server: ${officialDocxcelerateApiServer}`);
  const target = await askChoice(
    "Docxcelerate API target",
    ["official", "custom", "none"],
    "official",
  );

  if (target === "official") {
    return officialDocxcelerateApiEndpoint;
  }

  if (target === "none") {
    return "";
  }

  return normalizeDocxcelerateApiEndpoint(
    await askRequired("Docxcelerate API endpoint (server root or /api/letters URL)"),
  );
}

async function resolveInitTemplate(
  args: ParsedArgs,
  guided: boolean,
): Promise<WorkspaceProjectTemplate> {
  const useSample = args.flags.has("sample");
  const useBlank = args.flags.has("blank");

  if (useSample && useBlank) {
    fail("Choose only one of --sample or --blank.");
  }

  if (useBlank) {
    return "blank";
  }

  if (useSample || !guided) {
    return "sample";
  }

  return await askChoice("Project template", ["sample", "blank"], "sample");
}

async function askRequired(label: string, defaultValue?: string): Promise<string> {
  while (true) {
    const value = await askText(label, defaultValue);

    if (value.trim() !== "") {
      return value;
    }

    console.log(`${label} is required.`);
  }
}

async function askText(label: string, defaultValue?: string): Promise<string> {
  const suffix = defaultValue ? ` [${defaultValue}]` : "";
  process.stdout.write(`${label}${suffix}: `);

  const line = await readLine();
  if (line === null) {
    fail(`Expected ${label.toLowerCase()}.`);
  }

  const value = line.trim();
  return value === "" ? defaultValue ?? "" : value;
}

async function askBoolean(label: string, defaultValue: boolean): Promise<boolean> {
  const hint = defaultValue ? "Y/n" : "y/N";

  while (true) {
    const value = (await askText(`${label} (${hint})`)).toLowerCase();

    if (value === "") {
      return defaultValue;
    }

    if (value === "y" || value === "yes") {
      return true;
    }

    if (value === "n" || value === "no") {
      return false;
    }

    console.log("Please answer yes or no.");
  }
}

async function askChoice<const TChoice extends string>(
  label: string,
  choices: readonly TChoice[],
  defaultValue: TChoice,
): Promise<TChoice> {
  while (true) {
    const value = (await askText(`${label} (${choices.join("/")})`, defaultValue)).toLowerCase();
    const choice = choices.find((item) => item === value);

    if (choice) {
      return choice;
    }

    console.log(`Please choose one of: ${choices.join(", ")}.`);
  }
}

async function readLine(): Promise<string | null> {
  if (!inputIterator) {
    process.stdin.setEncoding("utf8");
    inputIterator = process.stdin[Symbol.asyncIterator]() as AsyncIterator<string>;
  }

  while (true) {
    const newlineIndex = inputBuffer.indexOf("\n");
    if (newlineIndex !== -1) {
      const line = inputBuffer.slice(0, newlineIndex).replace(/\r$/, "");
      inputBuffer = inputBuffer.slice(newlineIndex + 1);
      return line;
    }

    if (inputEnded) {
      if (inputBuffer === "") {
        return null;
      }

      const line = inputBuffer.replace(/\r$/, "");
      inputBuffer = "";
      return line;
    }

    const result = await inputIterator.next();
    if (result.done) {
      inputEnded = true;
      continue;
    }

    inputBuffer += result.value;
  }
}

function titleFromName(name: string): string {
  return name
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .split("-")
    .filter(Boolean)
    .map((part) => `${part.charAt(0).toUpperCase()}${part.slice(1)}`)
    .join(" ");
}

function printHelp(): void {
  console.log(`Docxcelerate project scaffolding

Usage:
  dxcl init [project-name] [--dir <parent-dir>] [--sample|--blank] [--official-server|--api-endpoint <url>|--no-api-endpoint]
  dxcl document new [name] [--title <title>] [--dir documents]
  dxcl document node [project-dir] [name] [--type paragraph|image|graph] [--mode static|dynamic]

Examples:
  dxcl init
  dxcl init housing-documents --blank
  dxcl init housing-documents --official-server
  dxcl init housing-documents --api-endpoint ${officialDocxcelerateApiServer}
  dxcl document new
  dxcl document new arrears-notice --title "Arrears Notice"
  dxcl document node
  dxcl document node documents/arrears-notice repayment-summary --type paragraph --mode dynamic

Aliases:
  dxcl new       -> dxcl document new
  dxcl node      -> dxcl document node
  dxcl letter .. -> dxcl document ..   (the old spelling, still accepted)
`);
}
