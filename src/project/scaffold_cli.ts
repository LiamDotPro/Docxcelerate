/**
 * The `dxcl` command line.
 *
 * Importing this module does nothing; {@linkcode runCli} is what reads
 * `process.argv` and exits. It used to happen at module scope, which made
 * importing an entrypoint enough to end the importing process.
 *
 * To drive the same work from code without a terminal, call
 * {@linkcode scaffoldWorkspaceProject}, {@linkcode scaffoldDocumentProject} or
 * {@linkcode generateNodeDefinition} from `/scaffold` instead.
 *
 * ```sh
 * npx docxcelerate init my-documents
 * npx dxcl document new tenancy-renewal
 * npx dxcl document node documents/tenancy-renewal balance --type paragraph
 * ```
 *
 * @module
 */

import {
  generateNodeDefinition,
  normalizeDocxcelerateApiEndpoint,
  officialDocxcelerateApiEndpoint,
  officialDocxcelerateApiServer,
  scaffoldDocumentProject,
  scaffoldWorkspaceProject,
  type WorkspaceProjectTemplate,
} from "./scaffold.ts";
import {
  findDocumentProjects,
  installRegistryEntry,
  resolveInstallOrder,
} from "../registry/install.ts";
import {
  COMPONENT_CATEGORIES,
  COMPONENTS,
  REGISTRY_THEMES,
  registryEntry,
} from "../registry/mod.ts";
import { spawn } from "node:child_process";
import { readFile, stat as statPath, writeFile } from "node:fs/promises";
import { dirname, join, relative, resolve } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

let inputIterator: AsyncIterator<string> | undefined;
let inputBuffer = "";
let inputEnded = false;

/**
 * Runs the command line.
 *
 * Reads `process.argv`, does the work and exits. Importing this module used
 * to do all of that at module scope, which made `docxcelerate/cli` an
 * entrypoint that ended your process if you imported it. `bin/dxcl.mjs` calls
 * this instead.
 *
 * @returns Nothing; every path through it exits.
 */
export async function runCli(): Promise<void> {
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
    const force = args.flags.has("force") ||
      (guided ? await askBoolean("Overwrite existing node file if needed?", false) : false);

    const result = await generateNodeDefinition({
      projectDir,
      name,
      type,
      force,
    });

    console.log(`Created ${type} node ${result.componentName} -> ${result.filePath}`);
    console.log(`Updated exports -> ${result.exportPath}`);
    process.exit(0);
  }

  if (command === "add") {
    const args = parseArgs(commandArgs);
    const guided = args.positionals.length === 0;
    const refs = guided ? await askForRegistryEntries() : args.positionals;
    const projectDir = await resolveTargetProject(args.options.project, guided);
    const force = args.flags.has("force");

    for (const id of resolveInstallOrder(refs)) {
      const result = await installRegistryEntry({ ref: id, projectDir, force });

      console.log(`Added ${result.kind} ${result.title} -> ${projectDir}`);
      for (const file of result.files) {
        console.log(`  ${file}`);
      }
      for (const note of result.followUp) {
        console.log(`  - ${note}`);
      }
    }

    process.exit(0);
  }

  if (command === "list" || command === "registry") {
    const args = parseArgs(commandArgs);
    const filter = args.positionals[0];

    if (filter !== undefined && filter !== "themes" && filter !== "components") {
      fail(`Unknown registry listing: ${filter}. Expected "themes" or "components".`);
    }

    if (filter !== "components") {
      console.log("Themes");
      for (const entry of REGISTRY_THEMES) {
        console.log(`  ${entry.id.padEnd(18)} ${entry.theme.summary}`);
      }
    }

    if (filter !== "themes") {
      if (filter === undefined) {
        console.log("");
      }

      console.log("Components");
      for (const category of COMPONENT_CATEGORIES) {
        for (const component of COMPONENTS.filter((entry) => entry.category === category)) {
          console.log(`  ${component.id.padEnd(18)} ${component.summary}`);
        }
      }
    }

    console.log("");
    console.log("Add one with: dxcl add <id> [--project documents/<name>]");
    console.log("Browse them at: https://docxcelerate.com/themes and /components");
    process.exit(0);
  }

  if (command === "show") {
    const args = parseArgs(commandArgs);
    const id = args.positionals[0] ?? await askRequired("Theme or component id");
    const entry = registryEntry(id);

    if (entry.kind === "theme") {
      const { theme } = entry;
      console.log(`${theme.title} (theme ${theme.id})`);
      console.log(theme.detail);
      console.log(`Category: ${theme.category}`);
      console.log(`Tags: ${theme.tags.join(", ")}`);
      console.log(`Fonts: ${theme.fonts.join(", ")}`);
      console.log(`Page: ${theme.style.page.size} ${theme.style.page.orientation}`);
      console.log(
        `Body: ${theme.style.typography.bodyFont} ${theme.style.typography.bodySizePt}pt`,
      );
    } else {
      console.log(`${entry.title} (component ${entry.id})`);
      console.log(entry.detail);
      console.log(`Category: ${entry.category}`);
      console.log(`Tags: ${entry.tags.join(", ")}`);
      console.log(`Exports: ${entry.exports.join(", ")}`);
      console.log("Reads:");
      for (const field of entry.dataFields) {
        console.log(`  ${field.path}: ${field.type} — ${field.summary}`);
      }
      console.log("Files:");
      for (const file of entry.files) {
        console.log(`  ${file.target}`);
      }
    }

    console.log("");
    console.log(`Add it with: dxcl add ${entry.id}`);
    process.exit(0);
  }

  console.error(`Unknown scaffold command: ${command}`);
  printHelp();
  process.exit(1);
}

/**
 * Which document project an install lands in.
 *
 * A workspace usually holds one document project, and typing `--project` for
 * the only candidate is the kind of ceremony that makes a tool feel slow. So
 * one is used without asking, several are offered, and none is an error that
 * says what to run instead.
 */
async function resolveTargetProject(
  explicit: string | undefined,
  guided: boolean,
): Promise<string> {
  if (explicit !== undefined) {
    return explicit;
  }

  const projects = await findDocumentProjects(".");

  if (projects.length === 1) {
    return projects[0];
  }

  if (projects.length === 0) {
    fail(
      "No document project here. Run dxcl document new to make one, " +
        "or pass --project with the directory holding document.project.ts.",
    );
  }

  if (!guided) {
    fail(
      `Several document projects here: ${projects.join(", ")}. ` +
        "Say which one with --project.",
    );
  }

  return await askChoice("Document project", projects, projects[0]);
}

/** The catalog, printed so a guided install can be answered without leaving. */
async function askForRegistryEntries(): Promise<string[]> {
  console.log("Themes");
  for (const entry of REGISTRY_THEMES) {
    console.log(`  ${entry.id.padEnd(18)} ${entry.theme.summary}`);
  }

  console.log("");
  console.log("Components");
  for (const component of COMPONENTS) {
    console.log(`  ${component.id.padEnd(18)} ${component.summary}`);
  }

  console.log("");
  const answer = await askRequired("What to add (ids, space separated)");

  return answer.split(/\s+/).filter(Boolean);
}

/**
 * Strips the namespace off `dxcl document new` and friends, leaving the bare
 * verb the dispatch below matches on.
 */
function normalizeCommand(
  command: string | undefined,
  args: string[],
): { command: string | undefined; commandArgs: string[] } {
  if (command !== "document") {
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
  dxcl document node [project-dir] [name] [--type paragraph|image|graph]
  dxcl add [ids...] [--project documents/<name>] [--force]
  dxcl list [themes|components]
  dxcl show <id>

Examples:
  dxcl init
  dxcl init housing-documents --blank
  dxcl init housing-documents --official-server
  dxcl init housing-documents --api-endpoint ${officialDocxcelerateApiServer}
  dxcl document new
  dxcl document new arrears-notice --title "Arrears Notice"
  dxcl document node
  dxcl document node documents/arrears-notice repayment-summary --type paragraph
  dxcl list
  dxcl add slate-report
  dxcl add letterhead signature-block --project documents/arrears-notice
  dxcl show payment-summary

Aliases:
  dxcl new       -> dxcl document new
  dxcl node      -> dxcl document node
  dxcl registry  -> dxcl list
`);
}
