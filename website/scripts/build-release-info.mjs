/**
 * Captures what a publish of the parent package contains.
 *
 * The homepage prints this checksum beside `npm i docxcelerate`, so it has to
 * be the one a reader gets back from:
 *
 *   npm view docxcelerate dist.integrity
 *
 * That makes the registry the source whenever the version in package.json is
 * published — and it must be, because publishing rebuilds the tarball, so a
 * local `npm pack` of the same version hashes differently. Only a version the
 * registry has never seen falls to `npm pack --dry-run --json`, which computes
 * the same fields locally and keeps an unreleased change visible on the site
 * before it ships.
 *
 * Last of all comes whatever `latest` points at: on Deno Deploy `npm` is a
 * shim over Deno, and if `pack` ever stops working through it, a published
 * release is a better answer than a failed build.
 *
 * Output: src/generated/release.json (gitignored — it's derived, not authored).
 */
import { execSync } from "node:child_process";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const PACKAGE_ROOT = resolve(ROOT, "..");
const OUT_FILE = resolve(ROOT, "src/generated/release.json");

const REGISTRY = "https://registry.npmjs.org";

function pack() {
  // Run through a shell: on Windows npm is a .cmd, which Node refuses to
  // spawn directly. The command is a fixed literal, so there's nothing to
  // interpolate or escape.
  const stdout = execSync("npm pack --dry-run --json", {
    cwd: PACKAGE_ROOT,
    encoding: "utf8",
    // npm writes its progress notices to stderr — except on Deno Deploy, where
    // the shim puts them on stdout alongside the JSON. extractResult() reads
    // around them.
    stdio: ["ignore", "pipe", "ignore"],
    maxBuffer: 8 * 1024 * 1024,
  });

  const result = extractResult(stdout);
  if (!result) {
    throw new Error("npm pack --json produced no result array.");
  }

  return {
    name: result.name,
    version: result.version,
    filename: result.filename,
    shasum: result.shasum,
    integrity: result.integrity,
    entryCount: result.entryCount,
    unpackedSize: result.unpackedSize,
  };
}

/**
 * The one result object out of `npm pack --json`, from output that may not be
 * only JSON.
 *
 * npm documents the notices as going to stderr, and locally they do. Deno
 * Deploy's npm is a shim over Deno, and there they land on stdout wrapped
 * around the payload — so slicing from the first `[` to the last `]` picks up
 * a bracket from a notice and parses nothing.
 *
 * Every `[` is tried in turn instead: scan to its balanced close, parse, and
 * accept the first array whose first entry looks like a pack result. Anything
 * else is a notice that happened to contain a bracket.
 */
function extractResult(stdout) {
  for (let start = stdout.indexOf("["); start >= 0; start = stdout.indexOf("[", start + 1)) {
    const end = matchingBracket(stdout, start);
    if (end < 0) continue;

    let parsed;
    try {
      parsed = JSON.parse(stdout.slice(start, end + 1));
    } catch {
      continue;
    }

    const [result] = Array.isArray(parsed) ? parsed : [];
    if (result?.version && result?.integrity) {
      return result;
    }
  }

  return null;
}

/** The index of the `]` closing the `[` at `start`, or -1 if it never closes. */
function matchingBracket(text, start) {
  let depth = 0;
  let inString = false;

  for (let i = start; i < text.length; i += 1) {
    const char = text[i];

    if (inString) {
      if (char === "\\") i += 1;
      else if (char === '"') inString = false;
      continue;
    }

    if (char === '"') inString = true;
    else if (char === "[") depth += 1;
    else if (char === "]" && (depth -= 1) === 0) return i;
  }

  return -1;
}

/** A release as the registry has it, or null if that spec is not published. */
async function fromRegistry(name, spec) {
  const packument = await fetchVersion(name, spec);
  if (!packument) {
    return null;
  }

  const dist = packument.dist ?? {};
  if (!dist.integrity) {
    throw new Error(`${name}@${packument.version} has no dist.integrity in the registry.`);
  }

  return {
    name: packument.name,
    version: packument.version,
    filename: dist.tarball?.split("/").pop() ?? `${name}-${packument.version}.tgz`,
    shasum: dist.shasum ?? null,
    integrity: dist.integrity,
    entryCount: dist.fileCount ?? null,
    unpackedSize: dist.unpackedSize ?? null,
  };
}

async function fetchVersion(name, spec) {
  const response = await fetch(`${REGISTRY}/${encodeURIComponent(name)}/${spec}`, {
    headers: { accept: "application/vnd.npm.install-v1+json, application/json" },
  });
  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    throw new Error(`Registry returned ${response.status} for ${name}@${spec}.`);
  }
  return await response.json();
}

const { name, version } = JSON.parse(
  await readFile(resolve(PACKAGE_ROOT, "package.json"), "utf8"),
);

let source = "the registry";
let release = await fromRegistry(name, version);

if (!release) {
  console.log(`release: ${name}@${version} is not published — packing it instead`);
  try {
    release = pack();
    source = "npm pack";
  } catch (error) {
    console.log(`release: npm pack is unavailable (${error.message.trim()})`);
    release = await fromRegistry(name, "latest");
    source = "the registry, at latest";
  }
}

if (!release) {
  throw new Error(`Found no release of ${name}: not at ${version}, not packable, none published.`);
}

await mkdir(dirname(OUT_FILE), { recursive: true });
await writeFile(OUT_FILE, `${JSON.stringify(release, null, 2)}\n`, "utf8");

console.log(
  `release: ${release.name}@${release.version} — ` +
    (release.entryCount === null || release.unpackedSize === null
      ? ""
      : `${release.entryCount} files, ${(release.unpackedSize / 1024).toFixed(0)} KB unpacked, `) +
    `via ${source}`,
);
