/**
 * Captures what a publish of the parent package contains.
 *
 * `npm pack --dry-run --json` computes the tarball's shasum and subresource
 * integrity without writing a file — the same values npm puts in the registry
 * as `dist.shasum` and `dist.integrity`. So the checksum the site prints is
 * one a reader can verify for themselves:
 *
 *   npm view docxcelerate dist.integrity
 *
 * On Deno Deploy `npm` is a shim over Deno, and only the everyday subcommands
 * go through it; `npm pack` is not one of them. There the registry answers the
 * same question directly — it is where those values end up, and where the
 * reader checks them — so a failed pack falls back to asking it. Local builds
 * still pack, which is what makes an unpublished change visible before it ships.
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
    // npm writes its progress notices to stderr; only stdout carries the JSON.
    stdio: ["ignore", "pipe", "ignore"],
    maxBuffer: 8 * 1024 * 1024,
  });

  // Be tolerant of anything npm prints around the payload.
  const start = stdout.indexOf("[");
  const end = stdout.lastIndexOf("]");
  if (start < 0 || end < 0) {
    throw new Error("npm pack --json produced no JSON array.");
  }

  const [result] = JSON.parse(stdout.slice(start, end + 1));
  if (!result?.version || !result?.integrity) {
    throw new Error("npm pack --json output is missing version or integrity.");
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
 * The published release, from the registry.
 *
 * Asks for the version in package.json first. A version that has been bumped
 * but not yet published is not there, so fall back to whatever `latest` points
 * at — the site would rather show a release a reader can actually install than
 * fail the build over one that does not exist yet.
 */
async function fromRegistry() {
  const manifest = JSON.parse(await readFile(resolve(PACKAGE_ROOT, "package.json"), "utf8"));
  const { name, version } = manifest;

  let packument = await fetchVersion(name, version);
  if (!packument) {
    console.log(`release: ${name}@${version} is not published — falling back to latest`);
    packument = await fetchVersion(name, "latest");
  }
  if (!packument) {
    throw new Error(`The registry has no published release of ${name}.`);
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

let release;
let source = "npm pack";

try {
  release = pack();
} catch (error) {
  console.log(`release: npm pack is unavailable (${error.message.trim()}) — asking the registry`);
  release = await fromRegistry();
  source = "registry";
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
