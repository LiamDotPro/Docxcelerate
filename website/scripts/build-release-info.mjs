/**
 * Captures what the next publish of the parent package would actually contain.
 *
 * `npm pack --dry-run --json` computes the tarball's shasum and subresource
 * integrity without writing a file — the same values npm puts in the registry
 * as `dist.shasum` and `dist.integrity`. So the checksum the site prints is
 * one a reader can verify for themselves:
 *
 *   npm view docxcelerate dist.integrity
 *
 * Output: src/generated/release.json (gitignored — it's derived, not authored).
 */
import { execSync } from "node:child_process";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "..");
const PACKAGE_ROOT = resolve(ROOT, "..");
const OUT_FILE = resolve(ROOT, "src/generated/release.json");

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
  return result;
}

const result = pack();

const release = {
  name: result.name,
  version: result.version,
  filename: result.filename,
  shasum: result.shasum,
  integrity: result.integrity,
  entryCount: result.entryCount,
  unpackedSize: result.unpackedSize,
};

await mkdir(dirname(OUT_FILE), { recursive: true });
await writeFile(OUT_FILE, `${JSON.stringify(release, null, 2)}\n`, "utf8");

console.log(
  `release: ${release.name}@${release.version} — ${release.entryCount} files, ` +
    `${(release.unpackedSize / 1024).toFixed(0)} KB unpacked`,
);
