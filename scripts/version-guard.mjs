/**
 * Refuses a version bump that ships nothing.
 *
 * Merging a version change to main publishes it, to npm and to JSR both. That
 * is the point of the release workflows and it is the right default — but it
 * means a bump made on a branch that only touched the website, or the
 * conformance suite, or a comment, cuts a release whose contents are byte for
 * byte the release before it. Every consumer is asked to upgrade to get
 * nothing, and the number stops meaning anything.
 *
 * So: if `package.json`'s version moved, something a consumer receives has to
 * have moved with it.
 *
 * The other direction is deliberately not checked. Landing a fix without a bump
 * is a normal thing to do — the release workflow stops when the version has not
 * changed, which is how several changes come to ship under one number.
 *
 * ```sh
 * node scripts/version-guard.mjs            # against origin/main
 * node scripts/version-guard.mjs <base-ref>
 * ```
 */
import { execFileSync } from "node:child_process";
import process from "node:process";

const base = process.argv[2] ?? "origin/main";

/**
 * What a consumer actually receives, worked out from the two manifests rather
 * than written down again here — a list that drifts from what is published is
 * worse than no list.
 *
 * `dist/` is the build of `src/`, so `src/` is what a diff sees. The manifests
 * themselves are handled separately: they always change on a bump, so counting
 * them as shipped would make this check pass every time.
 */
function shippedPrefixes(manifest, config) {
  const paths = new Set();

  // Stored without a trailing slash; a name is matched either exactly or as
  // the directory a path sits under, so `LICENSE` and `src` both work.
  const tidy = (entry) => (entry === "dist/" ? "src" : entry.replace(/\/$/, ""));

  for (const entry of manifest.files ?? []) paths.add(tidy(entry));
  for (const entry of config.publish?.include ?? []) paths.add(tidy(entry));

  paths.delete("package.json");
  paths.delete("deno.json");

  return [...paths].sort();
}

function at(ref, path) {
  return execFileSync("git", ["show", `${ref}:${path}`], { encoding: "utf8" });
}

function json(ref, path) {
  return JSON.parse(at(ref, path));
}

/** A manifest with its version taken out, so a bump alone does not read as a change. */
function withoutVersion(manifest) {
  const { version: _version, ...rest } = manifest;
  return JSON.stringify(rest);
}

const head = JSON.parse(execFileSync("git", ["show", "HEAD:package.json"], { encoding: "utf8" }));
const merge = execFileSync("git", ["merge-base", base, "HEAD"], { encoding: "utf8" }).trim();
const before = json(merge, "package.json");

if (head.version === before.version) {
  console.log(`version-guard: still ${head.version}; nothing to check.`);
  process.exit(0);
}

const changed = execFileSync("git", ["diff", "--name-only", merge, "HEAD"], { encoding: "utf8" })
  .split("\n")
  .map((line) => line.trim())
  .filter(Boolean);

const prefixes = shippedPrefixes(head, json("HEAD", "deno.json"));
const shipped = changed.filter((path) =>
  prefixes.some((prefix) => path === prefix || path.startsWith(`${prefix}/`))
);

// A manifest counts too, but only for something other than the number itself:
// a changed export map or a new dependency is a change a consumer receives.
for (const [path, name] of [["package.json", "npm"], ["deno.json", "JSR"]]) {
  if (withoutVersion(json("HEAD", path)) !== withoutVersion(json(merge, path))) {
    shipped.push(`${path} (${name} manifest, beyond the version)`);
  }
}

if (shipped.length > 0) {
  console.log(`version-guard: ${before.version} -> ${head.version}, and it ships:`);
  for (const path of shipped) console.log(`  ${path}`);
  process.exit(0);
}

console.error(
  `version-guard: the version went ${before.version} -> ${head.version}, and nothing a\n` +
    "consumer receives changed with it. Merging this would publish a release identical\n" +
    "to the one before it.\n\n" +
    `What ships: ${prefixes.join(", ")}\n` +
    `What this branch changed:\n${changed.map((path) => `  ${path}`).join("\n")}\n\n` +
    "Either leave the version alone — a change can land without one, and the release\n" +
    "workflow stops when the number has not moved — or include the change that earns it.",
);
process.exit(1);
