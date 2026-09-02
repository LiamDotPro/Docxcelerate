// Fails when a change would cost the package points on its JSR score.
//
// Two of the score's nine checks are about documentation, and both are things a
// pull request can quietly undo: adding an export without a doc comment, or
// adding an entrypoint without a module doc. Neither breaks a build, neither
// shows up in review, and the score is only visible after publishing. So the
// gate lives here instead.
//
// The other seven are not things a diff can regress. A README, an example, and
// provenance are already in place; the description and the runtime tags are
// package settings on jsr.io, which no commit can reach.
//
// Run it with `npm run jsr:doc`. It needs Deno on the path — it is Deno's own
// documentation linter underneath, which is the same tool JSR scores with.

import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import process from "node:process";

const jsr = JSON.parse(readFileSync(new URL("../deno.json", import.meta.url), "utf8"));
const entrypoints = Object.entries(jsr.exports ?? {});

if (entrypoints.length === 0) {
  fail("deno.json declares no exports.");
}

const failures = [];

// `deno doc --lint` reports three rules. Two are the ones JSR counts symbols
// with; the third, private-type-ref, also fires on types owned by a dependency,
// where there is nothing this repository can change. JSR scores slow types
// separately and passes, so it is reported and not enforced.
const enforced = new Set(["missing-jsdoc", "missing-explicit-type"]);

const lint = deno(["doc", "--lint", ...entrypoints.map(([, path]) => path)]);
const diagnostics = [...lint.output.matchAll(/error\[([a-z-]+)\]: (.+)\r?\n\s*-+> (.+)/g)].map(
  ([, rule, message, location]) => ({ rule, message, location: location.trim() }),
);

// A run that ends badly without saying why is a broken environment, not a clean
// package — usually a Deno too old for `doc --lint`, or an unresolvable import.
if (lint.status !== 0 && diagnostics.length === 0) {
  fail(`deno doc --lint failed without reporting a diagnostic:\n${lint.output}`);
}

for (const diagnostic of diagnostics) {
  if (enforced.has(diagnostic.rule)) {
    failures.push(`${diagnostic.location}\n    ${diagnostic.rule}: ${diagnostic.message}`);
  } else {
    console.log(`note: ${diagnostic.rule} at ${diagnostic.location} — not enforced.`);
  }
}

// A module doc is the one thing the linter above says nothing about, so it is
// read out of the documentation itself.
for (const [name, path] of entrypoints) {
  const json = deno(["doc", "--json", path]);

  if (!json.stdout.trim()) {
    fail(`deno doc --json ${path} produced nothing:\n${json.output}`);
  }

  const [module] = Object.values(JSON.parse(json.stdout).nodes ?? {});

  if (!module?.module_doc?.doc?.trim()) {
    failures.push(
      `${path}\n    entrypoint "${name}" has no module doc. ` +
        "Add a /** ... */ comment tagged @module at the top of the file.",
    );
  }
}

if (failures.length > 0) {
  fail(
    `${failures.length} problem${failures.length === 1 ? "" : "s"} that would lower the JSR ` +
      `score:\n\n${failures.join("\n\n")}\n`,
  );
}

console.log(
  `Every exported symbol across ${entrypoints.length} entrypoints is documented, ` +
    "and every entrypoint has a module doc.",
);

/**
 * Runs Deno and hands back what it said and how it ended.
 *
 * Both streams matter and for different reasons: `doc --json` writes the
 * documentation to stdout, while `doc --lint` writes its diagnostics to stderr.
 * The exit code decides nothing on its own — `--lint` exits non-zero for any
 * finding, including the rule this script only reports — so callers read the
 * output first and treat the status as a fallback.
 */
function deno(args) {
  const result = spawnSync("deno", args, {
    encoding: "utf8",
    // The diagnostics are parsed, so the escape codes that would colour them
    // are only in the way.
    env: { ...process.env, NO_COLOR: "1" },
    // The lint output on a package this size runs past the 1MB default, and a
    // truncated run reads as a clean one.
    maxBuffer: 32 * 1024 * 1024,
  });

  if (result.error) {
    fail(`Could not run deno: ${result.error.message}\nInstall it from https://deno.com.`);
  }

  return {
    stdout: result.stdout ?? "",
    output: `${result.stdout ?? ""}\n${result.stderr ?? ""}`,
    status: result.status,
  };
}

function fail(message) {
  console.error(`\n${message}`);
  process.exit(1);
}
