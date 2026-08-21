#!/usr/bin/env node
// Renders directory submissions from .launch/project.json, and remembers which
// ones are done.
//
// The problem this solves is not typing speed. It is drift: the same tagline
// retyped into twenty forms becomes twenty slightly different taglines, and by
// the time you notice, the wrong one is the one indexed on four aggregators and
// there is no way to edit half of them. So the copy lives in one file, every
// form is rendered from it, and a form is never the place a sentence is
// written for the first time.
//
// What this deliberately does not do is submit anything. Every destination in
// targets.json is a web form behind a login or a pull request against someone
// else's repository — there is no API to post to, and a script that pretended
// otherwise would just be a worse browser. It renders the answers and tracks
// the state; a human does the pasting.
//
//   node scripts/submit.mjs check          validate the copy against every limit
//   node scripts/submit.mjs list           every target, with what has shipped
//   node scripts/submit.mjs next           what to do next, highest value first
//   node scripts/submit.mjs render <id>    the copy pack for one destination
//   node scripts/submit.mjs render --all   every copy pack, for a working session
//   node scripts/submit.mjs done <id> [url]  record a submission
//   node scripts/submit.mjs undo <id>      un-record one
//   node scripts/submit.mjs checklist      the outstanding list, as markdown

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join, resolve } from "node:path";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const launch = join(root, ".launch");

const project = readJson(join(launch, "project.json"));
const { targets } = readJson(join(launch, "targets.json"));

// State is its own file rather than a field inside targets.json, so that a
// change to the destination list and a change to what has shipped never collide
// in the same diff.
const statePath = join(launch, "state.json");
const state = existsSync(statePath) ? readJson(statePath) : { submitted: {} };

function readJson(path) {
  return JSON.parse(readFileSync(path, "utf8"));
}

// Dot-path lookup, so a target names "urls.repo" and never holds a copy of it.
function field(path) {
  return path.split(".").reduce((at, key) => (at == null ? at : at[key]), project);
}

const bold = (s) => `\x1b[1m${s}\x1b[0m`;
const dim = (s) => `\x1b[2m${s}\x1b[0m`;
const red = (s) => `\x1b[31m${s}\x1b[0m`;
const green = (s) => `\x1b[32m${s}\x1b[0m`;
const yellow = (s) => `\x1b[33m${s}\x1b[0m`;

function byId(id) {
  const found = targets.find((t) => t.id === id);
  if (!found) {
    console.error(red(`No target "${id}".`));
    console.error(dim(`Known: ${targets.map((t) => t.id).join(", ")}`));
    process.exit(1);
  }
  return found;
}

// A field is rendered the way the form wants to receive it: a list of
// alternatives is one line per row because that is how they are entered, and an
// asset is a path you drag from, not a value you paste.
function renderValue(path, value) {
  if (value == null) return dim("(not set)");
  if (Array.isArray(value)) {
    if (value.length === 0) return yellow("(empty — nothing to submit here yet)");
    return value
      .map((item) =>
        typeof item === "object"
          ? `  • ${item.name}${item.url ? dim(` — ${item.url}`) : ""}\n    ${item.why ?? ""}`.trimEnd()
          : `  • ${item}`,
      )
      .join("\n");
  }
  if (path.startsWith("assets.")) {
    const exists = existsSync(join(root, value));
    return `${value} ${exists ? green("✓ on disk") : red("✗ missing")}`;
  }
  return String(value).replace(/\\n/g, "\n");
}

// Length is checked against the destination's cap, not a house style, because
// the only length that matters is the one the form silently truncates at.
function lengthNote(target, path, value) {
  if (typeof value !== "string") return "";
  const cap = target.limits?.[path];
  const n = value.length;
  if (!cap) return dim(`${n} chars`);
  return n > cap
    ? red(`${n}/${cap} chars — OVER by ${n - cap}`)
    : green(`${n}/${cap} chars`);
}

function renderTarget(target) {
  const done = state.submitted[target.id];
  const label = { pr: "pull request", form: "web form", post: "post", auto: "automatic" }[target.type];

  console.log("");
  console.log(bold(`── ${target.name} `.padEnd(72, "─")));
  console.log(`${dim("type")}  ${label}   ${dim("priority")}  ${target.priority}`);
  console.log(`${dim("go to")} ${target.url}`);
  if (done) console.log(green(`done  ${done.at}${done.url ? ` — ${done.url}` : ""}`));
  console.log("");

  for (const path of target.fields) {
    const value = field(path);
    console.log(`${bold(path)}  ${lengthNote(target, path, value)}`);
    console.log(renderValue(path, value));
    console.log("");
  }

  // The awesome-list line is the whole submission, so it is assembled here
  // rather than left as three fields for a human to reformat identically twice.
  if (target.type === "pr") {
    console.log(bold("paste this line:"));
    console.log(`- [${project.name}](${project.urls.repo}) - ${project.taglineShort}`);
    console.log("");
  }

  if (target.id === "show-hn") {
    const title = `Show HN: ${project.name} – ${project.taglineShort.replace(/\.$/, "")}`;
    console.log(`${bold("title:")} ${title}  ${title.length > 80 ? red(`${title.length}/80 — OVER`) : green(`${title.length}/80`)}`);
    console.log("");
  }

  if (target.notes) console.log(dim(wrap(target.notes, 76)));
}

function wrap(text, width) {
  const out = [];
  let line = "";
  for (const word of text.split(/\s+/)) {
    if (line.length + word.length + 1 > width) {
      out.push(line);
      line = word;
    } else {
      line = line ? `${line} ${word}` : word;
    }
  }
  if (line) out.push(line);
  return out.join("\n");
}

// Two kinds of finding, and only one of them is a build failure.
//
// Broken means the copy cannot be submitted as it stands: a sentence past a cap
// the form will truncate, a field a target names that project.json does not
// have, an asset path pointing at a file that has moved. Those are defects in
// something already written, they appear the moment someone edits the manifest,
// and CI should stop the pull request.
//
// Unready means a destination is waiting on work nobody has done yet — most of
// the time, screenshots that do not exist. That is a true statement about the
// project, but it is not a reason to fail every unrelated pull request until
// somebody opens a screenshot tool. It gets reported and exits zero.
function check() {
  const broken = [];
  const unready = [];

  for (const target of targets) {
    for (const path of target.fields) {
      const value = field(path);

      if (value === undefined) {
        broken.push(`${target.id}: project.json has no "${path}"`);
        continue;
      }

      const cap = target.limits?.[path];
      if (cap && typeof value === "string" && value.length > cap) {
        broken.push(`${target.id}: "${path}" is ${value.length} chars, cap is ${cap}`);
      }

      if (path.startsWith("assets.") && typeof value === "string" && !existsSync(join(root, value))) {
        broken.push(`${target.id}: asset "${value}" is not on disk`);
      }

      if (Array.isArray(value) && value.length === 0) {
        unready.push(`${target.id}: "${path}" is empty — the form asks for it`);
      }
    }
  }

  // Deduplicated: one missing asset referenced by six targets is one problem.
  const uniq = (list) => [...new Set(list)];

  if (broken.length > 0) {
    console.log(red(`${uniq(broken).length} thing(s) to fix:`));
    for (const p of uniq(broken)) console.log(`  ${red("✗")} ${p}`);
    process.exitCode = 1;
  } else {
    console.log(green("Every target's copy fits, and every asset it names exists."));
  }

  if (unready.length > 0) {
    console.log("");
    console.log(yellow(`${uniq(unready).length} destination(s) not ready to submit:`));
    for (const p of uniq(unready)) console.log(`  ${yellow("○")} ${p}`);
  }
}

function list() {
  const width = Math.max(...targets.map((t) => t.name.length));
  for (const target of [...targets].sort((a, b) => a.priority - b.priority)) {
    const done = state.submitted[target.id];
    const mark = done ? green("✓") : target.type === "auto" ? dim("·") : yellow("○");
    const trail = done ? dim(done.at.slice(0, 10)) : dim(target.type);
    console.log(`${mark} ${target.name.padEnd(width)}  ${dim(`p${target.priority}`)}  ${trail}  ${dim(target.id)}`);
  }
}

function next() {
  const todo = targets
    .filter((t) => !state.submitted[t.id] && t.type !== "auto")
    .sort((a, b) => a.priority - b.priority);

  if (todo.length === 0) {
    console.log(green("Nothing outstanding."));
    return;
  }

  console.log(bold(`${todo.length} outstanding, highest value first:`));
  for (const target of todo.slice(0, 5)) {
    console.log("");
    console.log(`  ${bold(target.name)} ${dim(`(${target.id})`)}`);
    console.log(`  ${target.url}`);
    console.log(dim(wrap(`  ${target.notes ?? ""}`, 76)));
  }
  console.log("");
  console.log(dim(`Render one with:  node scripts/submit.mjs render ${todo[0].id}`));
}

// Markdown rather than the coloured output above, because this one is read in a
// GitHub issue. Ticking a box there does not write back to state.json — `done`
// is still what records a submission — so the issue is a prompt, not a second
// source of truth about what has shipped.
function checklist() {
  const version = JSON.parse(readFileSync(join(root, "package.json"), "utf8")).version;
  const todo = targets
    .filter((t) => !state.submitted[t.id] && t.type !== "auto")
    .sort((a, b) => a.priority - b.priority);

  console.log(`Rendered from \`.launch/\` at v${version}. Copy for each one:`);
  console.log("");
  console.log("```sh");
  console.log("node scripts/submit.mjs render <id>");
  console.log("```");
  console.log("");

  if (todo.length === 0) {
    console.log("Nothing outstanding.");
    return;
  }

  let priority = null;
  for (const target of todo) {
    if (target.priority !== priority) {
      priority = target.priority;
      console.log("");
      console.log(`**Priority ${priority}**`);
      console.log("");
    }
    console.log(`- [ ] [${target.name}](${target.url}) — \`${target.id}\` (${target.type})`);
  }

  console.log("");
  console.log("Record each one as it lands, so the next release does not re-list it:");
  console.log("");
  console.log("```sh");
  console.log("node scripts/submit.mjs done <id> <url-of-the-listing>");
  console.log("```");
}

function done(id, url) {
  const target = byId(id);
  state.submitted[target.id] = {
    at: new Date().toISOString(),
    ...(url ? { url } : {}),
  };
  writeFileSync(statePath, `${JSON.stringify(state, null, 2)}\n`);
  console.log(green(`Recorded ${target.name}${url ? ` — ${url}` : ""}.`));
}

function undo(id) {
  const target = byId(id);
  delete state.submitted[target.id];
  writeFileSync(statePath, `${JSON.stringify(state, null, 2)}\n`);
  console.log(`Cleared ${target.name}.`);
}

const [command, ...rest] = process.argv.slice(2);

switch (command) {
  case "check":
    check();
    break;
  case "list":
    list();
    break;
  case "next":
    next();
    break;
  case "checklist":
    checklist();
    break;
  case "render":
    if (rest[0] === "--all") targets.forEach(renderTarget);
    else if (rest[0]) renderTarget(byId(rest[0]));
    else {
      console.error(red("render needs a target id, or --all."));
      process.exit(1);
    }
    break;
  case "done":
    if (!rest[0]) {
      console.error(red("done needs a target id."));
      process.exit(1);
    }
    done(rest[0], rest[1]);
    break;
  case "undo":
    if (!rest[0]) {
      console.error(red("undo needs a target id."));
      process.exit(1);
    }
    undo(rest[0]);
    break;
  default:
    console.log(readFileSync(fileURLToPath(import.meta.url), "utf8")
      .split("\n")
      .filter((line) => line.startsWith("//"))
      .map((line) => line.replace(/^\/\/ ?/, ""))
      .join("\n"));
}
