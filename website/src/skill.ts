/**
 * The agent skill, as the homepage prints it.
 *
 * The skill is Markdown that lives in the toolkit repository, and the homepage
 * prints it verbatim — the whole point of the block is that what a reader copies
 * off the page is the file that ships, so this reads the real one at build time
 * rather than keeping a second copy here to fall out of date.
 *
 * A missing file fails the build instead of falling back to something. There is
 * no version of this page that is better for showing a skill nobody has.
 *
 * Build-time only: the component calls `agentSkill()` in its frontmatter and
 * ships the result as markup, so `node:fs` never reaches the browser.
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

/**
 * The skill's path in the repository. `references/` beside it is the rest of
 * the skill; the homepage prints the file an agent reads first.
 */
const SKILL_PATH = "skills/docxcelerate/SKILL.md";

/**
 * Where to look, in order. Resolved against the working directory, which is
 * this package for `npm run build` and the repository root if the site is ever
 * built from there — both are tried, so neither has to be the one true answer.
 */
const CANDIDATES = [resolve("..", SKILL_PATH), resolve(SKILL_PATH)];

/** The skill's Markdown, read from the repository. */
export function agentSkill(): string {
  for (const path of CANDIDATES) {
    let source: string;
    try {
      source = readFileSync(path, "utf8").trim();
    } catch {
      continue;
    }

    if (source) {
      console.log(`skill: printing ${path}`);
      return source;
    }
  }

  throw new Error(
    `The homepage prints the agent skill and could not find it. Looked in:\n` +
      CANDIDATES.map((path) => `  ${path}`).join("\n"),
  );
}
