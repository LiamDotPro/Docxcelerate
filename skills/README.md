# Docxcelerate agent skills

A skill teaches a coding agent how to use Docxcelerate properly — the node model,
the hook rules, and the difference between building a document locally and
publishing one to the engine. Without it, agents guess: they add JSX pragmas that
are not needed, read data outside `useState`, reuse ids across branch arms, and
call `.map()` where a `<Repeat>` belongs.

```text
skills/docxcelerate/
  SKILL.md              the skill — mental model, rules, workflows, checklist
  references/
    api.md              every entrypoint, element prop, hook and build function
    patterns.md         copyable recipes for the shapes that come up
    publishing.md       derivers, build artifacts, engine semantics
    cli.md              every dxcl command and flag
```

It is plain Markdown with a YAML header, so it works as an Anthropic-format skill
and as a rules file for any other agent.

## Claude Code

Copy the directory into either skills folder — project-local for a team, personal
for every project:

```sh
# this workspace only
mkdir -p .claude/skills && cp -r path/to/skills/docxcelerate .claude/skills/

# every project on this machine
mkdir -p ~/.claude/skills && cp -r path/to/skills/docxcelerate ~/.claude/skills/
```

Claude picks it up on the next session and loads it when a Docxcelerate workspace
or import shows up in the task. Nothing to enable.

## Claude apps

Settings → Capabilities → Skills → **Upload skill**, with `skills/docxcelerate`
zipped:

```sh
cd skills && zip -r docxcelerate-skill.zip docxcelerate
```

## Any other agent

`SKILL.md` is Markdown, and the front matter is inert to anything that does not
read it. Point your agent's instruction file at it, or paste it in:

| Agent | Where it goes |
| --- | --- |
| Cursor | `.cursor/rules/docxcelerate.mdc`, or `@`-mention `SKILL.md` in chat |
| GitHub Copilot | `.github/copilot-instructions.md` |
| Codex, Gemini CLI, Aider, Cline | `AGENTS.md` at the repo root |
| An API-driven agent | Prepend `SKILL.md` to the system prompt |

For agents that read a single instruction file, a pointer is usually enough — they
will open the referenced files when the task needs them:

```markdown
When working with Docxcelerate, follow skills/docxcelerate/SKILL.md.
Deeper reference lives in skills/docxcelerate/references/.
```

If the agent cannot follow links, concatenate instead:

```sh
cat skills/docxcelerate/SKILL.md skills/docxcelerate/references/*.md > AGENTS.md
```

## Keeping it honest

The skill describes the published package. When the authoring surface changes —
a new element, a new hook, a renamed entrypoint, a new CLI flag — update
`references/` in the same change. A skill that has drifted is worse than no
skill, because an agent will trust it over the code in front of it.
