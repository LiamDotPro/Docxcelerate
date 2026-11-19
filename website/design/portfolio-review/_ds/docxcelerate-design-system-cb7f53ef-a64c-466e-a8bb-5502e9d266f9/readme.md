# Docxcelerate — design system

**Docxcelerate is a TypeScript toolkit and CLI for building DOCX document projects.** You author a document as a tree of typed components in JSX, preview it locally in the browser, and publish it to an engine that fills in your data, runs the AI prompts and returns a finished `.docx`. The tagline the brand leads with everywhere: *Documents as components. DOCX as output.*

The product is open source (MIT, npm + JSR), authored by Liam, and split into three surfaces:

| Surface | What it is | Represented here |
| --- | --- | --- |
| **The marketing site** — [docxcelerate.com](https://docxcelerate.com) | Astro + Tailwind 4, five languages, one long homepage plus a docs section | `ui_kits/website/` |
| **The preview app** — `dxcl dev` | A local split view: workspace files beside the document they render, wrapped in Word-shaped chrome | `ui_kits/preview-app/` |
| **The rendered document** | The A4 page the renderers emit — the artefact the recipient actually receives | `components/product/PaperPage.jsx` |
| The engine (hosted / self-hosted) | Generation API. Not yet public, so nothing is invented for it here | — |

## Sources this system was built from

- **GitHub:** <https://github.com/LiamDotPro/Docxcelerate> — branch `main`. The whole of `website/` (Astro components, `src/styles/global.css`, `src/i18n/ui/en.ts`, the demo document workspace, `public/brand/`) and `src/renderer/web_renderer.ts`. Explore it further before designing anything substantial: the site's component comments explain *why* almost every value is what it is, and they are the best available brief.
- **Related:** `LiamDotPro/docxcelerate-engine` (private) — the hosted engine. Not read; nothing here describes engine UI.
- Package pages: [npm](https://www.npmjs.com/package/docxcelerate) · [JSR](https://jsr.io/@docxcelerate/docxcelerate)
- No Figma file, no slide deck and no design-system definition were provided. Everything below is derived from the code.

See `github.md` for the sync record.

---

## CONTENT FUNDAMENTALS

**Voice: a working engineer explaining their own tool, plainly.** Short declaratives. Concrete nouns. No adjectives doing sales work.

- **Person.** "You" for the reader, "we"/"our" only for the project's own decisions ("Take advantage of our engine", "We build components which represent the contents of a document"). Never "I".
- **Casing.** Sentence case for every heading, button and card title — "Write documents like websites", "Publish once. Scale it out.", "See it on npm". Only mono micro-labels are uppercased, and those are letter-spaced (`Authoring`, `THE ENGINE`, `LICENSE`, `PLACEHOLDER`). Product names keep their own casing: `dxcl`, `.docx`, `AGENTS.md`, npm, JSR.
- **Headlines** are two short sentences, often full stops instead of a comma: *"Documents as components. DOCX as output."* Section titles do the same: *"Publish once. Scale it out."*
- **Body copy** explains the mechanism, then the consequence. Example: "Because a document is source, changing a sentence is a pull request that gets diffed, reviewed, and stays attributable a year later when someone asks who altered the arrears wording."
- **Specifics over superlatives.** The site never says "powerful" or "seamless". It says "Node.js 20 or newer", "one report or a hundred thousand is the same call", "renders identically on every run". Where a claim needs a picture, the picture is real code from the demo workspace.
- **Domain examples are British, public-sector and unglamorous** by design: tenancy renewals, arrears wording, offers of admission, repair appointments, policy renewals. Sectors are labelled Education / Housing / Insurance. Currency is £, dates are "12 April", spelling is British ("licence" the noun, "licensed" the verb, "behaviour"), and names are diverse (Maya Oyelaran, Dr Priya Raman, Mr Achterberg, Ms Ferreira).
- **CTA style.** A trailing arrow typed into the label: "Get started →", "How the engine works →", "View the source →", "Full documentation →". The quieter twin has no arrow: "How it works", "See it on npm".
- **Honesty markers are part of the voice.** Unshipped things carry a lowercase `soon` pill and are described in the future tense; the docs print a note when a page is only available in English; the demo caption says "no screenshots and no hidden tricks".
- **Emoji: never.** Not in the UI, not in headings, not in the README. Unicode arrows, a middot, a check mark and a caret do all the ornamental work.
- **Numbers** are zero-padded and tabular in mono (`01 · Authoring`, `v0.4.1`, `sha512-…`).

## VISUAL FOUNDATIONS

**One hue, hairlines, real paper, and no decoration that isn't information.** The system is a green-on-paper editorial layout that shares a screen with a Word document — so it holds two palettes at once and never mixes them.

**Colour.** A single ramp, "Ledger Green" (`--color-ledger-50 … 950`), a deep pine pulled slightly toward blue so it sits correctly beside the cool greys the DOCX renderer emits. Semantic tokens sit on top (`--paper`, `--surface`, `--surface-2`, `--border`, `--border-firm`, `--ink`, `--muted`, `--accent`, `--accent-deep`, `--accent-solid`, `--accent-tint`, `--on-accent`). The accent splits by theme: on paper one value does both text and fills; on a dark ground text goes lighter (`#8ac5ae`) while fills stay dark enough to carry light text (`#1d7357`). Every pairing clears WCAG AA and the contrast figure is written in the source beside the token. Exactly one non-green exists — `--rubric` (#9a3b31), named for the red a scribe used for the parts of a document that were not the text; it marks string literals and removed diff lines, *and nothing else*. `--wash` is a marker-pen highlight used under one licence clause. Two full-bleed bands (the pull quote and the footer) are fixed `ledger-900` in both themes; so is paper, which stays `#ffffff` because a letter preview that dims at night stops reading as paper.

**Type.** No webfonts at all — three system stacks: serif `Iowan Old Style, Palatino Linotype, Palatino, Book Antiqua, Georgia` for every heading and the pull quote; sans `system-ui, Segoe UI, Roboto, Helvetica, Arial`; mono `ui-monospace, Cascadia Code, Cascadia Mono, Consolas, SF Mono, Menlo` for all small type. A fourth stack, `Aptos, Segoe UI, Arial`, exists only for content standing in for a letter. Headings are 600 at `-0.015em` with `text-wrap: balance`; body prose is `--muted`, 0.95–1.125rem at 1.7–1.75, `text-wrap: pretty`. Mono carries every eyebrow (0.65rem, uppercase, `0.16em`), pill, file path and metric. Print type is separate and in points: Cambria 20pt title, 12pt bold headings, Calibri 11pt/1.42 body.

**Layout.** Centred containers — 64rem for a single-column block, 72rem for most sections, 80rem for the demo and the three-up points. Gutters 1.25rem, 2rem from `sm`. Sections are 6rem tall (7rem from `sm`) and alternate paper / surface, separated by a 1px hairline rather than a gap. The header is 3.5rem, sticky, translucent paper with a 4px backdrop blur. Docs use a 13rem left rail, a 46rem measure and a 12rem right rail, both rails sticky at 5rem. Peers are separated by vertical hairlines, not boxed as cards.

**Backgrounds.** Flat colour only. No photography, no illustration, no pattern, no texture, no grain. There is exactly one gradient in the whole system — a faint radial glow behind the pull quote (`--glow-band`) — plus two functional gradients: the 24px ruled grid of the Word workspace and the highlighter wash. The only photographic image anywhere is the author's portrait, in a circle, on the green band.

**Cards and borders.** A card is a 1px `--border` hairline, a 6px radius, a surface or paper ground, and no shadow. Radii run 2 / 3 / 4 / 6 / 8px plus pills for badges; nothing is heavily rounded. A 2px accent edge marks a selected file row or a blockquote — never the side of a card. Coloured left-border cards, tinted callout cards and gradient panels are not part of this brand.

**Shadow.** Shadow means paper or a floating shell, never a card and never a button: `--shadow-shell` under the demo, `--shadow-sheet` under the licence sheet, `--shadow-page` / `--shadow-page-docx` under an A4 page.

**Transparency and blur.** Rare and functional: the sticky header (`paper` at 85% + 4px blur), and white at 10–80% opacity on the fixed green bands. Elsewhere, tint tokens are used instead of alpha.

**Animation.** Colour transitions only, 0.12s (tree rows) to 0.18s (the file panel opening), always plain `ease`. The single transform in the system is a 2px horizontal nudge of a card's arrow on hover; a caret rotates 90° when a document expands. No entrances, no fades on scroll, no bounce, no scale. `prefers-reduced-motion` collapses everything to 0.01ms.

**Hover and press.** Primary buttons darken (`--accent-solid` → `--accent-deep`); secondary buttons fill with `--accent-tint`; tertiary buttons take an accent border and accent text; muted text goes accent; cards take an accent border plus tint. There is no press state — nothing shrinks, dims or lifts. Focus is a 2px `--accent` outline at 2px offset with a 2px radius.

**Imagery vibe.** Cool, desaturated, documentary. Brand exports (banner, social square, avatar) sit on the near-black green `#16211c`. The document surfaces are Word's own cool greys (`#e6e8ed` workspace, `#c7ccd6` page border, `#657287` muted, `#204c91` titlebar) — the one blue in the system, and only ever inside a drawing of Word.

## ICONOGRAPHY

**There is no icon library, and no icon font.** The source ships exactly three inline SVGs, each `currentColor` and sized in a 16–24px box: the GitHub mark (fill, 16px), a sun/theme glyph (stroke 1.6, 24px) and a panel-split glyph (stroke 1.3, 16px). Everything else that looks like an icon is a typed character: `→` on CTAs, `↗` on external links, `▸` for a tree caret, `✓` for a passing check, `−`/`+` for zoom, `·` as a separator, `$` as a shell prompt. Emoji are never used. Agent logos on the homepage are drawn as small monochrome marks per agent.

For a new design: reuse those three glyphs (they are reproduced verbatim in `components/navigation/SiteHeader.jsx` and `guidelines/iconography.html`) and typed characters for the rest. If a genuinely new icon is unavoidable, match the stroke: 1.3–1.6, round caps, no fill, no two-tone. **Do not import Lucide or Heroicons** — a stroke set would immediately read as another product's UI.

The brand mark itself is four sheared rectangles — a stack of text lines with the generated one accented — and lives in `assets/brand/`.

## Assets

`assets/brand/` — `mark.svg` (theme-aware), `mark-light.svg`, `mark-dark.svg`, `mark-mono.svg`, `mark-reversed.svg`, `lockup-light.svg`, `lockup-dark.svg`, `favicon.svg`, `icon-512.png`, `banner.png`, `social-square.png`, `og.png`, `avatar.png`.
`assets/people/liam.webp` — the author's portrait, used in the pull-quote attribution.

All exports are generated in the source repo by `npm run brand` in `website/`; regenerate there rather than editing an SVG here.

---

## Index

| Path | What it holds |
| --- | --- |
| `styles.css` | The entry point: `@import`s only. Consumers link this one file. |
| `tokens/` | `colors.css`, `typography.css`, `spacing.css`, `surfaces.css`, `motion.css`, `base.css` |
| `components/` | The reusable primitives (below) |
| `ui_kits/website/` | docxcelerate.com — homepage + docs page, click-through |
| `ui_kits/preview-app/` | `dxcl dev` — workspace beside the rendered document |
| `guidelines/` | 20 specimen cards: colours, type, spacing, surfaces, motion, brand, iconography |
| `assets/` | Brand marks, social images, portrait |
| `SKILL.md` | Agent-skill entry point for this design system |
| `github.md` | Source-repo association and sync record |

### Components

Grouped by concern; each directory has one `@dsCard` HTML showing its states.

**`components/brand/`** — `Mark`, `Wordmark`
**`components/core/`** — `Button`, `Badge`, `Eyebrow`, `Card`, `CopyField`, `PillTab`
**`components/navigation/`** — `SiteHeader`, `SiteFooter`, `DocsSidebar`, `PageToc`
**`components/content/`** — `PullQuote`, `NumberedSteps`, `DocCard`, `CodeFigure`, `DiffFigure`, `LicenceSheet`
**`components/product/`** — `FileTree`, `PanelBar`, `PaperPage` (with `PaperSection`, `PaperParagraph`, `PaperPlaceholder`), `WordShell`

Every component has a sibling `.d.ts` (props) and `.prompt.md` (what & when, plus a usage example).

### Intentional additions

The source is a site, not a component library, so a few families were factored out of repeated inline markup rather than found as components. Each is a literal extraction — same values, no new design:

- `Button`, `Badge`, `Eyebrow`, `Card` — the button, pill, eyebrow and card class strings repeated across `Home.astro`, `Nav.astro` and `Footer.astro`.
- `CodeFigure`, `DiffFigure` — the figure vocabulary of `PointFigure.astro` / `EngineFigure.astro` / `ScaleFigure.astro`, generalised to take lines as data.
- `PaperPage` family, `WordShell`, `FileTree`, `PanelBar` — the renderer's own HTML/CSS (`src/renderer/web_renderer.ts`) and `HeroDemo.astro`, split into mountable pieces.

Nothing else was added: there is no Toast, Avatar, Tooltip, Modal, Input or Tabs family here, because the product defines none.

### Known substitutions and gaps

- **No font files.** The brand uses system stacks by design, so nothing was substituted and nothing needs uploading. `Iowan Old Style` (Apple) and `Cascadia Code` (Windows) resolve to their next fallback on other platforms — Palatino/Georgia and Consolas/Menlo — which is the intended behaviour, but the serif does look different across platforms. If you want the heading face to be identical everywhere, that is a decision to make, not a bug to fix.
- **Syntax highlighting.** The site highlights code with Shiki at build time (github-light / github-dark-dimmed). The kits approximate it with `CodeFigure` tones.
- **Agent marks** (Claude Code, Cursor, Copilot logos) were not copied — they are third-party marks. The picker uses labels only.
- **No slide template** existed in the sources, so none was invented.
