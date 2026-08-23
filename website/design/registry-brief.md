# Docxcelerate registry — design brief

## What I need from you

Five surfaces need a visual design pass:

1. `/themes` — the theme gallery
2. `/themes/<id>` — one theme
3. `/components` — the component gallery
4. `/components/<id>` — one component
5. `RegistryCard` — the shared gallery card

The routes, the data, the translated strings and the previews are all built and
working. What exists today is **deliberately unstyled placeholder layout** —
structure put in place so a visual pass had something to replace. The source
files say so in their own comments:

> Deliberately plain. The galleries and the pages under them are laid out for
> structure rather than for effect — the visual design is a separate pass, and
> everything here is meant to be restyled rather than kept.

> The layout is structural, not designed. It exists so the routes, the data and
> the strings are in place for a visual pass to replace.

So: nothing in the current layout is precious. The **content inventory** and the
**constraints** below are what's fixed.

---

## What the registry is

Docxcelerate is a framework for writing DOCX documents as typed JSX components —
letters, notices, reports — where some of the prose is generated per recipient.

The registry is its version of the shadcn/ui model, applied to documents. It
holds two kinds of thing you can install into a document project with one
command:

- **Themes** (5) — fonts, colours, margins, spacing for the packed `.docx`
- **Components** (6) — prebuilt document nodes: a letterhead, an address block,
  a payment summary, a signature

```sh
npx dxcl list                        # print the catalog
npx dxcl show payment-summary        # one entry in full
npx dxcl add slate-report letterhead # install into this project
```

### The one idea the design has to carry

**Installing copies source code into your project. It does not add a dependency.**

`dxcl add letterhead` writes `nodes/letterhead.node.tsx` into your document
project and re-exports it. From that moment it is *your file*: it has no version,
it is never upgraded behind you, and editing it is the expected next step rather
than a fork. What the registry gives you is a first draft and the decisions
already made in it.

That should shape the whole feel. This is not a package index — nobody is
choosing a dependency to trust for the next three years. It is closer to a
pattern book or a set of worked examples. The emotional register is *"here is a
good first draft of the thing you were about to write, and you're about to own
it"*, not *"here is a library, check the download count."*

Two corollaries the design should reflect:

- **There are no versions, no download counts, no stars, no "last updated".**
  Don't design slots for them. The honest metadata is: what it does, what it
  reads from your data, and the code it will drop in your repo.
- **The source code is the product.** Every component page shows the actual file
  that gets copied, read out of the package at build time. It is not an
  illustration of the component — it *is* the component.

---

## The two kinds

### Themes — "what the document looks like"

A theme is data: a page size, four margins, two font families, a five-colour
palette, and a dozen numbers for type sizes and spacing. Installing one writes
the project's `document-style.ts`.

Categories: `Correspondence`, `Report`, `Legal`, `Marketing`

### Components — "a node someone already wrote"

A component is one `.tsx` file exporting one document node. It reads named
fields off the project's data (`account.balanceDue`, `sender.addressLines`) and
returns paragraphs or a section.

Categories, in document order: `Opening`, `Body`, `Generated`, `Closing`, `Legal`

Ids are unique across both kinds, so one lookup and one install command serve
both. `theme:legal-serif` / `component:terms-notice` prefixes exist but are
rarely needed.

---

## Content inventory

Design against these fields. They are all that exists — the site describes
nothing itself, everything is generated from the package at build time. Don't
invent fields; if a design needs something not on this list, flag it.

### Every theme carries

| Field | Example |
| --- | --- |
| `id` | `slate-report` |
| `title` | Slate Report |
| `summary` | one line, for cards |
| `detail` | one paragraph: what it's for, and the decision it makes |
| `category` | Report |
| `tags` | `sans`, `dense`, `corporate`, `skimmable` |
| `fonts` | `Calibri` — every font the theme asks for |
| `style.page` | size (`A4` / `LETTER`), orientation, 4 margins in mm |
| `style.typography` | body font, heading font, body pt, line height, colour |
| `style.palette` | 5 named colours: `heading`, `accent`, `muted`, `rule`, `page` |
| `style.paragraph` | spacing after, in pt |
| `style.title` / `style.sectionHeading` | font size pt, weight, spacing before/after |
| `previewUrl` | a real sample document, rendered in this theme |

### Every component carries

| Field | Example |
| --- | --- |
| `id` | `payment-summary` |
| `title` | Payment summary |
| `summary` | one line, for cards |
| `detail` | one paragraph: what it does, and the decision already made in it |
| `category` | Body |
| `tags` | `money`, `branching`, `section` |
| `exports` | `PaymentSummary` — the component names |
| `dataFields[]` | `{ path, type, summary }` — what it reads off your data |
| `files[]` | `{ source, target, code }` — the full source that gets copied |
| `previewData` | the JSON the preview was built against |
| `resolved[]` | the document nodes it produced — what a renderer is handed |
| `requires[]` | other entries installed alongside (all empty today) |
| `themeHint` | the theme it was drawn against, if any |
| `previewUrl` | the component, rendered on its own |

---

## Real content to design against

Use this copy in the artboards rather than lorem ipsum — the length and tone of
these strings is the design problem.

### The five themes

**Clean Minimal** · `clean-minimal` · Correspondence · Aptos + Cambria · A4 11pt
Palette: heading `#111827`, accent `#2F5FBD`, muted `#6B7280`, rule `#D1D5DB`, page `#FFFFFF`
> A4, one-inch margins, Aptos over Cambria. The default.

> What a document looks like when nobody has decided what it should look like — which is the right answer more often than it sounds. Nothing here competes with the words: one weight of grey, no rules, no accent, and spacing loose enough to read on a screen without turning two pages into three.

**Slate Report** · `slate-report` · Report · Calibri · A4 10.5pt
Palette: heading `#334155`, accent `#0F766E`, muted `#64748B`, rule `#CBD5E1`, page `#FFFFFF`
> Dense, sober, built to be skimmed. Small-cap headings on slate.

> For documents that are consulted rather than read: performance packs, board papers, anything with more sections than a reader intends to finish. Margins are narrower and leading tighter than a letter's, which buys about a fifth more on the page, and the headings carry the navigation — set small, in capitals, in a colder ink than the body.

**Warm Letter** · `warm-letter` · Correspondence · Georgia · A4 11.5pt
Palette: heading `#42301F`, accent `#A2571B`, muted `#8A7A6A`, rule `#E2D6C7`, page `#FFFDF9`
> Georgia on wide margins. Correspondence that reads like a person.

> A serif face, loose leading, and margins wide enough to shorten the line to something the eye finishes comfortably. Use it where the tone of the document matters as much as its contents — renewals, apologies, refusals — and where a reader should not feel they have been sent a form.

**Legal Serif** · `legal-serif` · Legal · Times New Roman · A4 12pt
Palette: heading `#000000`, accent `#000000`, muted `#3F3F3F`, rule `#000000`, page `#FFFFFF`
> Times New Roman at 12pt, all black. Convention, on purpose.

> Notices, terms, statements of case — anything that has to look like what it is at a glance and survive being photocopied twice. Nothing carries meaning by colour: headings are distinguished by weight and capitals alone, so the document reads the same in monochrome as in full colour.

**Bold Brief** · `bold-brief` · Marketing · Verdana · US Letter 12pt
Palette: heading `#141414`, accent `#C2185B`, muted `#6E6E6E`, rule `#E0E0E0`, page `#FFFFFF`
> US Letter, an outsized accent title, one page and one point.

> Briefing notes, one-pagers, anything summarised for somebody who will not turn the page. The title is set large and in the accent so the subject is legible before the document is picked up, and body text sits at 12pt on a short measure so a paragraph is over before attention is.

### The six components

**Letterhead** · `letterhead` · Opening · exports `Letterhead` · resolves to 3 nodes
Tags: letter, header, static. Reads: `sender.name`, `sender.addressLines`, `sentOn`
> Who sent this, from where, and when.

> The block at the top of the page. The address is one paragraph with the lines joined rather than one node per line, so a sender with two address lines and one with five both come out as a block instead of a ragged run of nodes.

**Recipient block** · `recipient-block` · Opening · exports `RecipientBlock` · resolves to 3 nodes
Tags: letter, address, branching. Reads: `recipient.name`, `recipient.addressLines`, `recipient.formalName`
> The address, and a greeting that survives a missing name.

> The salutation is the half worth having. A document generated in bulk meets recipients whose names it does not have — a joint tenancy, a company, a record where the field was never filled in — so the greeting branches, and the fallback is formal rather than clever.

**Payment summary** · `payment-summary` · Body · exports `PaymentSummary` · resolves to 1 node
Tags: money, branching, section. Reads: `account.reference`, `account.balanceDue`, `account.dueBy`, `account.currency`
> What is owed, by when — and what to say when nothing is.

> Three outcomes, three ids: in credit, clear, or owing. Branching on the amount rather than papering over it with one sentence that reads oddly at zero is the point — somebody who owes nothing should not be given a payment deadline.

**Next steps** · `next-steps` · Generated · exports `NextSteps` · resolves to 1 node
Tags: dynamic, prompts, ai. Reads: `actions`, `contact`
> A generated paragraph with all four prompts already fenced off.

> The node most likely to invent something, so all four prompts are set: the info prompt hands the engine the facts it may use, and the negative prompt closes off the two failures that matter in a document somebody acts on — an invented deadline and an invented way to contact you.

**Signature block** · `signature-block` · Closing · exports `SignatureBlock` · resolves to 4 nodes
Tags: letter, image, closing. Reads: `signatory.name`, `signatory.role`, `signatory.signatureImage`, `signatory.closing`
> A closing, a signature image, and who signed it.

> The image is optional and the name is not, which is the right way round: a letter signed by nobody is a letter nobody owns, whereas a missing image is a rendering detail. Without one the block closes with the typed name.

**Terms notice** · `terms-notice` · Legal · exports `TermsNotice` · resolves to 1 node · drawn against Legal Serif
Tags: legal, boilerplate, section. Reads: `terms.heading`, `terms.clauses`, `terms.version`
> Numbered small print, one paragraph and one stable id per clause.

> Boilerplate is the part of a document nobody rereads and everybody copies, so it is worth one node that owns it. Clauses arrive as data, each becomes its own paragraph, and each id is stable — which is what lets next year's terms be diffed against this year's clause by clause.

Note the shape: **six components, five in two categories.** The galleries are
small. A design that only works at forty cards is the wrong design; this needs
to look deliberate at six.

---

## The surfaces, and the design problem in each

### 1. `/themes` and `/components` — the galleries

**Now:** a page header (h1, lead paragraph, the `npx dxcl list …` command in
mono), then category groups. Each group is a tiny uppercase mono label over a
1/2/3-column grid of cards. A single text link at the bottom crosses to the
other gallery.

**Cards now:** title link, summary, five palette swatches (themes only), a
mono meta line (`Calibri · A4 10.5pt`, or the export names), and the install
command in mono at the bottom.

**The problems:**

- **Five white A4 documents look identical at card size.** This is the hardest
  problem on the page. The differences between themes are real and matter —
  Warm Letter's page is cream, Slate Report fits a fifth more on a page, Bold
  Brief is US Letter with a huge pink title — but none of that survives being
  shrunk to a thumbnail. The swatch row is the current attempt and it's weak:
  three of the five palettes are near-identical greys, and Legal Serif's is five
  shades of black. Find a way to make a theme legible at card size that isn't a
  screenshot and isn't five dots.
- **Themes and components are two galleries with one install command.** Right
  now they're two pages joined by a text link at the bottom. Is that right?
  Should there be one browse surface with two sections, or a shared shell? A
  person's actual question is "what can I add?", not "would I like a theme or a
  component today?"
- **Six components in five categories** means most category groups hold one
  card. Grouping may be the wrong organising device at this size.
- **Tags exist on every entry and are used nowhere on the galleries.** They're
  printed as plain text at the bottom of detail pages. `branching`, `dynamic`,
  `money`, `legal` are genuinely how someone would look for these.
- **The install command is the call to action and it's plain text.** It should
  probably be copyable. It appears on every card and every detail page.

### 2. `/themes/<id>` — one theme

**Now:** a back link, header (title, `detail` paragraph, install command), then
four stacked sections — Preview (a live rendered document in an iframe),
Colours (5 swatches with names and hexes), Fonts (the font list, then three
mono spec rows: page / typography / spacing), Category (category · tags).

**The problems:**

- The preview is the whole argument and it's the third thing on the page, below
  a paragraph and a command.
- The spec rows are dense mono strings — `body Calibri 10.5pt / 1.35 · headings
  Calibri · title 22pt bold · section heading 11pt semibold`. Everything a
  reader needs is there and none of it is scannable. A theme *is* data and
  printing it is honest, but this is a wall.
- Nothing lets you compare two themes, which is the actual task. Choosing a
  theme means holding two side by side.

### 3. `/components/<id>` — one component

**Now:** back link, header (title, `detail`, install command, an optional
"drawn against <theme>" link), then five stacked sections — Preview, Reads (a
3-column table: field / type / what it does), Source (the full `.tsx`, syntax
highlighted, with `lands: nodes/letterhead.node.tsx` above it), Resolves to
(the preview data as one long mono line, then the resolved nodes as JSON),
Exports (export names, category · tags).

**The problems:**

- **Five stacked full-width sections is a long scroll with no hierarchy.** The
  preview, the source and the data contract are three different questions asked
  by three different readers at three different moments.
- **"Resolves to" is raw JSON under a `JSON.stringify` one-liner of the preview
  data.** It's showing something genuinely interesting — data in, document nodes
  out — as an unformatted dump. This is the section most in need of a real idea.
- **The `dataFields` table is the practical payload.** After installing, the one
  manual step is adding these paths to your `types.ts`. The install prints them;
  the page tables them. Someone will copy from this. It should be built for
  copying.
- The relationship between preview data → rendered preview → resolved nodes →
  source code is the whole story of the framework, and right now it's four
  unconnected sections in a column.

### 4. The card

Shared between both galleries, with a `swatches` slot themes use and components
don't, and a `meta` slot that carries fonts + page size for themes and export
names for components. Consider whether one card should serve both, or whether a
theme card and a component card are honestly different objects.

---

## Constraints — these are fixed

**Stack.** Astro 5 + Tailwind v4, static output. No client framework. Small
amounts of vanilla JS are fine (the preview already uses a `ResizeObserver`).

**Theming.** Light and dark, via CSS custom properties. Three states: explicit
`data-theme="light"` / `data-theme="dark"` on the root, and unstamped (follow
`prefers-color-scheme`). Every colour must be defined in the base `:root` and
only *redefined* in the dark blocks.

**The site palette is "Ledger Green"** and already exists. Use these tokens
rather than new colours:

```
Light                          Dark
--paper       #ffffff          #0f1a15
--surface     #f7f9f8          #16211c
--surface-2   #eef2f0          #1c2a24
--border      #dfe5e2          #253530
--border-firm #c8d3ce          #33453d
--ink         #16211c          #d7e0da
--muted       #5c6b64          #93a29a
--accent      #17624a          #8ac5ae
--accent-deep #124e3b          #b9ddce
--accent-solid#17624a          #1d7357
--accent-tint #f0f7f4          #16281f
--on-accent   #ffffff          #f0f7f4
--rubric      #9a3b31          #e8998e   (string literals, removed diff lines — nothing else)
```

Fonts: `--font-sans` (system-ui stack), `--font-serif` (Iowan Old Style /
Palatino), `--font-mono` (ui-monospace / Cascadia), and `--font-letter` (Aptos —
reserved for anything representing actual letter content).

**Every pairing must clear WCAG AA.** The existing tokens are annotated with
their contrast ratios; keep that property.

**The previews are iframes and must stay iframes.** The rendered document ships
its own stylesheet — the frame keeps the letter's typography off the site and
the site's off the letter. Two consequences:

- **The paper stays white in both themes.** A letter preview that dims at night
  stops reading as paper, and the toolkit's claim is that this *is* the page.
  There is already a `box-shadow` + `border-radius: 6px` paper treatment; keep
  the intent even if you restyle it.
- **The frame measures its own height on load and on resize.** You can control
  its width, not its aspect ratio. Don't design a fixed-ratio thumbnail grid
  that the content has to fit.
- Each preview has a caption line: "Rendered by Docxcelerate at build time" and
  an "Open ↗" link.

**Five locales: en, de, es, nl, ru.** Both galleries and both detail pages exist
at `/themes/…` and `/<locale>/themes/…`. Two rules:

- **The chrome translates; the entries do not.** Headings, labels, table headers
  and lead paragraphs come from the i18n bundle. Theme and component titles,
  summaries, details and tags stay English in every language — a translation of
  them couldn't be regenerated from the package, so it would drift silently.
  Design for a page that is part translated and part not.
- **German and Russian labels run long.** Don't design fixed-width label
  columns or tight horizontal chips.

**All data is generated at build time** from the package's own catalog by
`website/scripts/build-registry.mjs`. The site is not allowed to describe a
theme or component itself — if the page says it, the package said it first. A
design that needs new prose per entry means new fields in the package catalog;
that's possible, but call it out rather than assuming it.

**Wide content scrolls in its own container.** Code blocks and tables must not
make the page scroll horizontally.

---

## Ranked, the problems worth your attention

1. **Making a theme legible at card size.** Five white documents; real
   differences that don't survive a thumbnail.
2. **The component detail page's four-part story** — preview data in, rendered
   document, resolved nodes, source code — currently four unrelated stacked
   sections.
3. **The install command as a real call to action** across cards and pages.
4. **Whether themes and components are one browse surface or two.**
5. **Making the theme spec (margins, sizes, spacing) scannable** without
   dishonestly hiding it.
6. **The `dataFields` table as something built to be copied**, since that's the
   one manual step after installing.
7. **Tags as a way in**, given they're already on every entry and used nowhere.

## Out of scope

The site chrome — nav, footer, docs pages, home page — is designed and shipped.
Match it; don't redesign it. The nav already carries "Themes" and "Components"
as top-level items alongside Docs, Cloud and GitHub.
