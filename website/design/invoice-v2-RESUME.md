# Invoice v2 — where we got to, and what to do next

Updated 2026-08-26 (first written 2026-08-25). Companion to
`invoice-v2-plan.md` (the what and why) and `../scripts/lib/VERIFY-CONTRACT.md`
(the harness spec).

**Branch: `release/0.3.1`.** Committed in four parts — the framework, the
preview renderer, the harness, then the document. Nothing pushed.

---

## The one-line status

```
56 pass / 11 fail / 0 blocked / 0 suspect     I1 clean  I2 clean  I3 clean  I4 clean
G1 PASS   G2 PASS   G3 FAIL   G4 FAIL   G5 PASS   G6 PASS
```

Repo: **559 tests, 0 fail.** `astro check`: **0 errors.** The board opened this
work at 4 pass / 57 fail / 6 blocked.

**Nothing is blocked any more.** The document publishes, the publish-path
artifact packs and is measured, and every objective that can be measured is.

Fourteen of the fifteen framework gaps are done (F1–F14). The one left, **F15
(rounded blocks), was decided against**: the blocks stay square in both engines,
for the reason recorded below.

## The preview and the Word file now agree

This was its own pass, and the numbers moved a long way:

| | before | after |
| --- | --- | --- |
| Preview page one | 1206.67px (a sheet is 1123) | **1122.53px** |
| Charge row, preview vs Word | 53.11 vs 48.00px | **55.05 vs 53.17px** |
| Footer strip, preview vs Word | 73.33 vs ~29px | **54 vs 32px** |
| Mean pixel diff, preview vs Word | 16.80% | **9.82%** |
| Mean vertical drift, page one | up to 74px | **8.6px** over 9 regions |

Four causes, all now fixed:

1. **A leading is now stated in points, exactly.** A leading written as a
   multiple is a multiple of different things in the two engines: Word
   multiplies the font's own line height, CSS multiplies the font size, and
   Aptos puts ~22% between those. That was the systematic drift — the same
   content laid out 10.6% apart, compounding down the page. `blockLine` now
   emits `w:line` in twentieths of a point with `w:lineRule="exact"`, which is
   the one form both engines read the same way. The theme's numbers were
   retuned once against the design afterwards.
2. **docx-preview dropped `w:tblInd` entirely** — `parseIndentation` looks for
   `w:left` and `w:tblInd` carries `w:w`. So a footer bar the file says reaches
   the paper's edge stopped at the margin. `docx-page.mjs` writes it back.
3. **docx-preview emits a `<div>` inside a `<p>` for an inline picture**, which
   is invalid HTML. jsdom keeps it; the browser that re-parses the saved page
   closes the paragraph at the picture, strands the words after it, and leaves
   an empty paragraph at the document's default leading. That is what made a
   one-line footer bar 73px deep. `docx-page.mjs` now rewrites the wrapper as
   the `<span>` it should always have been.
4. **The preview screenshots were cropped to the text column**, not the sheet,
   so every page-relative rect laid onto them landed a margin out and the
   cross-engine numbers moved for reasons that had nothing to do with the
   document. They are whole pages now, like the design fixtures and the Word
   raster.

Points 2 and 3 are in `docx-page.mjs` and are the same sanctioned category as
C6's page-number filling: **JS, not CSS, writing back what the file already
says and Word already draws.** Neither invents anything the document does not
declare, and I1 stays clean.

What is left between the two engines is mostly glyph rasterising — solid bars
diff at 0%, text regions at 4–12%. That is the floor for comparing two
typesetters pixel by pixel.

```sh
cd C:/Users/liam/Documents/Docxcelerate/website
npm run verify:invoice           # ~2.5 min
npm run verify:invoice -- --full # adds the variant builds, repo tests, astro check
```

---

## What the 11 failures are

| Cause | Objectives | Status |
| --- | --- | --- |
| **F15 — rounded blocks** | F15.A1, B1, C1, C2 | **Decided: the blocks stay square.** Left failing on purpose — a measured, chosen departure, not a defect. |
| **F13 leftovers** | F13.A3, B1, C1 | Small; see below. |
| **F2.B1** | F2.B1 | Page two's running header draws 33.91px; the objective wants 40–80. Our header is shorter than the design's, not an engine gap. |
| **F10.C2** | F10.C2 | The contract's own two numbers disagree; see below. |
| **Gates** | G3, G4 | Region positions against the *design*, which is the next tranche of work rather than a defect. |

### F15 — rounded blocks: DECIDED, the blocks stay square

**Settled 2026-08-26: the invoice does not round.** The escalation below was
answered — leave both engines square rather than round in Word alone.

This is a second deliberate departure from the canvas, alongside D12's
annotation chips, and it is recorded here for the same reason: so that a reader
comparing the document to the design finds the decision rather than a bug. The
plan's "dropping the radius is not an outcome" was written before the cost was
known; the cost is now known, and this is the call that was made against it.

**The four F15 objectives are left in place and failing on purpose.** They are
an accurate measurement of a departure that was chosen, not a defect anyone
intends to fix. Retiring them is a separate, deliberate edit — and one worth
making only if the board's four permanent red rows start being ignored rather
than read.

The evidence that forced the decision follows. The plan lists three mechanisms
and says to work down them, recording each rejection. All three are also
recorded in `.verify/history.json` under `attempts[]` (gitignored, hence the
copy here). The root cause is one thing: **docx-preview renders no shapes at
all.**

- **Mechanism 1, VML `v:roundrect`** — `parseVmlElement` (`:1254`) switches on
  the local name and handles `rect`, `oval`, `line`, `shape`, `textbox`. Every
  other name returns `null` and is dropped. `roundrect` is one of them.
- **Mechanism 2, DrawingML `prstGeom prst="roundRect"`** — `spPr` is read in
  exactly one place (`:2142`, inside `parsePicture`) and only for the transform.
  `prstGeom` appears nowhere in the package; there is no `wps:wsp` handling.
- **Mechanism 3, a packed picture behind the text** — needs the raster anchored
  with `behindDoc="1"`. docx-preview parses a picture only through
  `parsePicture` and lays it out in normal flow, so the ground would print
  *over* the text — and the pill would lose its selectable text in Word for
  nothing.

Word can do 1 and 2 natively, so the choice was: emit Word-native rounding and
accept that the preview squares it, or leave both square. **Square, in both.**

What decided it was the fidelity pass. Rounding in Word alone would have been
the single place the two engines knowingly disagree, immediately after a pass
spent making them agree everywhere else — a divergence that would stand out
rather than blend into the noise, in the one document held up as the example of
the two matching. A corner radius is not worth spending that on.

If the radius ever becomes worth it, the route is a preview engine that draws
shapes, not a renderer that emits geometry only one of the two can read.

### The small ones

- **F13.C1** — the code measures 108pt wide (exact) inside a 132pt card, but
  132.75pt tall: the card's row is a line taller than the picture. A cell height
  or a zero-leading paragraph around the image would close it.
- **F13.B1** — the design's `qr-canvas` region has no counterpart in our
  document; probe B finds the card but no inner box.
- **F13.A3** — wants a resolved build compared against an unresolved one.
  Now that D14 resolves the QR, the *unresolved* build is the one missing;
  `verify-build.mjs` needs an `--unresolved-image` variant.
- **F10.C2** — `ENGAGEMENT SUMMARY` renders 102.75pt against a 95.33pt design
  target. **The contract's own numbers disagree**: F10.A pins `w:spacing 18`,
  which is 0.12em × 7.5pt, and 7.5pt is what makes the label 102.75pt wide. At
  7pt the width lands but F10.A and F10.C1 both break. One of the two is wrong;
  the fixture is the tiebreak.

---

## What changed this session

### Framework — F1 to F14

| Gap | What it added |
| --- | --- |
| F1 | `<Section showTitle={false}>` — a section keeps its name and prints nothing |
| F2 | `<Document firstHeader/firstFooter>` — Word's `w:titlePg` |
| F3 | `bleed` on a table: negative `tblInd`, margins added back into the columns |
| F4 | A picture inside a paragraph, packed as a run in the same `w:p` |
| F5 | Every cell child takes the column's alignment and the cell's spacing |
| F6 | `font` on a block — a money column in tabular figures |
| F7 | `lineHeight` on a block — a table row set tighter than prose |
| F8 | `blocks.rowAlt` striped by the renderer as it counts rows |
| F9 | `borderSides: []` means "no edges", not "fall through to the default" |
| F10 | `letterSpacingEm` on a text block style |
| F11 | `valign` on a block |
| F12 | **Re-implemented** — the break rides a paragraph *style* |
| F13 | A picture draws its variant, in a single-cell card |
| F14 | `maxWidthMm` — a measure, taken off the right of the text column |
| — | `spacingAfterPt` on a block, so a rule is a strip and not a paragraph |

### Four real defects found by the harness and fixed

1. **Adjacent tables merged in Word.** Suppressing the F1 headings left three
   `w:tbl` back to back; Word reads those as one table and lays the second out
   on the first one's grid. Word saw 3 tables where the file had 6. A hairline
   separator paragraph now goes between them. **docx-preview never showed this**
   — it renders each table separately, so only the real deliverable was wrong.
2. **`firstHeader={false}` silently cost page one its footer.** `w:titlePg` is
   one switch for both strips: turning it on for the header made Word take the
   footer from a `first` part too, and finding none, print nothing. Absent now
   means "like every other page", which is the default part repeated.
3. **The preview stopped paginating.** Batch 1's direct `w:pageBreakBefore` is
   parsed by docx-preview (`:512`), emits no CSS (`:2462`), and is consulted for
   page splitting only off the *style* (`:3139`). The invoice preview on the
   website had become one unpaginated 2062px slab. Now the break rides a style,
   which both engines read.
4. **A data URI with a parameter was not a picture.** `imageSourceOf` read the
   whole head as the media type, so `image/svg+xml;utf8` matched nothing and a
   good picture came back as `[image: …]`. It now parses the parameters.

### Two `docx` package limitations, both pinned by tests

- `docx` builds a paragraph *style*'s properties with the same
  `ParagraphProperties` it uses for a paragraph, so `pageBreakBefore` packs
  correctly — but `IParagraphStylePropertiesOptions` omits it. One narrow
  documented cast; a test asserts the emitted XML.
- `styles.importedStyles` looks like the tidier door and is not:
  `options.styles` **spreads over** the default set (`index.mjs:18293`), so
  supplying it silently drops `docDefaults` and every built-in heading.
  `paragraphStyles` is the additive key.

### Document — C6, D1–D4, D6, D8–D11, D13, D14

- **D13** — the totals, the dates, the per-line arithmetic and the VAT rate all
  moved into derivers. **The document publishes.** Every figure travels as a
  `{{derived.…}}` token and none is baked (verified against `document.json`).
  The charge row became its own component, because a deriver is a hook and the
  callback inside a `.map()` is not a component.
- **D14** — the scan-to-pay QR is a deriver, not a prompt: an EPC069-12 payment
  payload, SVG for the screen and PNG for the `.docx`. No placeholder, so the
  preview shows a real scannable code.
- **C6** — the preview fills `PAGE`/`NUMPAGES` from the laid-out section count
  (docx-preview drops a field run entirely). Footers read `1 / 2` and `2 / 2` in
  both engines. This is JS in `docx-page.mjs`, not CSS: I1 stays clean.
- D1 totals bar, D2 label header rows, D3 the band as one row, D4 the reference
  panel as one cell, D6 the page-one closer, D8 the payment-page letterhead,
  D9 the muted second term, D10 the footer's solid ink, D11 "Terms & notes".
- Page one carries its own rule now: it has no running header, so the line under
  the letterhead belongs to the letterhead.

### Harness — from 19 probe gaps to none

Every objective now reads a fact some probe emits. Added: probe A's
break/heading style resolution, `maxConsecutiveEmptyParas`, `cardTable`, the
muted-note spacing, the footer's drawing-with-text count; probe B's
`pageNumberRightPx`, `trailingEmptyPara` and rendered-text anchor matching;
probe C's `vAlign`, `summary`, `footerTable`, `scanCard`, `suppressedHeadingFinds`,
footer height and visual line count, per-row line counts, the row-3 border, and
amount column indices; probe V's corner tests and per-page footer text (Word
evaluates `PAGE` once per story, so the per-page text exists only in the PDF).
The runner now passes invariants (G6), the variant builds, the border fixture,
the published document and the publish-path docx.

**Objectives corrected where they measured the wrong thing:** F3.B1 compared a
content-relative x against the paper's edge; F3.C1 compared a table's text
inset against its indent; F6.C2 folded three money columns into one spread;
F13.C1 compared a card against a picture; F7.C1 called a page spill
"not measured"; F15's corner test now skips a region with no fill to judge.

---

## Still current

**P0 — fonts.** Aptos Regular/Bold/Italic/BoldItalic copied to
`%LOCALAPPDATA%\Microsoft\Windows\Fonts\` and registered under
`HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`. Per-user and
reversible. Source cache: `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\Aptos\`.

**The design bundle** — `website/design/invoice-v2/`, with `fixtures/` holding
`design-p1.png`, `design-p2.png` and `design-regions.json` (21 regions).

**Housekeeping** — the harness uses port 8901; kill strays after a failed run:

```powershell
Get-Process WINWORD -ErrorAction SilentlyContinue
Get-NetTCPConnection -LocalPort 8901 -State Listen -ErrorAction SilentlyContinue
```

---

## The order of work from here

1. F13's three small ones, and F10.C2's contract contradiction.
2. G3 and G4 — region positions against the design. Both engines agree with
   each other now, so tuning one tunes both.
3. D5 (page two in two columns) and D7 (two names for the sender) are the only
   plan items never started.
4. Update `invoice-v2-plan.md` — it has still not been revised for any of the
   findings, now nine of them.

`npm test` (559) and `astro check` (0 errors) are both green, so the tree
commits cleanly whenever you want a checkpoint.
