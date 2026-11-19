/* @ds-bundle: {"format":4,"namespace":"DocxcelerateDesignSystem_cb7f53","components":[{"name":"Mark","sourcePath":"components/brand/Mark.jsx"},{"name":"Wordmark","sourcePath":"components/brand/Wordmark.jsx"},{"name":"CodeFigure","sourcePath":"components/content/CodeFigure.jsx"},{"name":"DiffFigure","sourcePath":"components/content/DiffFigure.jsx"},{"name":"DocCard","sourcePath":"components/content/DocCard.jsx"},{"name":"LicenceSheet","sourcePath":"components/content/LicenceSheet.jsx"},{"name":"NumberedSteps","sourcePath":"components/content/NumberedSteps.jsx"},{"name":"PullQuote","sourcePath":"components/content/PullQuote.jsx"},{"name":"Badge","sourcePath":"components/core/Badge.jsx"},{"name":"Button","sourcePath":"components/core/Button.jsx"},{"name":"Card","sourcePath":"components/core/Card.jsx"},{"name":"CopyField","sourcePath":"components/core/CopyField.jsx"},{"name":"Eyebrow","sourcePath":"components/core/Eyebrow.jsx"},{"name":"PillTab","sourcePath":"components/core/PillTab.jsx"},{"name":"DocsSidebar","sourcePath":"components/navigation/DocsSidebar.jsx"},{"name":"PageToc","sourcePath":"components/navigation/PageToc.jsx"},{"name":"SiteFooter","sourcePath":"components/navigation/SiteFooter.jsx"},{"name":"SiteHeader","sourcePath":"components/navigation/SiteHeader.jsx"},{"name":"FileTree","sourcePath":"components/product/FileTree.jsx"},{"name":"PanelBar","sourcePath":"components/product/PanelBar.jsx"},{"name":"PaperPage","sourcePath":"components/product/PaperPage.jsx"},{"name":"PaperSection","sourcePath":"components/product/PaperPage.jsx"},{"name":"PaperParagraph","sourcePath":"components/product/PaperPage.jsx"},{"name":"PaperPlaceholder","sourcePath":"components/product/PaperPage.jsx"},{"name":"WordShell","sourcePath":"components/product/WordShell.jsx"}],"sourceHashes":{"components/brand/Mark.jsx":"57c7e9ca0483","components/brand/Wordmark.jsx":"26aa74a0c8aa","components/content/CodeFigure.jsx":"24ac971bf492","components/content/DiffFigure.jsx":"bbcbbca9733c","components/content/DocCard.jsx":"987dc5f04d18","components/content/LicenceSheet.jsx":"a090fb4e8c3d","components/content/NumberedSteps.jsx":"d87fa0de2232","components/content/PullQuote.jsx":"70a2f2be301b","components/core/Badge.jsx":"136833a987ac","components/core/Button.jsx":"541dfd44a20f","components/core/Card.jsx":"fb70828e75f5","components/core/CopyField.jsx":"daf9367f7742","components/core/Eyebrow.jsx":"cb2793977999","components/core/PillTab.jsx":"2e35fa262f5b","components/navigation/DocsSidebar.jsx":"2e242f17b0f9","components/navigation/PageToc.jsx":"31b713f197d8","components/navigation/SiteFooter.jsx":"653c230627f4","components/navigation/SiteHeader.jsx":"3b42dbd2886a","components/product/FileTree.jsx":"149a59fca37d","components/product/PanelBar.jsx":"122626035e2c","components/product/PaperPage.jsx":"04062a6e58c7","components/product/WordShell.jsx":"c0bc6fb848c0","ui_kits/preview-app/PreviewApp.jsx":"8458c0639b0b","ui_kits/website/DocsScreen.jsx":"ff4989074b7a","ui_kits/website/Hero.jsx":"f1016d8a26e9","ui_kits/website/Points.jsx":"b74affb4d17a","ui_kits/website/Sections.jsx":"7a3c97412bf0","ui_kits/website/shared.jsx":"7855b24e3157"},"inlinedExternals":[],"unexposedExports":[]} */

(() => {

const __ds_ns = (window.DocxcelerateDesignSystem_cb7f53 = window.DocxcelerateDesignSystem_cb7f53 || {});

const __ds_scope = {};

(__ds_ns.__errors = __ds_ns.__errors || []);

// components/brand/Mark.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The Docxcelerate mark: four sheared text lines, the generated one accented.
 * Drawn inline (not loaded from assets/brand/mark.svg) so it follows the host
 * theme — three quiet lines in currentColor, one line in the accent.
 * The viewBox is the ink's exact bounding box, so size it by HEIGHT only.
 */
function Mark({
  height = "1.15em",
  accentColor = "var(--accent)",
  title,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("svg", _extends({
    viewBox: "0 0 47.14 43",
    role: title ? "img" : undefined,
    "aria-hidden": title ? undefined : "true",
    style: {
      height,
      width: "auto",
      flex: "none",
      display: "block",
      ...style
    }
  }, rest), title ? /*#__PURE__*/React.createElement("title", null, title) : null, /*#__PURE__*/React.createElement("g", {
    transform: "translate(-2.3094 -12) skewX(-12)"
  }, /*#__PURE__*/React.createElement("rect", {
    x: "14",
    y: "12",
    width: "38",
    height: "7",
    fill: "currentColor",
    opacity: "0.38"
  }), /*#__PURE__*/React.createElement("rect", {
    x: "14",
    y: "24",
    width: "28",
    height: "7",
    fill: "currentColor",
    opacity: "0.55"
  }), /*#__PURE__*/React.createElement("rect", {
    x: "14",
    y: "36",
    width: "42",
    height: "7",
    fill: accentColor
  }), /*#__PURE__*/React.createElement("rect", {
    x: "14",
    y: "48",
    width: "21",
    height: "7",
    fill: "currentColor",
    opacity: "0.55"
  })));
}
Object.assign(__ds_scope, { Mark });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/brand/Mark.jsx", error: String((e && e.message) || e) }); }

// components/brand/Wordmark.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * Mark plus wordmark. "Docx" in ink, "celerate" in the accent — the split is
 * the lockup, and it never changes. On a fixed brand-green ground pass
 * tone="reversed" so both the mark's accent line and the second half of the
 * word take ledger-300 instead of the theme accent.
 */
function Wordmark({
  size = "1.125rem",
  tone = "theme",
  weight = 600,
  as = "span",
  style,
  ...rest
}) {
  const reversed = tone === "reversed";
  const accent = reversed ? "var(--color-ledger-300)" : "var(--accent)";
  const Tag = as;
  return /*#__PURE__*/React.createElement(Tag, _extends({
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: "0.625rem",
      fontFamily: "var(--font-serif)",
      fontSize: size,
      fontWeight: weight,
      letterSpacing: "-0.015em",
      color: reversed ? "#ffffff" : "var(--ink)",
      textDecoration: "none",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement(__ds_scope.Mark, {
    height: "1.35em",
    accentColor: accent
  }), /*#__PURE__*/React.createElement("span", null, "Docx", /*#__PURE__*/React.createElement("span", {
    style: {
      color: accent
    }
  }, "celerate")));
}
Object.assign(__ds_scope, { Wordmark });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/brand/Wordmark.jsx", error: String((e && e.message) || e) }); }

// components/content/CodeFigure.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const toneColor = {
  tag: "var(--muted)",
  ink: "var(--ink)",
  accent: "var(--accent)",
  string: "var(--rubric)",
  muted: "var(--muted)"
};

/**
 * A code figure: mono, unwrapped, no chrome. This is how the brand draws a
 * diagram — the source itself, indented, with tags in muted, component names in
 * the accent and string literals in rubric red. Lines are data, so a figure can
 * be composed rather than hand-marked-up.
 */
function CodeFigure({
  lines = [],
  size = "var(--text-code-figure)",
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      overflowX: "auto",
      fontFamily: "var(--font-mono)",
      fontSize: size,
      lineHeight: 1.75,
      color: "var(--muted)",
      ...style
    }
  }, rest), lines.map((line, i) => /*#__PURE__*/React.createElement("div", {
    key: i,
    style: {
      paddingLeft: `${(line.indent ?? 0) * 0.875}rem`,
      whiteSpace: "pre"
    }
  }, (line.spans ?? [{
    text: line.text ?? "",
    tone: "muted"
  }]).map((span, j) => /*#__PURE__*/React.createElement("span", {
    key: j,
    style: {
      color: toneColor[span.tone] ?? toneColor.muted
    }
  }, span.text)))));
}
Object.assign(__ds_scope, { CodeFigure });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/content/CodeFigure.jsx", error: String((e && e.message) || e) }); }

// components/content/DiffFigure.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * A diff figure. Removed lines sit on rubric-tint, added lines on accent-tint,
 * and each tinted row bleeds 1.5 past the text so a change reads as a band
 * rather than a highlight. Hunks are separated by space alone — a rule between
 * them would be a third weight in a drawing that needs two.
 */
function DiffFigure({
  hunks = [],
  checks = [],
  style,
  ...rest
}) {
  const row = tone => ({
    margin: "0 -0.375rem",
    padding: "0 0.375rem",
    whiteSpace: "pre",
    background: tone === "add" ? "var(--accent-tint)" : "var(--rubric-tint)",
    color: tone === "add" ? "var(--accent)" : "var(--rubric)"
  });
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      margin: "0 -0.375rem",
      padding: "0 0.375rem",
      overflowX: "auto",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-code-figure)",
      lineHeight: 1.75,
      color: "var(--muted)",
      ...style
    }
  }, rest), hunks.map((hunk, i) => /*#__PURE__*/React.createElement("div", {
    key: hunk.file,
    style: {
      marginTop: i === 0 ? 0 : "1.25rem"
    }
  }, /*#__PURE__*/React.createElement("div", null, hunk.file), hunk.removed?.map(text => /*#__PURE__*/React.createElement("div", {
    key: text,
    style: row("remove")
  }, "\u2212 ", text)), hunk.added?.map(text => /*#__PURE__*/React.createElement("div", {
    key: text,
    style: row("add")
  }, "+ ", text)))), checks.length ? /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: "1.25rem"
    }
  }, checks.map(check => /*#__PURE__*/React.createElement("div", {
    key: check
  }, "\u2713 ", check))) : null);
}
Object.assign(__ds_scope, { DiffFigure });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/content/DiffFigure.jsx", error: String((e && e.message) || e) }); }

// components/content/LicenceSheet.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The project's licence drawn as the sheet of paper it is: two blank sheets
 * fanned behind a front sheet that carries real text, the granted permissions
 * marked with a highlighter wash. Tilts are under a degree.
 */
function LicenceSheet({
  licence = "MIT",
  copyright = `Copyright (c) ${new Date().getFullYear()} Docxcelerate`,
  pages = "1 of 1",
  verbs = ["Fork it", "Vendor it", "Ship it"],
  readLabel = "Read it →",
  readHref = "https://opensource.org/licenses/MIT",
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      position: "relative",
      paddingTop: "0.625rem",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("div", {
    "aria-hidden": "true",
    style: {
      position: "absolute",
      top: "1.625rem",
      left: "1.25rem",
      right: "-4px",
      height: "96%",
      borderRadius: "var(--radius-sm)",
      border: "1px solid var(--border)",
      background: "var(--surface-2)",
      transform: "rotate(1.1deg)"
    }
  }), /*#__PURE__*/React.createElement("div", {
    "aria-hidden": "true",
    style: {
      position: "absolute",
      top: "1rem",
      left: "0.625rem",
      right: "0.125rem",
      height: "98%",
      borderRadius: "var(--radius-sm)",
      border: "1px solid var(--border)",
      background: "var(--paper)",
      transform: "rotate(0.4deg)"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative",
      borderRadius: "var(--radius-sm)",
      border: "1px solid var(--border)",
      background: "var(--paper)",
      padding: "1.875rem 2rem 2rem",
      boxShadow: "var(--shadow-sheet)",
      transform: "rotate(-0.9deg)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "baseline",
      justifyContent: "space-between",
      gap: "1rem",
      borderBottom: "1px solid var(--border)",
      paddingBottom: "1.125rem",
      fontFamily: "var(--font-mono)",
      fontSize: "0.6rem",
      color: "var(--muted)"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      letterSpacing: "0.18em"
    }
  }, "LICENSE"), /*#__PURE__*/React.createElement("span", {
    style: {
      letterSpacing: "0.04em",
      opacity: 0.7
    }
  }, pages)), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "1.625rem 0 0",
      fontFamily: "var(--font-serif)",
      fontSize: "3.25rem",
      lineHeight: 0.9,
      fontWeight: 700,
      letterSpacing: "-0.03em"
    }
  }, licence), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "0.875rem 0 0",
      fontFamily: "var(--font-mono)",
      fontSize: "0.66rem",
      color: "var(--muted)"
    }
  }, copyright), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "0.875rem 0 0",
      fontFamily: "var(--font-serif)",
      fontSize: "0.97rem",
      lineHeight: 1.7,
      color: "var(--muted)",
      textWrap: "pretty"
    }
  }, "Permission is hereby granted, free of charge, to any person obtaining a copy of this software,", " ", /*#__PURE__*/React.createElement("span", {
    style: {
      background: "var(--wash-underlay)",
      color: "var(--ink)"
    }
  }, "to use, copy, modify, merge, publish, distribute, sublicense"), " ", "and sell copies of it."), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      alignItems: "center",
      justifyContent: "space-between",
      gap: "0.75rem 1rem",
      marginTop: "1.75rem",
      borderTop: "1px solid var(--border)",
      paddingTop: "1.125rem"
    }
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      alignItems: "center",
      gap: "0.25rem 1rem",
      margin: 0,
      fontSize: "0.845rem",
      color: "var(--ink)"
    }
  }, verbs.map((verb, index) => /*#__PURE__*/React.createElement(React.Fragment, {
    key: verb
  }, index > 0 ? /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true",
    style: {
      color: "var(--border-firm)"
    }
  }, "\xB7") : null, /*#__PURE__*/React.createElement("span", null, verb)))), /*#__PURE__*/React.createElement("a", {
    href: readHref,
    style: {
      fontSize: "0.78rem",
      fontWeight: 600,
      whiteSpace: "nowrap",
      color: "var(--accent)",
      textDecoration: "none"
    }
  }, readLabel))));
}
Object.assign(__ds_scope, { LicenceSheet });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/content/LicenceSheet.jsx", error: String((e && e.message) || e) }); }

// components/content/NumberedSteps.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * A numbered sequence down a spine: filled accent discs joined by a hairline,
 * because the steps are one sequence and the rule is what says so.
 */
function NumberedSteps({
  steps = [],
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("ol", _extends({
    style: {
      display: "flex",
      flexDirection: "column",
      listStyle: "none",
      margin: 0,
      padding: 0,
      ...style
    }
  }, rest), steps.map((step, index) => {
    const last = index === steps.length - 1;
    return /*#__PURE__*/React.createElement("li", {
      key: step.title,
      style: {
        margin: 0,
        display: "grid",
        gridTemplateColumns: "3.25rem 1fr"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center"
      }
    }, /*#__PURE__*/React.createElement("span", {
      style: {
        display: "grid",
        placeItems: "center",
        width: "1.75rem",
        height: "1.75rem",
        flex: "none",
        borderRadius: "var(--radius-pill)",
        background: "var(--accent-solid)",
        color: "var(--on-accent)",
        fontFamily: "var(--font-mono)",
        fontSize: "0.62rem",
        fontVariantNumeric: "tabular-nums"
      }
    }, String(index + 1).padStart(2, "0")), last ? null : /*#__PURE__*/React.createElement("span", {
      "aria-hidden": "true",
      style: {
        margin: "0.5rem 0",
        width: "1px",
        flex: 1,
        background: "var(--border-firm)"
      }
    })), /*#__PURE__*/React.createElement("div", {
      style: {
        minWidth: 0,
        paddingBottom: last ? 0 : "2rem"
      }
    }, /*#__PURE__*/React.createElement("h3", {
      style: {
        margin: 0,
        fontSize: "var(--text-h3)",
        fontWeight: 600
      }
    }, step.title), /*#__PURE__*/React.createElement("p", {
      style: {
        margin: "0.375rem 0 0.875rem",
        fontSize: "var(--text-sm)",
        lineHeight: 1.65,
        color: "var(--muted)",
        textWrap: "pretty"
      }
    }, step.detail), step.figure));
  }));
}
Object.assign(__ds_scope, { NumberedSteps });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/content/NumberedSteps.jsx", error: String((e && e.message) || e) }); }

// components/content/PullQuote.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The thesis band: fixed ledger-900 in both themes, a faint radial glow behind
 * the type, serif quote, and a circular portrait with a mono caption under it.
 * The one place on a page that spends a full band of green.
 */
function PullQuote({
  quote,
  name,
  role,
  portrait,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("section", _extends({
    style: {
      position: "relative",
      overflow: "hidden",
      background: "var(--color-ledger-900)",
      color: "#ffffff",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "absolute",
      inset: 0,
      pointerEvents: "none",
      background: "var(--glow-band)"
    },
    "aria-hidden": "true"
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative",
      maxWidth: "48rem",
      margin: "0 auto",
      padding: "8rem var(--gutter-sm)",
      textAlign: "center"
    }
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      fontFamily: "var(--font-serif)",
      fontSize: "var(--text-quote)",
      lineHeight: "var(--leading-quote)",
      letterSpacing: "-0.012em",
      textWrap: "balance",
      color: "inherit",
      textShadow: "var(--shadow-quote-glow)"
    }
  }, quote), name ? /*#__PURE__*/React.createElement("figure", {
    style: {
      margin: "2.5rem 0 0",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      gap: "0.75rem"
    }
  }, portrait ? /*#__PURE__*/React.createElement("div", {
    style: {
      borderRadius: "var(--radius-pill)",
      background: "rgba(255,255,255,0.1)",
      padding: "0.25rem",
      boxShadow: "0 10px 30px rgba(0,0,0,0.25)"
    }
  }, /*#__PURE__*/React.createElement("img", {
    src: portrait,
    width: "104",
    height: "104",
    alt: `${name}, ${role}`,
    style: {
      display: "block",
      width: "6.5rem",
      height: "6.5rem",
      borderRadius: "var(--radius-pill)",
      objectFit: "cover"
    }
  })) : null, /*#__PURE__*/React.createElement("figcaption", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-meta)",
      lineHeight: 1.7,
      letterSpacing: "var(--tracking-label)",
      textTransform: "uppercase",
      color: "rgba(255,255,255,0.6)"
    }
  }, name, /*#__PURE__*/React.createElement("br", null), /*#__PURE__*/React.createElement("span", {
    style: {
      color: "rgba(255,255,255,0.45)"
    }
  }, role))) : null));
}
Object.assign(__ds_scope, { PullQuote });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/content/PullQuote.jsx", error: String((e && e.message) || e) }); }

// components/core/Badge.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const tones = {
  /** The nav and footer "soon" marker on a page ground. */
  outline: {
    border: "1px solid var(--border-firm)",
    color: "var(--muted)",
    background: "transparent"
  },
  /** Same pill on the fixed green ground. */
  reversed: {
    border: "1px solid rgba(255,255,255,0.2)",
    color: "rgba(255,255,255,0.55)",
    background: "transparent"
  },
  /** Selected state, e.g. the agent picker. */
  accent: {
    border: "1px solid var(--accent)",
    color: "var(--accent)",
    background: "var(--accent-tint)"
  }
};

/**
 * A mono capsule: "soon" beside a nav item, "MIT" / "npm · JSR" in the footer,
 * DATA / PLACEHOLDER tags inside a figure. Never coloured semantically —
 * there is no red badge in this system.
 */
function Badge({
  tone = "outline",
  shape = "pill",
  children,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("span", _extends({
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: "0.25rem",
      borderRadius: shape === "pill" ? "var(--radius-pill)" : "var(--radius-sm)",
      padding: shape === "pill" ? "0.0625rem 0.375rem" : "0.125rem 0.375rem",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-badge)",
      lineHeight: 1.6,
      letterSpacing: "var(--tracking-pill)",
      textTransform: "uppercase",
      whiteSpace: "nowrap",
      ...(tones[tone] ?? tones.outline),
      ...style
    }
  }, rest), children);
}
Object.assign(__ds_scope, { Badge });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Badge.jsx", error: String((e && e.message) || e) }); }

// components/core/Button.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const base = {
  display: "inline-flex",
  alignItems: "center",
  gap: "0.5rem",
  borderRadius: "var(--radius)",
  padding: "0.625rem 1.25rem",
  fontFamily: "var(--font-sans)",
  fontSize: "var(--text-sm)",
  fontWeight: 500,
  lineHeight: 1.4,
  textDecoration: "none",
  cursor: "pointer",
  transition: "var(--transition-colors)"
};
const variants = {
  primary: {
    background: "var(--accent-solid)",
    color: "var(--on-accent)",
    border: "1px solid transparent",
    hover: {
      background: "var(--accent-deep)"
    }
  },
  secondary: {
    background: "transparent",
    color: "var(--accent)",
    border: "1px solid var(--border-firm)",
    hover: {
      background: "var(--accent-tint)"
    }
  },
  tertiary: {
    background: "var(--paper)",
    color: "var(--ink)",
    border: "1px solid var(--border-firm)",
    hover: {
      color: "var(--accent)",
      borderColor: "var(--accent)"
    }
  },
  quiet: {
    background: "transparent",
    color: "var(--muted)",
    border: "1px solid var(--border-firm)",
    hover: {
      color: "var(--accent)"
    }
  },
  link: {
    background: "transparent",
    color: "var(--accent)",
    border: "1px solid transparent",
    padding: 0,
    fontWeight: 500,
    hover: {
      textDecoration: "underline"
    }
  }
};

/**
 * The site's call to action. Primary is the filled green, secondary the
 * hairline-bordered accent, tertiary a paper button with ink text; every one
 * changes colour on hover and nothing else — no lift, no shadow, no scale.
 */
function Button({
  variant = "primary",
  href,
  children,
  disabled,
  style,
  ...rest
}) {
  const v = variants[variant] ?? variants.primary;
  const {
    hover,
    ...rest_v
  } = v;
  const [over, setOver] = React.useState(false);
  const Tag = href ? "a" : "button";
  return /*#__PURE__*/React.createElement(Tag, _extends({
    href: href,
    disabled: Tag === "button" ? disabled : undefined,
    onMouseEnter: () => setOver(true),
    onMouseLeave: () => setOver(false),
    style: {
      ...base,
      ...rest_v,
      ...(over && !disabled ? hover : null),
      ...(disabled ? {
        opacity: 0.4,
        cursor: "default"
      } : null),
      ...style
    }
  }, rest), children);
}
Object.assign(__ds_scope, { Button });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Button.jsx", error: String((e && e.message) || e) }); }

// components/core/Card.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The system's only box: a hairline border, a 6px radius, no shadow. `surface`
 * sits on paper, `paper` sits on a surface band. `interactive` adds the
 * hover the doc cards use — border and ground go accent, nothing moves.
 */
function Card({
  variant = "surface",
  interactive = false,
  as = "div",
  children,
  style,
  ...rest
}) {
  const [over, setOver] = React.useState(false);
  const Tag = as;
  return /*#__PURE__*/React.createElement(Tag, _extends({
    onMouseEnter: interactive ? () => setOver(true) : undefined,
    onMouseLeave: interactive ? () => setOver(false) : undefined,
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0,
      border: "1px solid var(--border)",
      borderRadius: "var(--radius-md)",
      background: variant === "paper" ? "var(--paper)" : "var(--surface)",
      padding: "var(--pad-card)",
      textDecoration: "none",
      transition: "var(--transition-colors)",
      ...(over ? {
        borderColor: "var(--accent)",
        background: "var(--accent-tint)"
      } : null),
      ...style
    }
  }, rest), children);
}
Object.assign(__ds_scope, { Card });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Card.jsx", error: String((e && e.message) || e) }); }

// components/content/DocCard.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * A documentation link card: mono index, serif title with a nudging arrow, one
 * line of blurb. Hover takes the border and ground to the accent and moves the
 * arrow 2px — the only motion on the page.
 */
function DocCard({
  index,
  title,
  blurb,
  href,
  style,
  ...rest
}) {
  const [over, setOver] = React.useState(false);
  return /*#__PURE__*/React.createElement(__ds_scope.Card, _extends({
    as: "a",
    href: href,
    interactive: true,
    onMouseEnter: () => setOver(true),
    onMouseLeave: () => setOver(false),
    style: style
  }, rest), index != null ? /*#__PURE__*/React.createElement("span", {
    style: {
      marginBottom: "0.875rem",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      letterSpacing: "var(--tracking-label)",
      color: "var(--muted)",
      fontVariantNumeric: "tabular-nums"
    }
  }, String(index).padStart(2, "0")) : null, /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "var(--font-serif)",
      fontSize: "var(--text-h3-card)",
      fontWeight: 600,
      color: over ? "var(--accent)" : "var(--ink)",
      transition: "var(--transition-colors)"
    }
  }, title, /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true",
    style: {
      display: "inline-block",
      marginLeft: "0.25rem",
      transform: over ? "translateX(var(--nudge-x))" : "none",
      transition: "var(--transition-transform)"
    }
  }, "\u2192")), /*#__PURE__*/React.createElement("span", {
    style: {
      marginTop: "0.375rem",
      fontSize: "var(--text-sm)",
      lineHeight: 1.65,
      color: "var(--muted)"
    }
  }, blurb));
}
Object.assign(__ds_scope, { DocCard });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/content/DocCard.jsx", error: String((e && e.message) || e) }); }

// components/core/CopyField.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The install command, in a bordered mono field with a copy button. The button
 * takes its labels from props so a localised page carries no English in script.
 */
function CopyField({
  command = "npx docxcelerate init my-documents",
  prompt = "$",
  labelIdle = "Copy",
  labelCopied = "Copied",
  style,
  ...rest
}) {
  const [label, setLabel] = React.useState(labelIdle);
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(command);
      setLabel(labelCopied);
    } catch {
      setLabel("Press ⌘C");
    }
    setTimeout(() => setLabel(labelIdle), 1600);
  };
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: "0.75rem",
      border: "1px solid var(--border)",
      borderRadius: "var(--radius)",
      background: "var(--surface)",
      padding: "0.5rem 1rem",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-sm)",
      ...style
    }
  }, rest), prompt ? /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--muted)",
      userSelect: "none"
    }
  }, prompt) : null, /*#__PURE__*/React.createElement("code", {
    style: {
      color: "var(--ink)"
    }
  }, command), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: copy,
    style: {
      marginLeft: "0.25rem",
      border: 0,
      background: "transparent",
      borderRadius: "var(--radius-sm)",
      padding: "0.125rem 0.375rem",
      fontFamily: "var(--font-mono)",
      fontSize: "0.75rem",
      color: label === labelIdle ? "var(--muted)" : "var(--accent)",
      cursor: "pointer",
      transition: "var(--transition-colors)"
    }
  }, label));
}
Object.assign(__ds_scope, { CopyField });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/CopyField.jsx", error: String((e && e.message) || e) }); }

// components/core/Eyebrow.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The mono, uppercase, wide-tracked label that opens every section — and, with
 * `index`, the numbered variant the mid-page points use ("01 · Authoring").
 */
function Eyebrow({
  index,
  children,
  tone = "accent",
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("span", _extends({
    style: {
      display: "block",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      fontWeight: 500,
      letterSpacing: "var(--tracking-eyebrow)",
      textTransform: "uppercase",
      color: tone === "muted" ? "var(--muted)" : "var(--accent)",
      fontVariantNumeric: "tabular-nums",
      ...style
    }
  }, rest), index != null ? /*#__PURE__*/React.createElement(React.Fragment, null, String(index).padStart(2, "0"), /*#__PURE__*/React.createElement("span", {
    style: {
      margin: "0 0.5rem",
      color: "var(--muted)"
    },
    "aria-hidden": "true"
  }, "\xB7")) : null, children);
}
Object.assign(__ds_scope, { Eyebrow });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/Eyebrow.jsx", error: String((e && e.message) || e) }); }

// components/core/PillTab.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * A pressed-state pill row — the homepage's agent picker. It chooses which
 * instruction is showing, so the buttons carry aria-pressed rather than
 * behaving like tabs over panels.
 */
function PillTab({
  items = [],
  value,
  onChange,
  style,
  ...rest
}) {
  const [internal, setInternal] = React.useState(value ?? items[0]?.id);
  const active = value ?? internal;
  return /*#__PURE__*/React.createElement("div", _extends({
    role: "group",
    style: {
      display: "flex",
      flexWrap: "wrap",
      gap: "0.5rem",
      ...style
    }
  }, rest), items.map(item => {
    const on = item.id === active;
    return /*#__PURE__*/React.createElement("button", {
      key: item.id,
      type: "button",
      "aria-pressed": on,
      onClick: () => {
        setInternal(item.id);
        onChange?.(item.id);
      },
      style: {
        display: "inline-flex",
        alignItems: "center",
        gap: "0.5rem",
        borderRadius: "var(--radius-pill)",
        border: on ? "1px solid var(--accent)" : "1px solid var(--border)",
        background: on ? "var(--accent-tint)" : "transparent",
        color: on ? "var(--accent)" : "var(--muted)",
        padding: "0.375rem 0.875rem",
        fontFamily: "var(--font-sans)",
        fontSize: "var(--text-sm)",
        cursor: "pointer",
        transition: "var(--transition-colors)"
      }
    }, item.label);
  }));
}
Object.assign(__ds_scope, { PillTab });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/core/PillTab.jsx", error: String((e && e.message) || e) }); }

// components/navigation/DocsSidebar.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The docs left rail. Mono uppercase group headings, plain links beneath, and
 * the current page marked with the accent tint. A subgroup nests one level in
 * behind a hairline.
 */
function DocsSidebar({
  groups = [],
  currentSlug,
  style,
  ...rest
}) {
  const link = page => {
    const on = page.slug === currentSlug;
    return /*#__PURE__*/React.createElement("li", {
      key: page.slug,
      style: {
        margin: 0
      }
    }, /*#__PURE__*/React.createElement("a", {
      href: `/docs/${page.slug}/`,
      "aria-current": on ? "page" : undefined,
      style: {
        display: "block",
        borderRadius: "var(--radius)",
        padding: "0.375rem 0.75rem",
        textDecoration: "none",
        transition: "var(--transition-colors)",
        background: on ? "var(--accent-tint)" : "transparent",
        color: on ? "var(--accent)" : "var(--muted)",
        fontWeight: on ? 600 : 400
      }
    }, page.title));
  };
  return /*#__PURE__*/React.createElement("nav", _extends({
    "aria-label": "Documentation",
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1.75rem",
      fontSize: "var(--text-sm)",
      ...style
    }
  }, rest), groups.map(group => /*#__PURE__*/React.createElement("div", {
    key: group.label,
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.5rem"
    }
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      fontWeight: 500,
      letterSpacing: "var(--tracking-label)",
      textTransform: "uppercase",
      color: "var(--muted)"
    }
  }, group.label), group.pages?.length ? /*#__PURE__*/React.createElement("ul", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.125rem",
      listStyle: "none",
      margin: "0 0 0 -0.75rem",
      padding: 0
    }
  }, group.pages.map(link)) : null, (group.subgroups ?? []).map(sub => /*#__PURE__*/React.createElement("div", {
    key: sub.label,
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.125rem"
    }
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      margin: "0.25rem 0 0",
      fontSize: "0.7rem",
      fontFamily: "var(--font-sans)",
      fontWeight: 500,
      letterSpacing: "var(--tracking-pill)",
      textTransform: "uppercase",
      color: "var(--muted)",
      opacity: 0.7
    }
  }, sub.label), /*#__PURE__*/React.createElement("ul", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.125rem",
      listStyle: "none",
      margin: 0,
      padding: "0 0 0 0.5rem",
      borderLeft: "1px solid var(--border)"
    }
  }, sub.pages.map(link)))))));
}
Object.assign(__ds_scope, { DocsSidebar });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/navigation/DocsSidebar.jsx", error: String((e && e.message) || e) }); }

// components/navigation/PageToc.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The docs right rail: "On this page", h2s flush and h3s indented, the active
 * heading in the accent.
 */
function PageToc({
  headings = [],
  activeSlug,
  label = "On this page",
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("nav", _extends({
    "aria-label": label,
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.625rem",
      fontSize: "var(--text-sm)",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      fontWeight: 500,
      letterSpacing: "var(--tracking-label)",
      textTransform: "uppercase",
      color: "var(--muted)"
    }
  }, label), /*#__PURE__*/React.createElement("ul", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.375rem",
      listStyle: "none",
      margin: 0,
      padding: 0
    }
  }, headings.map(heading => /*#__PURE__*/React.createElement("li", {
    key: heading.slug,
    style: {
      margin: 0,
      paddingLeft: heading.depth > 2 ? "0.75rem" : 0
    }
  }, /*#__PURE__*/React.createElement("a", {
    href: `#${heading.slug}`,
    style: {
      textDecoration: "none",
      lineHeight: 1.5,
      color: heading.slug === activeSlug ? "var(--accent)" : "var(--muted)",
      fontWeight: heading.slug === activeSlug ? 600 : 400
    }
  }, heading.text)))));
}
Object.assign(__ds_scope, { PageToc });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/navigation/PageToc.jsx", error: String((e && e.message) || e) }); }

// components/navigation/SiteFooter.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const DEFAULT_COLUMNS = [{
  heading: "Docs",
  links: [{
    label: "Start here",
    href: "/docs/start-here/"
  }, {
    label: "The node model",
    href: "/docs/nodes/overview/"
  }, {
    label: "CLI commands",
    href: "/docs/cli/commands/"
  }, {
    label: "Package entrypoints",
    href: "/docs/reference/entrypoints/"
  }]
}, {
  heading: "Project",
  links: [{
    label: "GitHub",
    href: "https://github.com/LiamDotPro/Docxcelerate",
    external: true
  }, {
    label: "npm",
    href: "https://www.npmjs.com/package/docxcelerate",
    external: true
  }, {
    label: "JSR",
    href: "https://jsr.io/@docxcelerate/docxcelerate",
    external: true
  }, {
    label: "Releases",
    href: "https://github.com/LiamDotPro/Docxcelerate/releases",
    external: true
  }]
}, {
  heading: "Engine",
  links: [{
    label: "The engine",
    href: "/docs/generation/endpoint/"
  }, {
    label: "Build artifacts",
    href: "/docs/projects/artifacts/"
  }, {
    label: "Cloud",
    href: "#",
    external: true,
    soon: true
  }]
}];

/**
 * The footer, on fixed brand green in both themes: the page ends on the brand
 * rather than fading back into the ground it started on.
 */
function SiteFooter({
  columns = DEFAULT_COLUMNS,
  blurb = "Documents as components, DOCX as output. Built in the open, MIT licensed, and running on your own machine unless you decide otherwise.",
  version = "v0.4.1",
  checksum = "sha512-Qy8Xk3PmM1vB7fCq…",
  year = new Date().getFullYear(),
  style,
  ...rest
}) {
  const linkStyle = {
    display: "inline-flex",
    alignItems: "center",
    gap: "0.375rem",
    color: "rgba(255,255,255,0.8)",
    textDecoration: "none",
    transition: "var(--transition-colors)"
  };
  return /*#__PURE__*/React.createElement("footer", _extends({
    style: {
      background: "var(--color-ledger-900)",
      color: "#ffffff",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-page)",
      margin: "0 auto",
      padding: "5rem var(--gutter-sm) 2.25rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "1.6fr 1fr 1fr 1fr",
      gap: "2.5rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "24rem"
    }
  }, /*#__PURE__*/React.createElement(__ds_scope.Wordmark, {
    tone: "reversed",
    size: "1.5rem",
    weight: 700
  }), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "0.875rem 0 0",
      fontSize: "0.95rem",
      lineHeight: 1.65,
      color: "rgba(255,255,255,0.6)",
      textWrap: "pretty"
    }
  }, blurb), /*#__PURE__*/React.createElement("p", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      gap: "0.5rem",
      margin: "1.25rem 0 0"
    }
  }, /*#__PURE__*/React.createElement(__ds_scope.Badge, {
    tone: "reversed",
    style: {
      padding: "0.25rem 0.625rem",
      fontSize: "var(--text-meta)"
    }
  }, version), /*#__PURE__*/React.createElement(__ds_scope.Badge, {
    tone: "reversed",
    style: {
      padding: "0.25rem 0.625rem",
      fontSize: "var(--text-meta)"
    }
  }, "MIT"), /*#__PURE__*/React.createElement(__ds_scope.Badge, {
    tone: "reversed",
    style: {
      padding: "0.25rem 0.625rem",
      fontSize: "var(--text-meta)"
    }
  }, "npm \xB7 JSR"))), columns.map(column => /*#__PURE__*/React.createElement("nav", {
    key: column.heading,
    "aria-label": column.heading
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "0 0 0.875rem",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-meta)",
      letterSpacing: "var(--tracking-label)",
      textTransform: "uppercase",
      color: "rgba(255,255,255,0.45)"
    }
  }, column.heading), /*#__PURE__*/React.createElement("ul", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "0.625rem",
      listStyle: "none",
      margin: 0,
      padding: 0,
      fontSize: "0.9rem"
    }
  }, column.links.map(link => /*#__PURE__*/React.createElement("li", {
    key: link.label,
    style: {
      margin: 0
    }
  }, /*#__PURE__*/React.createElement("a", {
    href: link.href,
    style: linkStyle
  }, link.label, link.external ? /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\u2197") : null, link.soon ? /*#__PURE__*/React.createElement(__ds_scope.Badge, {
    tone: "reversed",
    style: {
      fontSize: "0.55rem"
    }
  }, "soon") : null))))))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      alignItems: "center",
      justifyContent: "space-between",
      gap: "0.5rem 1.5rem",
      marginTop: "3.5rem",
      borderTop: "1px solid rgba(255,255,255,0.15)",
      paddingTop: "1.25rem",
      fontFamily: "var(--font-mono)",
      fontSize: "0.68rem",
      letterSpacing: "0.04em",
      color: "rgba(255,255,255,0.45)"
    }
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0
    }
  }, "\xA9 ", year, " Docxcelerate contributors."), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0
    }
  }, checksum))));
}
Object.assign(__ds_scope, { SiteFooter });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/navigation/SiteFooter.jsx", error: String((e && e.message) || e) }); }

// components/navigation/SiteHeader.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
const GITHUB_PATH = "M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27s1.36.09 2 .27c1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.01 8.01 0 0 0 16 8c0-4.42-3.58-8-8-8Z";
const DEFAULT_LINKS = [{
  label: "Docs",
  href: "/docs/start-here/",
  current: true
}, {
  label: "Cloud",
  href: "#",
  external: true,
  soon: true
}, {
  label: "GitHub",
  href: "https://github.com/LiamDotPro/Docxcelerate",
  external: true,
  icon: "github"
}];

/**
 * The sticky site header: lockup left, links right, then the language control
 * and the theme toggle. Translucent paper with a small backdrop blur, one
 * hairline underneath, 3.5rem tall.
 */
function SiteHeader({
  links = DEFAULT_LINKS,
  locale = "EN",
  onToggleTheme,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("header", _extends({
    style: {
      position: "sticky",
      top: 0,
      zIndex: 40,
      borderBottom: "1px solid var(--border)",
      background: "color-mix(in srgb, var(--paper) 85%, transparent)",
      backdropFilter: "blur(4px)",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: "1.5rem",
      height: "var(--header-height)",
      maxWidth: "var(--container-page)",
      margin: "0 auto",
      padding: "0 var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement(__ds_scope.Wordmark, {
    as: "a",
    href: "/",
    size: "1.125rem"
  }), /*#__PURE__*/React.createElement("nav", {
    "aria-label": "Main",
    style: {
      marginLeft: "auto",
      display: "flex",
      alignItems: "center",
      gap: "1.25rem",
      fontSize: "var(--text-sm)"
    }
  }, links.map(link => /*#__PURE__*/React.createElement("a", {
    key: link.label,
    href: link.href,
    "aria-current": link.current ? "page" : undefined,
    title: link.icon ? link.label : undefined,
    "aria-label": link.icon ? link.label : undefined,
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: "0.375rem",
      textDecoration: "none",
      color: link.current ? "var(--accent)" : "var(--muted)",
      fontWeight: link.current ? 600 : 400,
      transition: "var(--transition-colors)"
    }
  }, link.icon === "github" ? /*#__PURE__*/React.createElement("svg", {
    viewBox: "0 0 16 16",
    width: "16",
    height: "16",
    fill: "currentColor",
    "aria-hidden": "true"
  }, /*#__PURE__*/React.createElement("path", {
    d: GITHUB_PATH
  })) : /*#__PURE__*/React.createElement(React.Fragment, null, link.label, link.external ? /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\u2197") : null), link.soon ? /*#__PURE__*/React.createElement(__ds_scope.Badge, null, "soon") : null)), /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      letterSpacing: "var(--tracking-pill)",
      color: "var(--muted)"
    }
  }, locale), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: onToggleTheme,
    "aria-label": "Switch theme",
    style: {
      display: "grid",
      placeItems: "center",
      width: "2rem",
      height: "2rem",
      borderRadius: "var(--radius)",
      border: "1px solid var(--border)",
      background: "transparent",
      color: "var(--muted)",
      cursor: "pointer",
      transition: "var(--transition-colors)"
    }
  }, /*#__PURE__*/React.createElement("svg", {
    viewBox: "0 0 24 24",
    width: "16",
    height: "16",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "1.6",
    "aria-hidden": "true"
  }, /*#__PURE__*/React.createElement("path", {
    d: "M12 3v1.5M12 19.5V21M3 12h1.5M19.5 12H21M5.6 5.6l1.1 1.1M17.3 17.3l1.1 1.1M18.4 5.6l-1.1 1.1M6.7 17.3l-1.1 1.1",
    strokeLinecap: "round"
  }), /*#__PURE__*/React.createElement("circle", {
    cx: "12",
    cy: "12",
    r: "3.8"
  }))))));
}
Object.assign(__ds_scope, { SiteHeader });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/navigation/SiteHeader.jsx", error: String((e && e.message) || e) }); }

// components/product/FileTree.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The workspace file tree from the preview app: mono rows, indentation encoding
 * real directory depth, documents expandable with a caret and their sector set
 * to the right. The selected file takes the accent tint and a 2px accent bar on
 * its left edge.
 */
function FileTree({
  root = "my-documents/",
  nodes = [],
  selected,
  expanded = [],
  onSelect,
  onToggle,
  style,
  ...rest
}) {
  const rowBase = {
    display: "block",
    width: "100%",
    textAlign: "left",
    fontFamily: "var(--font-mono)",
    fontSize: "var(--text-tree)",
    lineHeight: 1.6,
    whiteSpace: "nowrap",
    paddingTop: "0.15rem",
    paddingBottom: "0.15rem",
    paddingRight: "0.6rem",
    border: 0,
    borderLeft: "2px solid transparent",
    background: "transparent",
    cursor: "pointer",
    transition: "var(--transition-colors)"
  };
  const pad = depth => `calc(0.6rem + ${depth} * 0.7rem)`;
  const render = node => {
    if (node.kind === "dir") {
      return /*#__PURE__*/React.createElement("span", {
        key: node.label + node.depth,
        style: {
          ...rowBase,
          cursor: "default",
          color: "var(--muted)",
          opacity: 0.75,
          paddingLeft: pad(node.depth)
        }
      }, node.label);
    }
    if (node.kind === "document") {
      const open = expanded.includes(node.id);
      return /*#__PURE__*/React.createElement(React.Fragment, {
        key: node.id
      }, /*#__PURE__*/React.createElement("button", {
        type: "button",
        "aria-expanded": open,
        onClick: () => onToggle?.(node.id),
        style: {
          ...rowBase,
          display: "flex",
          alignItems: "baseline",
          gap: "0.4rem",
          paddingLeft: pad(node.depth),
          color: open ? "var(--accent)" : "var(--ink)",
          fontWeight: open ? 600 : 400
        }
      }, /*#__PURE__*/React.createElement("span", {
        "aria-hidden": "true",
        style: {
          flex: "none",
          fontSize: "0.6rem",
          transform: open ? "rotate(90deg)" : "none",
          transition: "var(--transition-transform)"
        }
      }, "\u25B8"), /*#__PURE__*/React.createElement("span", {
        style: {
          overflow: "hidden",
          textOverflow: "ellipsis"
        }
      }, node.label), node.sector ? /*#__PURE__*/React.createElement("span", {
        style: {
          marginLeft: "auto",
          flex: "none",
          fontSize: "0.6rem",
          letterSpacing: "var(--tracking-pill)",
          textTransform: "uppercase",
          color: "var(--muted)",
          opacity: 0.75,
          fontWeight: 400
        }
      }, node.sector) : null), open ? (node.children ?? []).map(render) : null);
    }
    const on = node.id === selected;
    return /*#__PURE__*/React.createElement("button", {
      key: node.id,
      type: "button",
      role: "tab",
      "aria-selected": on,
      onClick: () => onSelect?.(node.id),
      style: {
        ...rowBase,
        paddingLeft: pad(node.depth),
        color: on ? "var(--accent)" : "var(--muted)",
        background: on ? "var(--accent-tint)" : "transparent",
        borderLeftColor: on ? "var(--accent-solid)" : "transparent",
        fontWeight: on ? 600 : 400
      }
    }, node.label);
  };
  return /*#__PURE__*/React.createElement("nav", _extends({
    role: "tablist",
    "aria-orientation": "vertical",
    "aria-label": "Workspace files",
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1px",
      padding: "0.5rem 0",
      overflow: "auto",
      background: "var(--surface)",
      borderRight: "1px solid var(--border)",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-tree)",
      lineHeight: 1.6,
      color: "var(--muted)",
      opacity: 0.75,
      padding: "0.15rem 0.6rem"
    }
  }, root), nodes.map(render));
}
Object.assign(__ds_scope, { FileTree });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/product/FileTree.jsx", error: String((e && e.message) || e) }); }

// components/product/PanelBar.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The mono status bar that tops each half of the preview app: a path on the
 * left, controls or a count on the right. Uppercase, 0.65rem, always muted.
 */
function PanelBar({
  path,
  hint,
  children,
  style,
  ...rest
}) {
  const label = {
    fontFamily: "var(--font-mono)",
    fontSize: "var(--text-eyebrow)",
    letterSpacing: "0.12em",
    textTransform: "uppercase",
    color: "var(--muted)",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis"
  };
  return /*#__PURE__*/React.createElement("header", _extends({
    style: {
      display: "flex",
      flex: "none",
      alignItems: "center",
      justifyContent: "space-between",
      gap: "0.75rem",
      height: "2.5rem",
      padding: "0 0.75rem",
      background: "var(--surface)",
      borderBottom: "1px solid var(--border)",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("span", {
    style: label
  }, path), children ?? (hint ? /*#__PURE__*/React.createElement("span", {
    style: label
  }, hint) : null));
}
Object.assign(__ds_scope, { PanelBar });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/product/PanelBar.jsx", error: String((e && e.message) || e) }); }

// components/product/PaperPage.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * An A4 sheet at 96dpi — 210mm × 297mm, 25.4mm margins, always white with
 * #1f2933 ink whatever the theme is doing. A letter preview that dims at night
 * stops reading as paper, and the product's claim is that the preview is the
 * document.
 */
function PaperPage({
  title,
  children,
  scale = 1,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("article", _extends({
    "aria-label": title,
    style: {
      width: "210mm",
      minHeight: "297mm",
      padding: "25.4mm",
      background: "var(--paper-fixed)",
      color: "var(--paper-fixed-ink)",
      border: "1px solid var(--paper-fixed-border)",
      boxShadow: "var(--shadow-page-docx)",
      transform: scale === 1 ? undefined : `scale(${scale})`,
      transformOrigin: "top left",
      fontFamily: "var(--font-letter)",
      ...style
    }
  }, rest), title ? /*#__PURE__*/React.createElement("h1", {
    style: {
      margin: "0 0 18pt",
      fontFamily: "Cambria, Georgia, serif",
      fontSize: "20pt",
      lineHeight: 1.2,
      fontWeight: 700,
      color: "var(--docx-text)"
    }
  }, title) : null, children);
}

/** A section inside a rendered document: 12pt heading, 12pt of space after. */
function PaperSection({
  heading,
  children,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("section", _extends({
    style: {
      margin: "0 0 12pt",
      ...style
    }
  }, rest), heading ? /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: "16pt 0 7pt",
      fontSize: "12pt",
      lineHeight: 1.3,
      fontWeight: 700,
      color: "var(--docx-text)"
    }
  }, heading) : null, children);
}

/** A paragraph node. `dynamic` prints the node's id above it, as the renderer does. */
function PaperParagraph({
  id,
  dynamic = false,
  children,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("p", _extends({
    style: {
      margin: "0 0 10pt",
      fontFamily: 'Calibri, "Aptos", "Segoe UI", Arial, sans-serif',
      fontSize: "11pt",
      lineHeight: 1.42,
      color: "var(--docx-text)",
      ...style
    }
  }, rest), dynamic && id ? /*#__PURE__*/React.createElement("span", {
    style: {
      display: "block",
      marginBottom: "3pt",
      color: "var(--docx-text-muted)",
      fontSize: "8pt",
      lineHeight: 1.2,
      textTransform: "uppercase"
    }
  }, "Dynamic: ", id) : null, children);
}

/** The dashed stand-in the renderers emit for an image, graph or TOC node. */
function PaperPlaceholder({
  children,
  style,
  ...rest
}) {
  return /*#__PURE__*/React.createElement("div", _extends({
    style: {
      margin: "10pt 0",
      padding: "9pt",
      border: "1px dashed #9aa6b8",
      color: "#59677a",
      fontSize: "9pt",
      ...style
    }
  }, rest), children);
}
Object.assign(__ds_scope, { PaperPage, PaperSection, PaperParagraph, PaperPlaceholder });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/product/PaperPage.jsx", error: String((e && e.message) || e) }); }

// components/product/WordShell.jsx
try { (() => {
function _extends() { return _extends = Object.assign ? Object.assign.bind() : function (n) { for (var e = 1; e < arguments.length; e++) { var t = arguments[e]; for (var r in t) ({}).hasOwnProperty.call(t, r) && (n[r] = t[r]); } return n; }, _extends.apply(null, arguments); }
/**
 * The Word-shaped chrome the web renderer wraps a document in: blue titlebar,
 * a ribbon of inert tabs and tool blocks, then a ruled grey workspace holding
 * the page. Cosmetic by design — nothing in the ribbon does anything.
 */
function WordShell({
  title = "Document",
  live = false,
  pageWidth = "210mm",
  children,
  style,
  ...rest
}) {
  const tabs = ["Home", "Insert", "Layout", "Review", "View"];
  const tool = {
    height: "26px",
    border: "1px solid #d4d9e2",
    borderRadius: "4px",
    background: "linear-gradient(#ffffff, #f0f2f6)"
  };
  return /*#__PURE__*/React.createElement("main", _extends({
    style: {
      display: "grid",
      gridTemplateRows: "auto auto 1fr",
      height: "100%",
      background: "var(--docx-workspace)",
      color: "#1f2933",
      fontFamily: "var(--font-letter)",
      ...style
    }
  }, rest), /*#__PURE__*/React.createElement("header", {
    style: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      minHeight: "44px",
      padding: "0 18px",
      background: "var(--docx-titlebar)",
      color: "#ffffff"
    }
  }, /*#__PURE__*/React.createElement("strong", {
    style: {
      fontSize: "14px",
      fontWeight: 600
    }
  }, title), /*#__PURE__*/React.createElement("span", {
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: "8px",
      fontSize: "12px",
      opacity: 0.94
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      width: "8px",
      height: "8px",
      borderRadius: "50%",
      background: live ? "#44c767" : "#a9bad8"
    }
  }), live ? "Live preview connected" : "Preview")), /*#__PURE__*/React.createElement("section", {
    "aria-label": "Document toolbar",
    style: {
      background: "var(--docx-chrome)",
      borderBottom: "1px solid var(--docx-chrome-border)",
      boxShadow: "0 1px 0 rgb(15 23 42 / 4%)"
    }
  }, /*#__PURE__*/React.createElement("nav", {
    "aria-label": "Ribbon tabs",
    style: {
      display: "flex",
      gap: "2px",
      minHeight: "34px",
      padding: "0 16px",
      borderBottom: "1px solid var(--docx-chrome-border)"
    }
  }, tabs.map((tab, i) => /*#__PURE__*/React.createElement("span", {
    key: tab,
    style: {
      display: "inline-flex",
      alignItems: "center",
      padding: "0 14px",
      fontSize: "13px",
      color: i === 0 ? "var(--docx-accent)" : "#384252",
      fontWeight: i === 0 ? 600 : 400,
      borderBottom: i === 0 ? "2px solid var(--docx-accent)" : "2px solid transparent"
    }
  }, tab))), /*#__PURE__*/React.createElement("div", {
    "aria-hidden": "true",
    style: {
      display: "flex",
      alignItems: "center",
      gap: "22px",
      minHeight: "74px",
      padding: "10px 18px 12px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      ...tool,
      width: "116px",
      background: "#ffffff"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(3, 30px)",
      gap: "6px",
      paddingRight: "20px",
      borderRight: "1px solid var(--docx-chrome-border)"
    }
  }, Array.from({
    length: 6
  }).map((_, i) => /*#__PURE__*/React.createElement("span", {
    key: i,
    style: tool
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      ...tool,
      width: "116px",
      background: "#ffffff"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      ...tool,
      width: "116px",
      background: "#ffffff"
    }
  }))), /*#__PURE__*/React.createElement("section", {
    "aria-label": "Document workspace",
    style: {
      overflow: "auto",
      background: "linear-gradient(90deg, rgb(0 0 0 / 4%) 0 1px, transparent 1px) 0 0 / 24px 24px, var(--docx-workspace)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "0 clamp(16px, 6%, 96px) 80px"
    }
  }, /*#__PURE__*/React.createElement("div", {
    "aria-hidden": "true",
    style: {
      position: "sticky",
      top: 0,
      zIndex: 1,
      width: pageWidth,
      height: "28px",
      margin: "0 auto 18px",
      background: "repeating-linear-gradient(90deg, #f8f9fb 0, #f8f9fb 9.5mm, #c7ccd6 9.5mm, #c7ccd6 10mm)",
      border: "1px solid #cfd4dd",
      borderTop: 0,
      boxShadow: "0 1px 2px rgb(15 23 42 / 10%)"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      placeItems: "start center",
      minWidth: "100%"
    }
  }, children))));
}
Object.assign(__ds_scope, { WordShell });
})(); } catch (e) { __ds_ns.__errors.push({ path: "components/product/WordShell.jsx", error: String((e && e.message) || e) }); }

// ui_kits/preview-app/PreviewApp.jsx
try { (() => {
const {
  PanelBar,
  FileTree,
  WordShell,
  PaperPage,
  PaperSection,
  PaperParagraph,
  PaperPlaceholder,
  CodeFigure
} = window.DocxcelerateDesignSystem_cb7f53;
const t = (text, kind) => ({
  text,
  tone: kind
});
const DOCUMENTS = {
  "offer-of-admission": {
    name: "Offer of Admission",
    sector: "Education",
    files: ["greeting.node.tsx", "tutor-note.node.tsx", "conditions.node.tsx", "document.tsx", "types.ts"],
    page: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Your offer"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "Dear Maya, we are delighted to offer you a place on the BSc Architecture programme beginning in September."), /*#__PURE__*/React.createElement(PaperParagraph, {
      id: "tutor-note",
      dynamic: true
    }, "A short personal note from Dr Priya Raman about Maya Oyelaran's interview.")), /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Conditions"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "This offer is conditional on the grades listed in your application being confirmed by the awarding body."), /*#__PURE__*/React.createElement(PaperParagraph, null, "Please confirm your acceptance within 30 days of this notice."), /*#__PURE__*/React.createElement(PaperPlaceholder, null, "Image: campus map")), /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Next steps"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "Your tutor will contact you before the end of the month to arrange enrolment.")))
  },
  "repairs-appointment": {
    name: "Repair Appointment",
    sector: "Housing",
    files: ["greeting.node.tsx", "appointment.node.tsx", "access.node.tsx", "what-to-expect.node.tsx", "document.tsx"],
    page: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Your appointment"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "Dear Mr Achterberg, an engineer will visit Flat 12, Sanderling Court on Tuesday 4 March between 09:00 and 13:00.")), /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Access"
    }, /*#__PURE__*/React.createElement(PaperParagraph, {
      id: "access",
      dynamic: true
    }, "Whether someone over 18 needs to be at home, given the work booked.")), /*#__PURE__*/React.createElement(PaperSection, {
      heading: "What to expect"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "The engineer will isolate the supply, replace the failed valve and test the system before leaving."), /*#__PURE__*/React.createElement(PaperPlaceholder, null, "Table of contents")))
  },
  "policy-renewal": {
    name: "Policy Renewal",
    sector: "Insurance",
    files: ["greeting.node.tsx", "renewal.node.tsx", "premium-change.node.tsx", "shopping.node.tsx", "document.tsx"],
    page: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Your renewal"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "Dear Ms Ferreira, your home insurance policy HB-4471902 renews on 12 April.")), /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Premium change"
    }, /*#__PURE__*/React.createElement(PaperParagraph, {
      id: "premium-change",
      dynamic: true
    }, "An explanation of the change from \xA341.20 to \xA346.85 a month."), /*#__PURE__*/React.createElement(PaperPlaceholder, null, "Bar graph: premium over four renewals")), /*#__PURE__*/React.createElement(PaperSection, {
      heading: "Shopping around"
    }, /*#__PURE__*/React.createElement(PaperParagraph, null, "You do not have to renew with us. If you do nothing, the policy renews automatically on the date above.")))
  }
};
const SOURCE = {
  "document.tsx": [{
    spans: [t("import", "accent"), t(" { Document, Section, template } ", "ink"), t("from", "accent"), t(" ", "ink"), t('"docxcelerate/template"', "string"), t(";", "muted")]
  }, {
    spans: [t("", "muted")]
  }, {
    spans: [t("export const", "accent"), t(" documentTemplate = ", "ink"), t("template", "accent"), t("<Data>(", "muted")]
  }, {
    indent: 1,
    spans: [t("<", "muted"), t("Document", "accent"), t(" id=", "ink"), t('"offer-of-admission"', "string"), t(" title=", "ink"), t('"Offer of Admission"', "string"), t(">", "muted")]
  }, {
    indent: 2,
    spans: [t("<", "muted"), t("Section", "accent"), t(" id=", "ink"), t('"offer"', "string"), t(" title=", "ink"), t('"Your offer"', "string"), t(">", "muted")]
  }, {
    indent: 3,
    spans: [t("<", "muted"), t("Greeting", "accent"), t(" />", "muted")]
  }, {
    indent: 3,
    spans: [t("<", "muted"), t("TutorNote", "accent"), t(" />", "muted")]
  }, {
    indent: 2,
    spans: [t("</", "muted"), t("Section", "accent"), t(">", "muted")]
  }, {
    indent: 1,
    spans: [t("</", "muted"), t("Document", "accent"), t(">", "muted")]
  }, {
    spans: [t(");", "muted")]
  }],
  node: [{
    spans: [t("export const", "accent"), t(" TutorNote: ", "ink"), t("Paragraph", "accent"), t(" = () => {", "muted")]
  }, {
    indent: 1,
    spans: [t("const", "accent"), t(" [state] = ", "ink"), t("useState", "accent"), t("((data: Data) => ({", "muted")]
  }, {
    indent: 2,
    spans: [t("interviewer: data.interviewer,", "ink")]
  }, {
    indent: 1,
    spans: [t("}));", "muted")]
  }, {
    spans: [t("", "muted")]
  }, {
    indent: 1,
    spans: [t("useSetPrompts", "accent"), t("({", "muted")]
  }, {
    indent: 2,
    spans: [t("generalPrompt: ", "muted"), t('"A short personal note."', "string"), t(",", "muted")]
  }, {
    indent: 2,
    spans: [t("tone: ", "muted"), t('"warm, never effusive"', "string"), t(",", "muted")]
  }, {
    indent: 1,
    spans: [t("});", "muted")]
  }, {
    spans: [t("", "muted")]
  }, {
    indent: 1,
    spans: [t("return", "accent"), t(" <", "muted"), t("Paragraph", "accent"), t(" id=", "ink"), t('"tutor-note"', "string"), t(" />;", "muted")]
  }, {
    spans: [t("};", "muted")]
  }],
  "types.ts": [{
    spans: [t("export interface", "accent"), t(" Data {", "muted")]
  }, {
    indent: 1,
    spans: [t("recipientName: ", "ink"), t("string", "accent"), t(";", "muted")]
  }, {
    indent: 1,
    spans: [t("interviewer: ", "ink"), t("string", "accent"), t(";", "muted")]
  }, {
    indent: 1,
    spans: [t("conditions: ", "ink"), t("string", "accent"), t("[];", "muted")]
  }, {
    spans: [t("}", "muted")]
  }],
  "docxcelerate.config.json": [{
    spans: [t("{", "muted")]
  }, {
    indent: 1,
    spans: [t('"documentsDir"', "ink"), t(": ", "muted"), t('"documents"', "string"), t(",", "muted")]
  }, {
    indent: 1,
    spans: [t('"outDir"', "ink"), t(": ", "muted"), t('"dist"', "string"), t(",", "muted")]
  }, {
    indent: 1,
    spans: [t('"engine"', "ink"), t(": { ", "muted"), t('"url"', "ink"), t(": ", "muted"), t('"http://localhost:8787"', "string"), t(" }", "muted")]
  }, {
    spans: [t("}", "muted")]
  }]
};
const sourceFor = file => SOURCE[file] ?? (file.endsWith(".node.tsx") ? SOURCE.node : SOURCE["document.tsx"]);
const treeFor = openId => {
  const nodes = [{
    kind: "file",
    id: "docxcelerate.config.json",
    label: "docxcelerate.config.json",
    depth: 1
  }, {
    kind: "dir",
    label: "documents/",
    depth: 1
  }];
  for (const [id, doc] of Object.entries(DOCUMENTS)) {
    const files = doc.files.map(label => ({
      kind: "file",
      id: `${id}/${label}`,
      label,
      depth: label.endsWith(".node.tsx") ? 4 : 3
    }));
    nodes.push({
      kind: "document",
      id,
      label: `${id}/`,
      depth: 2,
      sector: doc.sector,
      children: [{
        kind: "dir",
        label: "nodes/",
        depth: 3
      }, ...files]
    });
  }
  return nodes;
};
const zoomBtn = active => ({
  border: active ? "1px solid var(--accent-solid)" : "1px solid var(--border)",
  borderRadius: "3px",
  padding: "0.1rem 0.5rem",
  fontFamily: "var(--font-mono)",
  fontSize: "0.65rem",
  lineHeight: 1.5,
  letterSpacing: "0.1em",
  textTransform: "uppercase",
  color: active ? "var(--accent)" : "var(--muted)",
  background: "transparent",
  cursor: "pointer",
  transition: "var(--transition-colors)"
});

/** `dxcl dev` — the local preview app: workspace on the left, the document it produces on the right. */
function PreviewApp() {
  const [docId, setDocId] = React.useState("offer-of-admission");
  const [file, setFile] = React.useState("offer-of-admission/document.tsx");
  const [treeOpen, setTreeOpen] = React.useState(true);
  const [zoom, setZoom] = React.useState(0.5);
  const [fitting, setFitting] = React.useState(true);
  const doc = DOCUMENTS[docId];
  const fileName = file.split("/").pop();
  const select = id => {
    setFile(id);
    const owner = id.includes("/") ? id.split("/")[0] : null;
    if (owner && DOCUMENTS[owner]) setDocId(owner);
  };
  return /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "50% 0.25rem minmax(0,1fr)",
      height: "100vh",
      background: "var(--surface-2)"
    }
  }, /*#__PURE__*/React.createElement("section", {
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0,
      borderRight: "1px solid var(--border)"
    }
  }, /*#__PURE__*/React.createElement(PanelBar, {
    path: "my-documents/"
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: "0.625rem"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      letterSpacing: "0.12em",
      textTransform: "uppercase",
      color: "var(--muted)"
    }
  }, "3 documents \xB7 ", Object.values(DOCUMENTS).reduce((n, d) => n + d.files.length, 0) + 1, " files"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => setTreeOpen(!treeOpen),
    "aria-expanded": treeOpen,
    "aria-label": treeOpen ? "Hide workspace files" : "Show workspace files",
    style: {
      display: "grid",
      placeItems: "center",
      width: "1.5rem",
      height: "1.5rem",
      border: 0,
      borderRadius: "3px",
      background: treeOpen ? "var(--surface-2)" : "transparent",
      color: treeOpen ? "var(--accent)" : "var(--muted)",
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("svg", {
    viewBox: "0 0 16 16",
    width: "15",
    height: "15",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "1.3",
    "aria-hidden": "true"
  }, /*#__PURE__*/React.createElement("rect", {
    x: "1.5",
    y: "2.5",
    width: "13",
    height: "11",
    rx: "1.5"
  }), /*#__PURE__*/React.createElement("line", {
    x1: "6",
    y1: "2.5",
    x2: "6",
    y2: "13.5"
  }))))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: treeOpen ? "15rem minmax(0,1fr)" : "0 minmax(0,1fr)",
      flex: 1,
      minHeight: 0,
      transition: "grid-template-columns var(--duration-slow) var(--ease)"
    }
  }, /*#__PURE__*/React.createElement(FileTree, {
    nodes: treeFor(docId),
    selected: file,
    expanded: [docId],
    onSelect: select,
    onToggle: id => {
      setDocId(id);
      setFile(`${id}/document.tsx`);
    },
    style: {
      visibility: treeOpen ? "visible" : "hidden"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0,
      minHeight: 0,
      background: "var(--paper)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: "0.5rem",
      padding: "0.5rem 1.1rem 0",
      fontFamily: "var(--font-mono)",
      fontSize: "0.65rem",
      color: "var(--muted)",
      whiteSpace: "nowrap",
      overflow: "hidden"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--accent)"
    }
  }, fileName), /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\xB7"), /*#__PURE__*/React.createElement("span", {
    style: {
      overflow: "hidden",
      textOverflow: "ellipsis"
    }
  }, file)), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      minHeight: 0,
      overflow: "auto",
      padding: "0.75rem 1.1rem 1rem"
    }
  }, /*#__PURE__*/React.createElement(CodeFigure, {
    lines: sourceFor(fileName),
    size: "0.76rem"
  }))))), /*#__PURE__*/React.createElement("div", {
    role: "separator",
    "aria-orientation": "vertical",
    style: {
      background: "var(--border)",
      cursor: "col-resize"
    }
  }), /*#__PURE__*/React.createElement("section", {
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0
    }
  }, /*#__PURE__*/React.createElement(PanelBar, {
    path: `${doc.name} · A4`
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: "0.25rem"
    }
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => {
      setFitting(true);
      setZoom(0.5);
    },
    "aria-pressed": fitting,
    style: zoomBtn(fitting)
  }, "Fit"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => {
      setFitting(false);
      setZoom(Math.max(0.1, zoom / 1.25));
    },
    "aria-label": "Zoom out",
    style: {
      ...zoomBtn(false),
      minWidth: "1.6rem",
      fontSize: "0.8rem"
    }
  }, "\u2212"), /*#__PURE__*/React.createElement("output", {
    style: {
      minWidth: "3.1rem",
      textAlign: "center",
      fontFamily: "var(--font-mono)",
      fontSize: "0.65rem",
      fontVariantNumeric: "tabular-nums",
      color: "var(--muted)"
    }
  }, Math.round(zoom * 100), "%"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => {
      setFitting(false);
      setZoom(Math.min(3, zoom * 1.25));
    },
    "aria-label": "Zoom in",
    style: {
      ...zoomBtn(false),
      minWidth: "1.6rem",
      fontSize: "0.8rem"
    }
  }, "+"))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      minHeight: 0
    }
  }, /*#__PURE__*/React.createElement(WordShell, {
    title: doc.name,
    live: true,
    pageWidth: `calc(210mm * ${zoom})`
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: `calc(210mm * ${zoom})`,
      height: `calc(297mm * ${zoom})`
    }
  }, /*#__PURE__*/React.createElement(PaperPage, {
    title: doc.name,
    scale: zoom
  }, doc.page))))));
}
Object.assign(window, {
  PreviewApp,
  DOCUMENTS,
  treeFor,
  sourceFor
});
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/preview-app/PreviewApp.jsx", error: String((e && e.message) || e) }); }

// ui_kits/website/DocsScreen.jsx
try { (() => {
const {
  DocsSidebar,
  PageToc,
  Eyebrow,
  CodeFigure,
  Button
} = window.DocxcelerateDesignSystem_cb7f53;
const GROUPS = [{
  label: "Start Here",
  pages: [{
    slug: "start-here",
    title: "Start here"
  }, {
    slug: "writing-nodes",
    title: "Writing nodes"
  }, {
    slug: "document-projects",
    title: "Document projects"
  }]
}, {
  label: "Essentials",
  pages: [{
    slug: "essentials/documents-and-nodes",
    title: "Documents and nodes"
  }, {
    slug: "essentials/templates",
    title: "Templates"
  }]
}, {
  label: "Nodes",
  pages: [{
    slug: "nodes/overview",
    title: "Overview"
  }],
  subgroups: [{
    label: "Node types",
    pages: [{
      slug: "nodes/paragraph",
      title: "Paragraph"
    }, {
      slug: "nodes/section",
      title: "Section"
    }, {
      slug: "nodes/image",
      title: "Image"
    }, {
      slug: "nodes/graph",
      title: "Graph"
    }, {
      slug: "nodes/table-of-contents",
      title: "Table of contents"
    }]
  }]
}, {
  label: "CLI",
  pages: [{
    slug: "cli/commands",
    title: "CLI commands"
  }]
}, {
  label: "Projects",
  pages: [{
    slug: "projects/workspace",
    title: "Workspace"
  }, {
    slug: "projects/artifacts",
    title: "Build artifacts"
  }]
}, {
  label: "Generation",
  pages: [{
    slug: "generation/endpoint",
    title: "The engine"
  }]
}, {
  label: "Reference",
  pages: [{
    slug: "reference/entrypoints",
    title: "Package entrypoints"
  }]
}];
const HEADINGS = [{
  slug: "what-a-node-is",
  text: "What a node is",
  depth: 2
}, {
  slug: "state",
  text: "Reading data with useState",
  depth: 2
}, {
  slug: "prompts",
  text: "Handing work to a model",
  depth: 2
}, {
  slug: "set-prompts",
  text: "useSetPrompts",
  depth: 3
}, {
  slug: "testing",
  text: "Testing a node",
  depth: 2
}];
const t = (text, kind) => ({
  text,
  tone: kind
});
const NODE_SOURCE = [{
  spans: [t("import", "accent"), t(" { Paragraph, useSetPrompts, useState } ", "ink"), t("from", "accent"), t(" ", "ink"), t('"docxcelerate/template"', "string"), t(";", "muted")]
}, {
  spans: [t("", "muted")]
}, {
  spans: [t("export const", "accent"), t(" Balance: ", "ink"), t("Paragraph", "accent"), t(" = () => {", "muted")]
}, {
  indent: 1,
  spans: [t("const", "accent"), t(" [state] = ", "ink"), t("useState", "accent"), t("((data: TenancyData) => ({", "muted")]
}, {
  indent: 2,
  spans: [t("name: data.recipientName,", "ink")]
}, {
  indent: 2,
  spans: [t("settled: data.balanceDue === ", "ink"), t("0", "accent"), t(",", "muted")]
}, {
  indent: 1,
  spans: [t("}));", "muted")]
}, {
  spans: [t("", "muted")]
}, {
  indent: 1,
  spans: [t("useSetPrompts", "accent"), t("({ generalPrompt: ", "muted"), t("`Explain the balance to ${state.name}.`", "string"), t(" });", "muted")]
}, {
  spans: [t("", "muted")]
}, {
  indent: 1,
  spans: [t("if", "accent"), t(" (state.settled) {", "muted")]
}, {
  indent: 2,
  spans: [t("return", "accent"), t(" <", "muted"), t("Paragraph", "accent"), t(" id=", "ink"), t('"settled"', "string"), t(">Nothing outstanding.</", "muted"), t("Paragraph", "accent"), t(">;", "muted")]
}, {
  indent: 1,
  spans: [t("}", "muted")]
}, {
  spans: [t("", "muted")]
}, {
  indent: 1,
  spans: [t("return", "accent"), t(" <", "muted"), t("Paragraph", "accent"), t(" id=", "ink"), t('"arrears"', "string"), t(" />;", "muted")]
}, {
  spans: [t("};", "muted")]
}];
const prose = {
  maxWidth: "46rem",
  lineHeight: 1.7,
  minWidth: 0,
  flex: 1
};
const h2 = {
  margin: "2.75rem 0 0.75rem",
  fontSize: "1.6rem",
  fontWeight: 600
};
const p = {
  margin: "0 0 1.1rem",
  color: "var(--ink)"
};
const inlineCode = {
  fontSize: "0.86em",
  background: "var(--accent-tint)",
  color: "var(--accent)",
  padding: "0.1em 0.35em",
  borderRadius: "3px",
  fontFamily: "var(--font-mono)"
};
const th = {
  borderBottom: "1px solid var(--border)",
  padding: "0.55rem 0.85rem",
  textAlign: "left",
  fontFamily: "var(--font-mono)",
  fontSize: "0.68rem",
  letterSpacing: "0.1em",
  textTransform: "uppercase",
  color: "var(--muted)",
  fontWeight: 500,
  background: "var(--surface-2)"
};
const td = {
  borderBottom: "1px solid var(--border)",
  padding: "0.55rem 0.85rem",
  textAlign: "left",
  fontSize: "0.9rem"
};

/** A documentation page: two rails around a 46rem measure. */
function DocsScreen() {
  const [slug, setSlug] = React.useState("writing-nodes");
  return /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: "3rem",
      maxWidth: "var(--container-page)",
      width: "100%",
      margin: "0 auto",
      padding: "2.5rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("aside", {
    style: {
      width: "var(--rail-left)",
      flex: "none"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      position: "sticky",
      top: "var(--sticky-top)"
    },
    onClick: e => {
      const a = e.target.closest("a");
      if (a) {
        e.preventDefault();
        setSlug(a.getAttribute("href").replace("/docs/", "").replace(/\/$/, ""));
      }
    }
  }, /*#__PURE__*/React.createElement(DocsSidebar, {
    groups: GROUPS,
    currentSlug: slug
  }))), /*#__PURE__*/React.createElement("article", {
    style: prose
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "0 0 0.5rem",
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      letterSpacing: "var(--tracking-label)",
      textTransform: "uppercase",
      color: "var(--muted)"
    }
  }, "Start Here"), /*#__PURE__*/React.createElement("h1", {
    style: {
      margin: "0 0 0.75rem",
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, "Writing nodes"), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "0 0 2.5rem",
      fontSize: "1.125rem",
      lineHeight: 1.65,
      color: "var(--muted)"
    }
  }, "The pieces a document is made of, and where AI fits."), /*#__PURE__*/React.createElement("h2", {
    id: "what-a-node-is",
    style: h2
  }, "What a node is"), /*#__PURE__*/React.createElement("p", {
    style: p
  }, "A node is a component that returns one element of a Word document \u2014 a ", /*#__PURE__*/React.createElement("code", {
    style: inlineCode
  }, "Paragraph"), ", an", " ", /*#__PURE__*/React.createElement("code", {
    style: inlineCode
  }, "Image"), ", a ", /*#__PURE__*/React.createElement("code", {
    style: inlineCode
  }, "Graph"), ". Nodes are ordinary TypeScript files under ", /*#__PURE__*/React.createElement("code", {
    style: inlineCode
  }, "nodes/"), ", and they compose the same way components do on the web."), /*#__PURE__*/React.createElement("div", {
    style: {
      margin: "0 0 1.35rem",
      padding: "1rem 1.15rem",
      border: "1px solid var(--border)",
      borderRadius: "var(--radius-md)",
      background: "var(--paper)"
    }
  }, /*#__PURE__*/React.createElement(CodeFigure, {
    lines: NODE_SOURCE,
    size: "0.82rem"
  })), /*#__PURE__*/React.createElement("h2", {
    id: "state",
    style: h2
  }, "Reading data with useState"), /*#__PURE__*/React.createElement("p", {
    style: p
  }, /*#__PURE__*/React.createElement("code", {
    style: inlineCode
  }, "useState"), " takes a selector over the document's data and returns exactly the fields the node needs. Everything a node reads is declared, so a preview build knows what to ask for."), /*#__PURE__*/React.createElement("blockquote", {
    style: {
      margin: "0 0 1.35rem",
      padding: "0.9rem 1.15rem",
      border: 0,
      borderLeft: "3px solid var(--accent-solid)",
      borderRadius: "0 4px 4px 0",
      background: "var(--accent-tint)"
    }
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      fontSize: "0.95rem",
      color: "var(--ink)"
    }
  }, "A node never reaches outside its selector. If it needs a new field, the type changes and the compiler says so.")), /*#__PURE__*/React.createElement("h2", {
    id: "prompts",
    style: h2
  }, "Handing work to a model"), /*#__PURE__*/React.createElement("table", {
    style: {
      width: "100%",
      margin: "0 0 1.35rem",
      borderCollapse: "collapse",
      fontSize: "0.9rem"
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    style: th
  }, "Hook"), /*#__PURE__*/React.createElement("th", {
    style: th
  }, "Resolves"), /*#__PURE__*/React.createElement("th", {
    style: th
  }, "Status"))), /*#__PURE__*/React.createElement("tbody", null, [["useState", "Data selector", "Stable"], ["useSetPrompts", "Generated text", "Stable"], ["useCondition", "Branch decision", "Stable"], ["useDeriver", "Computed value", "No helper yet"]].map(row => /*#__PURE__*/React.createElement("tr", {
    key: row[0]
  }, /*#__PURE__*/React.createElement("td", {
    style: {
      ...td,
      fontFamily: "var(--font-mono)",
      color: "var(--accent)"
    }
  }, row[0]), /*#__PURE__*/React.createElement("td", {
    style: {
      ...td,
      color: "var(--muted)"
    }
  }, row[1]), /*#__PURE__*/React.createElement("td", {
    style: {
      ...td,
      color: "var(--muted)"
    }
  }, row[2]))))), /*#__PURE__*/React.createElement("h3", {
    id: "set-prompts",
    style: {
      margin: "2rem 0 0.5rem",
      fontSize: "1.2rem",
      fontWeight: 600
    }
  }, "useSetPrompts"), /*#__PURE__*/React.createElement("p", {
    style: p
  }, "The prompt is written where the node is written, not in a separate prompt file. A preview build stops at the placeholder; a request-time build sends the prompt."), /*#__PURE__*/React.createElement("h2", {
    id: "testing",
    style: h2
  }, "Testing a node"), /*#__PURE__*/React.createElement("p", {
    style: p
  }, "Because a document is source, a node can be asserted like any other unit. CI catches the mistake before a recipient does."), /*#__PURE__*/React.createElement("hr", {
    style: {
      margin: "3rem 0",
      border: 0,
      borderTop: "1px solid var(--border)"
    }
  }), /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: "var(--text-sm)",
      color: "var(--muted)"
    }
  }, /*#__PURE__*/React.createElement("a", {
    href: "#",
    style: {
      color: "var(--accent)"
    }
  }, "Edit this page on GitHub \u2197"))), /*#__PURE__*/React.createElement("aside", {
    style: {
      width: "var(--rail-right)",
      flex: "none"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      position: "sticky",
      top: "var(--sticky-top)"
    }
  }, /*#__PURE__*/React.createElement(PageToc, {
    headings: HEADINGS,
    activeSlug: "prompts"
  }))));
}
Object.assign(window, {
  DocsScreen,
  GROUPS,
  HEADINGS
});
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/website/DocsScreen.jsx", error: String((e && e.message) || e) }); }

// ui_kits/website/Hero.jsx
try { (() => {
const {
  Button,
  CopyField,
  Eyebrow,
  PanelBar,
  FileTree,
  PaperPage,
  PaperSection,
  PaperParagraph,
  PaperPlaceholder,
  CodeFigure
} = window.DocxcelerateDesignSystem_cb7f53;
const DEMO_TREE = [{
  kind: "file",
  id: "config",
  label: "docxcelerate.config.json",
  depth: 1
}, {
  kind: "dir",
  label: "documents/",
  depth: 1
}, {
  kind: "document",
  id: "offer-of-admission",
  label: "offer-of-admission/",
  depth: 2,
  sector: "Education",
  children: [{
    kind: "dir",
    label: "nodes/",
    depth: 3
  }, {
    kind: "file",
    id: "greeting",
    label: "greeting.node.tsx",
    depth: 4
  }, {
    kind: "file",
    id: "tutor-note",
    label: "tutor-note.node.tsx",
    depth: 4
  }, {
    kind: "file",
    id: "document",
    label: "document.tsx",
    depth: 3
  }, {
    kind: "file",
    id: "preview-data",
    label: "preview-data.ts",
    depth: 3
  }, {
    kind: "file",
    id: "types",
    label: "types.ts",
    depth: 3
  }]
}, {
  kind: "document",
  id: "repairs-appointment",
  label: "repairs-appointment/",
  depth: 2,
  sector: "Housing"
}, {
  kind: "document",
  id: "policy-renewal",
  label: "policy-renewal/",
  depth: 2,
  sector: "Insurance"
}];
const t = (text, kind) => ({
  text,
  tone: kind
});
const CODE = {
  document: [{
    spans: [t("import", "accent"), t(" { Document, Section, template } ", "ink"), t("from", "accent"), t(" ", "ink"), t('"docxcelerate/template"', "string"), t(";", "muted")]
  }, {
    spans: [t("import", "accent"), t(" { Greeting, Offer, TutorNote } ", "ink"), t("from", "accent"), t(" ", "ink"), t('"./nodes/index.ts"', "string"), t(";", "muted")]
  }, {
    spans: [t("", "muted")]
  }, {
    spans: [t("export const", "accent"), t(" documentTemplate = ", "ink"), t("template", "accent"), t("<OfferData>(", "muted")]
  }, {
    indent: 1,
    spans: [t("<", "muted"), t("Document", "accent"), t(" id=", "ink"), t('"offer-of-admission"', "string"), t(" title=", "ink"), t('"Offer of Admission"', "string"), t(">", "muted")]
  }, {
    indent: 2,
    spans: [t("<", "muted"), t("Section", "accent"), t(" id=", "ink"), t('"offer"', "string"), t(" title=", "ink"), t('"Your offer"', "string"), t(">", "muted")]
  }, {
    indent: 3,
    spans: [t("<", "muted"), t("Greeting", "accent"), t(" />", "muted")]
  }, {
    indent: 3,
    spans: [t("<", "muted"), t("Offer", "accent"), t(" />", "muted")]
  }, {
    indent: 3,
    spans: [t("<", "muted"), t("TutorNote", "accent"), t(" />", "muted")]
  }, {
    indent: 2,
    spans: [t("</", "muted"), t("Section", "accent"), t(">", "muted")]
  }, {
    indent: 1,
    spans: [t("</", "muted"), t("Document", "accent"), t(">", "muted")]
  }, {
    spans: [t(");", "muted")]
  }],
  "tutor-note": [{
    spans: [t("export const", "accent"), t(" TutorNote: ", "ink"), t("Paragraph", "accent"), t(" = () => {", "muted")]
  }, {
    indent: 1,
    spans: [t("const", "accent"), t(" [state] = ", "ink"), t("useState", "accent"), t("((data: OfferData) => ({", "muted")]
  }, {
    indent: 2,
    spans: [t("interviewer: data.interviewer,", "ink")]
  }, {
    indent: 2,
    spans: [t("portfolioTheme: data.portfolioTheme,", "ink")]
  }, {
    indent: 1,
    spans: [t("}));", "muted")]
  }, {
    spans: [t("", "muted")]
  }, {
    indent: 1,
    spans: [t("useSetPrompts", "accent"), t("({ generalPrompt: ", "muted"), t('"A short personal note."', "string"), t(" });", "muted")]
  }, {
    spans: [t("", "muted")]
  }, {
    indent: 1,
    spans: [t("return", "accent"), t(" <", "muted"), t("Paragraph", "accent"), t(" id=", "ink"), t('"tutor-note"', "string"), t(" />;", "muted")]
  }, {
    spans: [t("};", "muted")]
  }]
};
CODE.greeting = CODE["tutor-note"];
CODE.config = [{
  spans: [t("{", "muted")]
}, {
  indent: 1,
  spans: [t('"documentsDir"', "ink"), t(": ", "muted"), t('"documents"', "string"), t(",", "muted")]
}, {
  indent: 1,
  spans: [t('"outDir"', "ink"), t(": ", "muted"), t('"dist"', "string"), t(",", "muted")]
}, {
  indent: 1,
  spans: [t('"engine"', "ink"), t(": { ", "muted"), t('"url"', "ink"), t(": ", "muted"), t('"http://localhost:8787"', "string"), t(" }", "muted")]
}, {
  spans: [t("}", "muted")]
}];
CODE["preview-data"] = CODE.config;
CODE.types = CODE["tutor-note"];

/**
 * The homepage hero: headline, lead, CTAs, install command, release line — then
 * the workspace demo running wider than the prose above it, because it is the
 * argument rather than an illustration of one.
 */
function Hero() {
  const [file, setFile] = React.useState("document");
  const [open, setOpen] = React.useState(["offer-of-admission"]);
  const [treeOpen, setTreeOpen] = React.useState(false);
  return /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("section", {
    style: {
      maxWidth: "var(--container-wide)",
      margin: "0 auto",
      padding: "4rem var(--gutter-sm) 2.5rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      gap: "1.5rem",
      textAlign: "center"
    }
  }, /*#__PURE__*/React.createElement("h1", {
    style: {
      margin: 0,
      maxWidth: "48rem",
      fontSize: "clamp(2.6rem,7vw,4.5rem)",
      fontWeight: 600,
      lineHeight: 1.02
    }
  }, "Documents as components.", /*#__PURE__*/React.createElement("br", null), "DOCX as output."), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      maxWidth: "36rem",
      fontSize: "1.125rem",
      lineHeight: 1.65,
      color: "var(--muted)"
    }
  }, "Compose documents from small typed components, using the JSX you already write. Easily incorporate AI to write or make decisions about the contents. Take advantage of our engine to generate documents at scale."), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      justifyContent: "center",
      gap: "0.75rem"
    }
  }, /*#__PURE__*/React.createElement(Button, {
    href: "#docs"
  }, "Get started \u2192"), /*#__PURE__*/React.createElement(Button, {
    variant: "secondary",
    href: "#engine"
  }, "How it works")), /*#__PURE__*/React.createElement(CopyField, {
    style: {
      marginTop: "0.25rem"
    }
  }), /*#__PURE__*/React.createElement("p", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: "0.5rem",
      margin: 0,
      fontFamily: "var(--font-mono)",
      fontSize: "0.68rem",
      color: "var(--muted)"
    }
  }, /*#__PURE__*/React.createElement("span", null, "v0.4.1"), /*#__PURE__*/React.createElement("span", null, "npm"), /*#__PURE__*/React.createElement("span", null, "JSR"), /*#__PURE__*/React.createElement("span", {
    "aria-hidden": "true"
  }, "\xB7"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontVariantNumeric: "tabular-nums"
    }
  }, "sha512-Qy8Xk3PmM1vB7fCq\u2026")))), /*#__PURE__*/React.createElement("section", {
    style: {
      maxWidth: "var(--container-wide)",
      margin: "0 auto",
      padding: "0 var(--gutter-sm) 4rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "50% 0.25rem minmax(0,1fr)",
      height: "38rem",
      border: "1px solid var(--border-firm)",
      borderRadius: "var(--radius-lg)",
      overflow: "hidden",
      background: "var(--surface-2)",
      boxShadow: "var(--shadow-shell)"
    }
  }, /*#__PURE__*/React.createElement("section", {
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0
    }
  }, /*#__PURE__*/React.createElement(PanelBar, {
    path: "my-documents/",
    hint: "3 documents \xB7 24 files"
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: "0.5rem"
    }
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => setTreeOpen(!treeOpen),
    "aria-expanded": treeOpen,
    "aria-label": treeOpen ? "Hide workspace files" : "Show workspace files",
    style: {
      display: "grid",
      placeItems: "center",
      width: "1.5rem",
      height: "1.5rem",
      border: 0,
      borderRadius: "3px",
      background: treeOpen ? "var(--surface-2)" : "transparent",
      color: treeOpen ? "var(--accent)" : "var(--muted)",
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("svg", {
    viewBox: "0 0 16 16",
    width: "15",
    height: "15",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "1.3",
    "aria-hidden": "true"
  }, /*#__PURE__*/React.createElement("rect", {
    x: "1.5",
    y: "2.5",
    width: "13",
    height: "11",
    rx: "1.5"
  }), /*#__PURE__*/React.createElement("line", {
    x1: "6",
    y1: "2.5",
    x2: "6",
    y2: "13.5"
  }))), /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-eyebrow)",
      letterSpacing: "0.12em",
      textTransform: "uppercase",
      color: "var(--muted)"
    }
  }, "3 documents \xB7 24 files"))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: treeOpen ? "14rem minmax(0,1fr)" : "0 minmax(0,1fr)",
      flex: 1,
      minHeight: 0,
      transition: "grid-template-columns var(--duration-slow) var(--ease)"
    }
  }, /*#__PURE__*/React.createElement(FileTree, {
    nodes: DEMO_TREE,
    selected: file,
    expanded: open,
    onSelect: setFile,
    onToggle: id => setOpen(open.includes(id) ? [] : [id]),
    style: {
      visibility: treeOpen ? "visible" : "hidden",
      borderRight: treeOpen ? "1px solid var(--border)" : "0"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 0,
      minHeight: 0,
      overflow: "auto",
      background: "var(--paper)",
      padding: "1rem 1.1rem"
    }
  }, /*#__PURE__*/React.createElement(CodeFigure, {
    lines: CODE[file] ?? CODE.document,
    size: "0.76rem"
  })))), /*#__PURE__*/React.createElement("div", {
    style: {
      background: "var(--border)",
      cursor: "col-resize"
    },
    role: "separator",
    "aria-orientation": "vertical"
  }), /*#__PURE__*/React.createElement("section", {
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0
    }
  }, /*#__PURE__*/React.createElement(PanelBar, {
    path: "Offer of Admission \xB7 A4",
    hint: "Fit \xB7 100% \xB7 \u2212 +"
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      minHeight: 0,
      overflow: "auto",
      padding: "1rem",
      display: "grid",
      placeContent: "safe center",
      background: "var(--surface-2)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      boxShadow: "var(--shadow-page)",
      width: "calc(210mm * 0.52)",
      height: "calc(297mm * 0.52)"
    }
  }, /*#__PURE__*/React.createElement(PaperPage, {
    title: "Offer of Admission",
    scale: 0.52,
    style: {
      boxShadow: "none"
    }
  }, /*#__PURE__*/React.createElement(PaperSection, {
    heading: "Your offer"
  }, /*#__PURE__*/React.createElement(PaperParagraph, null, "Dear Maya, we are delighted to offer you a place on the BSc Architecture programme beginning in September."), /*#__PURE__*/React.createElement(PaperParagraph, {
    id: "tutor-note",
    dynamic: true
  }, "A short personal note from Dr Priya Raman about Maya Oyelaran's interview.")), /*#__PURE__*/React.createElement(PaperSection, {
    heading: "Conditions"
  }, /*#__PURE__*/React.createElement(PaperParagraph, null, "This offer is conditional on the grades listed in your application being confirmed by the awarding body."), /*#__PURE__*/React.createElement(PaperParagraph, null, "Please confirm your acceptance within 30 days of this notice."), /*#__PURE__*/React.createElement(PaperPlaceholder, null, "Image: campus map")), /*#__PURE__*/React.createElement(PaperSection, {
    heading: "Next steps"
  }, /*#__PURE__*/React.createElement(PaperParagraph, null, "Your tutor will contact you before the end of the month to arrange enrolment."))))))), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: "3rem 0 0",
      textAlign: "center",
      fontSize: "0.875rem",
      color: "var(--muted)"
    }
  }, "Real document preview being rendered, no screenshots and no hidden tricks.")));
}
Object.assign(window, {
  Hero,
  DEMO_TREE,
  CODE
});
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/website/Hero.jsx", error: String((e && e.message) || e) }); }

// ui_kits/website/Points.jsx
try { (() => {
const {
  Eyebrow,
  CodeFigure,
  DiffFigure,
  NumberedSteps,
  Button,
  Badge,
  Card
} = window.DocxcelerateDesignSystem_cb7f53;
const POINTS = [{
  label: "Authoring",
  title: "Write documents like websites",
  body: "A document is a tree of typed components. If you have written React you already know the shape of this: props, composition, small files. A frontend engineer is productive on the first afternoon rather than learning a template language first."
}, {
  label: "AI",
  title: "AI at the component level",
  body: "AI goes in through hooks, inside the components you already write. A component hands the model its context and what you want written, so it produces that one part of the document while everything around it stays deterministic."
}, {
  label: "Change control",
  title: "Documents live in your repo",
  body: "Because a document is source, changing a sentence is a pull request that gets diffed, reviewed, and stays attributable a year later when someone asks who altered the arrears wording. Tests assert a document renders what you expect."
}];
const t = (text, kind) => ({
  text,
  tone: kind
});
const TREE_FIGURE = [{
  spans: [t("<", "muted"), t("Document", "accent"), t(" title=", "ink"), t('"Offer of Admission"', "string"), t(">", "muted")]
}, {
  indent: 1,
  spans: [t("<", "muted"), t("Section", "accent"), t(" title=", "ink"), t('"Your offer"', "string"), t(">", "muted")]
}, {
  indent: 2,
  spans: [t("<", "muted"), t("Greeting", "accent"), t(" />", "muted")]
}, {
  indent: 2,
  spans: [t("<", "muted"), t("Offer", "accent"), t(" />", "muted")]
}, {
  indent: 2,
  spans: [t("<", "muted"), t("TutorNote", "accent"), t(" />", "muted")]
}, {
  indent: 1,
  spans: [t("</", "muted"), t("Section", "accent"), t(">", "muted")]
}, {
  indent: 1,
  spans: [t("<", "muted"), t("Section", "accent"), t(" title=", "ink"), t('"Conditions"', "string"), t(">", "muted")]
}, {
  indent: 2,
  spans: [t("<", "muted"), t("Conditions", "accent"), t(" />", "muted")]
}, {
  indent: 1,
  spans: [t("</", "muted"), t("Section", "accent"), t(">", "muted")]
}, {
  spans: [t("</", "muted"), t("Document", "accent"), t(">", "muted")]
}];

/** The AI point's figure: data in, a prompted node, the placeholder that stands in until the engine runs. */
function AiFigure() {
  const arrow = /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      justifyContent: "center",
      padding: "0.375rem 0"
    },
    "aria-hidden": "true"
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      width: 0,
      height: 0,
      borderLeft: "4px solid transparent",
      borderRight: "4px solid transparent",
      borderTop: "5px solid var(--border-firm)"
    }
  }));
  return /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "var(--text-code-figure)",
      lineHeight: 1.75,
      color: "var(--muted)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      border: "1px solid var(--border-firm)",
      borderRadius: "var(--radius-sm)",
      background: "var(--paper)",
      padding: "0.75rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "baseline",
      justifyContent: "space-between",
      gap: "0.75rem"
    }
  }, /*#__PURE__*/React.createElement("span", null, "state: ", /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--rubric)"
    }
  }, "{ interviewer, portfolioTheme }")), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: "0.6rem",
      letterSpacing: "0.1em"
    }
  }, "DATA")), /*#__PURE__*/React.createElement("div", null, "tone: ", /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--rubric)"
    }
  }, "\"warm, never effusive\""))), arrow, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: "0.75rem",
      borderRadius: "var(--radius-sm)",
      background: "var(--accent-solid)",
      padding: "0.625rem 0.75rem"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--on-accent)"
    }
  }, "<TutorNote />"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: "0.6rem",
      letterSpacing: "0.1em",
      color: "var(--on-accent)",
      opacity: 0.8
    }
  }, "useSetPrompts")), arrow, /*#__PURE__*/React.createElement("div", {
    style: {
      borderLeft: "2px solid var(--accent-solid)",
      paddingLeft: "0.625rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "var(--font-serif)",
      fontSize: "0.8rem",
      lineHeight: 1.4,
      color: "var(--ink)"
    }
  }, "A short personal note from Dr Priya Raman about Maya Oyelaran's interview."), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: "0.5rem",
      fontSize: "0.6rem",
      letterSpacing: "0.1em"
    }
  }, "PLACEHOLDER")));
}

/** Three peers separated by vertical hairlines — not cards. */
function Points() {
  const figures = [/*#__PURE__*/React.createElement(CodeFigure, {
    lines: TREE_FIGURE
  }), /*#__PURE__*/React.createElement(AiFigure, null), /*#__PURE__*/React.createElement(DiffFigure, {
    hunks: [{
      file: "nodes/arrears.node.tsx",
      removed: ["within 14 days of this notice"],
      added: ["within 30 days of this notice"]
    }, {
      file: "nodes/contact.node.tsx",
      removed: ["call us to discuss"],
      added: ["call us to agree a payment plan"]
    }],
    checks: ["arrears-notice renders as expected", "no other document changed"]
  })];
  return /*#__PURE__*/React.createElement("section", {
    style: {
      borderBottom: "1px solid var(--border)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-wide)",
      margin: "0 auto",
      padding: "7rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(3, minmax(0,1fr))"
    }
  }, POINTS.map((point, index) => /*#__PURE__*/React.createElement("div", {
    key: point.label,
    style: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0,
      borderLeft: index > 0 ? "1px solid var(--border)" : undefined,
      paddingLeft: index > 0 ? "2.75rem" : 0,
      paddingRight: index < POINTS.length - 1 ? "2.75rem" : 0
    }
  }, /*#__PURE__*/React.createElement(Eyebrow, {
    index: index + 1,
    style: {
      marginBottom: "1rem"
    }
  }, point.label), /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: "1.5rem"
    }
  }, figures[index]), /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: "0 0 0.75rem",
      maxWidth: "11em",
      fontSize: "1.75rem",
      lineHeight: 1.15,
      fontWeight: 600,
      textWrap: "balance"
    }
  }, point.title), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      fontSize: "0.975rem",
      lineHeight: 1.75,
      color: "var(--muted)",
      textWrap: "pretty"
    }
  }, point.body))))));
}

/** The engine: copy on the left, the numbered spine on the right. */
function Engine() {
  const artefact = (label, value) => /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: "0.75rem",
      border: "1px solid var(--border)",
      borderRadius: "var(--radius-sm)",
      background: "var(--paper)",
      padding: "0.5rem 0.75rem",
      fontFamily: "var(--font-mono)",
      fontSize: "0.7rem",
      color: "var(--muted)"
    }
  }, /*#__PURE__*/React.createElement("span", null, value), /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: "0.6rem",
      letterSpacing: "0.1em",
      textTransform: "uppercase"
    }
  }, label));
  return /*#__PURE__*/React.createElement("section", {
    id: "engine",
    style: {
      borderBottom: "1px solid var(--border)",
      background: "var(--surface)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-page)",
      margin: "0 auto",
      padding: "7rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "1fr 1fr",
      gap: "4rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1.25rem",
      maxWidth: "36rem"
    }
  }, /*#__PURE__*/React.createElement(Eyebrow, null, "The engine"), /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, "Publish once. Scale it out."), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, "The engine is where documents are actually written. It fills in your data, runs the AI, and returns the finished document. Any kind of node can use AI, not only paragraphs."), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, "You publish a template to the engine once. After that, any system can call its API with a set of data and get a document back. A free engine is available to self-host."), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      gap: "0.75rem",
      marginTop: "0.5rem"
    }
  }, /*#__PURE__*/React.createElement(Button, {
    href: "#docs"
  }, "How the engine works \u2192"), /*#__PURE__*/React.createElement(Button, {
    variant: "quiet",
    href: "#"
  }, "Managed cloud ", /*#__PURE__*/React.createElement(Badge, null, "soon")))), /*#__PURE__*/React.createElement(NumberedSteps, {
    steps: [{
      title: "Build",
      detail: "The framework turns your document into a package. This step runs on your machine.",
      figure: artefact("artifact", "dist/offer-of-admission.dxcl")
    }, {
      title: "Publish",
      detail: "You send the package to an engine. The engine stores it and gives it a name.",
      figure: artefact("template", "offer-of-admission@3")
    }, {
      title: "Write",
      detail: "Your application sends a set of data. The engine returns the finished document.",
      figure: artefact("response", "POST /generate → 200 .docx")
    }]
  }))));
}
Object.assign(window, {
  Points,
  Engine,
  AiFigure,
  TREE_FIGURE
});
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/website/Points.jsx", error: String((e && e.message) || e) }); }

// ui_kits/website/Sections.jsx
try { (() => {
const {
  Eyebrow,
  Button,
  Card,
  DocCard,
  LicenceSheet,
  PillTab,
  Badge
} = window.DocxcelerateDesignSystem_cb7f53;
const TRAITS = [{
  title: "One template, every recipient",
  body: "Publish the template once, then call it per person. One report or a hundred thousand is the same call, repeated.",
  figure: "volume"
}, {
  title: "The same document every time",
  body: "Anything you did not mark as generated renders identically on every run. Only the parts you chose can vary.",
  figure: "determinism"
}, {
  title: "Called by the systems you run",
  body: "Your CRM, case system or billing platform posts its data and gets a .docx back. Nobody exports a spreadsheet or opens Word.",
  figure: "integration"
}, {
  title: "Reproducible a year later",
  body: "Templates are versioned, so any document can be rebuilt from the same template and the same data. An audit gets a build, not an archive.",
  figure: "reproducibility"
}];
const mono = {
  fontFamily: "var(--font-mono)",
  fontSize: "0.7rem",
  lineHeight: 1.75,
  color: "var(--muted)"
};
function TraitFigure({
  kind
}) {
  if (kind === "volume") {
    return /*#__PURE__*/React.createElement("div", {
      style: {
        ...mono,
        display: "grid",
        gap: "0.375rem"
      }
    }, ["offer-of-admission@3 → 41,208 recipients", "repairs-appointment@7 → 12,940 recipients", "policy-renewal@2 → 388,116 recipients"].map(row => /*#__PURE__*/React.createElement("div", {
      key: row,
      style: {
        display: "flex",
        justifyContent: "space-between",
        gap: "1rem",
        borderBottom: "1px solid var(--border)",
        paddingBottom: "0.375rem"
      }
    }, /*#__PURE__*/React.createElement("span", null, row.split(" → ")[0]), /*#__PURE__*/React.createElement("span", {
      style: {
        color: "var(--ink)",
        fontVariantNumeric: "tabular-nums"
      }
    }, row.split(" → ")[1]))));
  }
  if (kind === "determinism") {
    return /*#__PURE__*/React.createElement("div", {
      style: mono
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        background: "var(--accent-tint)",
        color: "var(--accent)",
        margin: "0 -0.375rem",
        padding: "0 0.375rem"
      }
    }, "generated \xB7 tutor-note"), /*#__PURE__*/React.createElement("div", null, "static \xB7 offer"), /*#__PURE__*/React.createElement("div", null, "static \xB7 conditions"), /*#__PURE__*/React.createElement("div", null, "static \xB7 signoff"));
  }
  if (kind === "integration") {
    return /*#__PURE__*/React.createElement("div", {
      style: mono
    }, /*#__PURE__*/React.createElement("div", null, "POST /v1/documents/offer-of-admission"), /*#__PURE__*/React.createElement("div", {
      style: {
        color: "var(--rubric)"
      }
    }, '{ "recipientName": "Maya Oyelaran" }'), /*#__PURE__*/React.createElement("div", {
      style: {
        color: "var(--accent)"
      }
    }, "200 \xB7 application/vnd\u2026wordprocessingml"));
  }
  return /*#__PURE__*/React.createElement("div", {
    style: mono
  }, /*#__PURE__*/React.createElement("div", null, "offer-of-admission@1 \xB7 12 Jan"), /*#__PURE__*/React.createElement("div", null, "offer-of-admission@2 \xB7 04 Mar"), /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--accent)"
    }
  }, "offer-of-admission@3 \xB7 19 Jun \u2190 current"));
}

/** Scale: one wide card carrying the section's argument, three narrower under it. */
function Scale() {
  return /*#__PURE__*/React.createElement("section", null, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-page)",
      margin: "0 auto",
      padding: "7rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1rem",
      maxWidth: "42rem",
      marginBottom: "3rem"
    }
  }, /*#__PURE__*/React.createElement(Eyebrow, null, "Enterprise"), /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, "Built for the reports you already write by hand"), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, "A firm producing the same report hundreds of times a month already has the layout and most of the wording. What changes is the person it is written for. Rebuild that report as components, keep every run identical, and hand a model only the parts that depend on who is reading it.")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(3, minmax(0,1fr))",
      gap: "1.25rem"
    }
  }, TRAITS.map((trait, index) => /*#__PURE__*/React.createElement(Card, {
    key: trait.title,
    variant: "paper",
    style: {
      padding: "var(--pad-card-lg)",
      gridColumn: index === 0 ? "span 3" : undefined,
      display: index === 0 ? "grid" : "flex",
      gridTemplateColumns: index === 0 ? "1fr 1fr" : undefined,
      alignItems: index === 0 ? "center" : undefined,
      gap: index === 0 ? "3.5rem" : undefined
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      marginBottom: "0.75rem",
      fontFamily: "var(--font-mono)",
      fontSize: "0.7rem",
      color: "var(--muted)",
      fontVariantNumeric: "tabular-nums"
    }
  }, String(index + 1).padStart(2, "0")), /*#__PURE__*/React.createElement("h3", {
    style: {
      margin: "0 0 0.5rem",
      fontSize: "1.125rem",
      fontWeight: 600
    }
  }, trait.title), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      fontSize: "0.95rem",
      lineHeight: 1.75,
      color: "var(--muted)",
      textWrap: "pretty"
    }
  }, trait.body)), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: index === 0 ? 0 : "auto",
      paddingTop: index === 0 ? 0 : "1.75rem"
    }
  }, /*#__PURE__*/React.createElement(TraitFigure, {
    kind: trait.figure
  })))))));
}

/** Open source: prose beside the licence, drawn as paper. */
function OpenSource() {
  return /*#__PURE__*/React.createElement("section", {
    style: {
      borderTop: "1px solid var(--border)",
      background: "var(--surface)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-page)",
      margin: "0 auto",
      padding: "7rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "1fr 420px",
      gap: "4rem",
      alignItems: "start"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1.25rem",
      maxWidth: "36rem"
    }
  }, /*#__PURE__*/React.createElement(Eyebrow, null, "Open source"), /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, "Open source, and built to stay that way"), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)",
      textWrap: "pretty"
    }
  }, "The framework, the renderers, the node model and the CLI are MIT licensed and developed in the open. Read the code that writes your documents, fork it, or vendor it into your own build."), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)",
      textWrap: "pretty"
    }
  }, "The engine is free to self-host, so running documents at scale never depends on a vendor staying in business or a price list staying the same. It is the convenience, not the way in."), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      gap: "0.75rem",
      marginTop: "0.5rem"
    }
  }, /*#__PURE__*/React.createElement(Button, {
    href: "https://github.com/LiamDotPro/Docxcelerate"
  }, "View the source \u2192"), /*#__PURE__*/React.createElement(Button, {
    variant: "tertiary",
    href: "https://www.npmjs.com/package/docxcelerate"
  }, "See it on npm"))), /*#__PURE__*/React.createElement(LicenceSheet, null))));
}
const CARDS = [{
  title: "Start here",
  blurb: "Scaffold a workspace and open your first document."
}, {
  title: "Documents and nodes",
  blurb: "The component model a document is built from."
}, {
  title: "The node model",
  blurb: "Every node type, with a preview of each."
}, {
  title: "Writing nodes",
  blurb: "The pieces a document is made of."
}, {
  title: "CLI commands",
  blurb: "dxcl init, new, node."
}, {
  title: "Package entrypoints",
  blurb: "Everything importable, and what it exports."
}];

/** The documentation index, on paper so the page keeps alternating grounds. */
function DocsIndex({
  onOpenDocs
}) {
  return /*#__PURE__*/React.createElement("section", {
    id: "docs",
    style: {
      borderTop: "1px solid var(--border)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-page)",
      margin: "0 auto",
      padding: "7rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexWrap: "wrap",
      alignItems: "flex-end",
      justifyContent: "space-between",
      gap: "1.25rem 2.5rem",
      marginBottom: "2.5rem"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1rem",
      maxWidth: "42rem"
    }
  }, /*#__PURE__*/React.createElement(Eyebrow, null, "Documentation"), /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, "Everything is documented"), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, "Every node type, every CLI flag, and every file a build writes, with previews rendered by the real renderer, so nothing on the page can describe a helper that no longer exists.")), /*#__PURE__*/React.createElement("a", {
    href: "#",
    onClick: onOpenDocs,
    style: {
      fontSize: "var(--text-sm)",
      fontWeight: 500,
      whiteSpace: "nowrap",
      color: "var(--accent)",
      textDecoration: "none"
    }
  }, "Full documentation \u2192")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gridTemplateColumns: "repeat(3, minmax(0,1fr))",
      gap: "1.25rem"
    }
  }, CARDS.map((card, index) => /*#__PURE__*/React.createElement(DocCard, {
    key: card.title,
    index: index + 1,
    title: card.title,
    blurb: card.blurb,
    href: "#",
    onClick: onOpenDocs
  })))));
}
const SKILL = `---
name: docxcelerate
description: Build DOCX document projects with Docxcelerate — typed
  components, preview, and the generation engine.
---

# Docxcelerate

A document is a tree of typed components. Compose it with JSX from
\`docxcelerate/template\`, preview it locally, then publish it to an
engine that fills in data and runs the prompts.

## The model

- \`Document\` is the root. It owns the id, title and styling.
- \`Section\` groups nodes and prints a heading.
- \`Paragraph\`, \`Image\`, \`Graph\` and \`TableOfContents\` are the leaves.

## Where AI fits

\`useSetPrompts\` hands the model a node's context and what you want
written. Everything you do not mark as generated renders identically
on every run.`;
const AGENTS = [{
  id: "claude-code",
  label: "Claude Code",
  path: ".claude/skills/docxcelerate/",
  note: "Copy the folder into .claude/skills/ for one project, or ~/.claude/skills/ for all of them. It loads itself when a document project turns up."
}, {
  id: "cursor",
  label: "Cursor",
  path: ".cursor/rules/docxcelerate.mdc",
  note: "Save it as a rule file and Cursor reads it in that project. Or @-mention the file in chat when you want it."
}, {
  id: "copilot",
  label: "Copilot",
  path: ".github/copilot-instructions.md",
  note: "Paste it into .github/copilot-instructions.md and Copilot applies it across the repository."
}, {
  id: "agents-md",
  label: "AGENTS.md",
  path: "AGENTS.md",
  note: "Codex, Gemini CLI, Aider and Cline all read AGENTS.md at the root of the repo. Paste it in, or link to it if your agent opens files."
}];

/** The agent skill, printed in full at the foot of the page. */
function AgentSkill() {
  const [agent, setAgent] = React.useState(AGENTS[0].id);
  const current = AGENTS.find(a => a.id === agent);
  const [label, setLabel] = React.useState("Copy");
  return /*#__PURE__*/React.createElement("section", {
    style: {
      borderTop: "1px solid var(--border)",
      background: "var(--surface)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: "var(--container-prose)",
      margin: "0 auto",
      padding: "7rem var(--gutter-sm)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      gap: "1rem",
      maxWidth: "42rem",
      margin: "0 auto 2rem",
      textAlign: "center"
    }
  }, /*#__PURE__*/React.createElement(Eyebrow, null, "Agent skills"), /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, "Hand the whole thing to your agent"), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, "One Markdown file teaches a coding agent how documents are put together here: the component model, the rules that catch agents out, and every command.")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      justifyContent: "center",
      marginBottom: "1.5rem"
    }
  }, /*#__PURE__*/React.createElement(PillTab, {
    items: AGENTS,
    value: agent,
    onChange: setAgent
  })), /*#__PURE__*/React.createElement("figure", {
    style: {
      margin: 0,
      overflow: "hidden",
      borderRadius: "var(--radius-md)",
      border: "1px solid var(--border)"
    }
  }, /*#__PURE__*/React.createElement("figcaption", {
    style: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: "0.75rem",
      borderBottom: "1px solid var(--border)",
      background: "var(--surface-2)",
      padding: "0.625rem 1rem"
    }
  }, /*#__PURE__*/React.createElement("code", {
    style: {
      fontFamily: "var(--font-mono)",
      fontSize: "0.75rem",
      color: "var(--muted)"
    }
  }, current.path), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => {
      navigator.clipboard?.writeText(SKILL);
      setLabel("Copied");
      setTimeout(() => setLabel("Copy"), 1600);
    },
    style: {
      border: 0,
      background: "transparent",
      borderRadius: "var(--radius-sm)",
      padding: "0.125rem 0.375rem",
      fontFamily: "var(--font-mono)",
      fontSize: "0.75rem",
      color: label === "Copy" ? "var(--muted)" : "var(--accent)",
      cursor: "pointer"
    }
  }, label)), /*#__PURE__*/React.createElement("pre", {
    style: {
      margin: 0,
      padding: "1.15rem 1.3rem",
      maxHeight: "22rem",
      overflow: "auto",
      background: "var(--paper)",
      fontFamily: "var(--font-mono)",
      fontSize: "0.78rem",
      lineHeight: 1.7,
      color: "var(--ink)"
    }
  }, SKILL)), /*#__PURE__*/React.createElement("p", {
    style: {
      maxWidth: "34rem",
      margin: "1.5rem auto 0",
      textAlign: "center",
      fontSize: "var(--text-sm)",
      lineHeight: 1.7,
      color: "var(--muted)",
      textWrap: "pretty"
    }
  }, current.note), /*#__PURE__*/React.createElement("p", {
    style: {
      maxWidth: "34rem",
      margin: "0.5rem auto 0",
      textAlign: "center",
      fontSize: "var(--text-sm)",
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, "It is plain Markdown, and four reference files sit beside it in skills/docxcelerate/.")));
}
Object.assign(window, {
  Scale,
  OpenSource,
  DocsIndex,
  AgentSkill,
  TraitFigure
});
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/website/Sections.jsx", error: String((e && e.message) || e) }); }

// ui_kits/website/shared.jsx
try { (() => {
const DS = window.DocxcelerateDesignSystem_cb7f53;

/** Section frame: max-width container, 6/7rem vertical band, optional ground + hairlines. */
function Band({
  ground = "paper",
  width = "var(--container-page)",
  top,
  bottom = true,
  style,
  children
}) {
  return /*#__PURE__*/React.createElement("section", {
    style: {
      background: ground === "surface" ? "var(--surface)" : "var(--paper)",
      borderTop: top ? "1px solid var(--border)" : undefined,
      borderBottom: bottom ? "1px solid var(--border)" : undefined,
      ...style
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      maxWidth: width,
      margin: "0 auto",
      padding: "6rem var(--gutter-sm)"
    }
  }, children));
}

/** Section heading block: eyebrow, serif h2, lead. */
function SectionHead({
  eyebrow,
  title,
  lead,
  width = "42rem",
  style
}) {
  const {
    Eyebrow
  } = DS;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: "1rem",
      maxWidth: width,
      ...style
    }
  }, eyebrow ? /*#__PURE__*/React.createElement(Eyebrow, null, eyebrow) : null, /*#__PURE__*/React.createElement("h2", {
    style: {
      margin: 0,
      fontSize: "2.25rem",
      fontWeight: 600
    }
  }, title), lead ? /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.7,
      color: "var(--muted)"
    }
  }, lead) : null);
}
const tone = (text, t) => ({
  text,
  tone: t
});
Object.assign(window, {
  Band,
  SectionHead,
  tone,
  DS
});
})(); } catch (e) { __ds_ns.__errors.push({ path: "ui_kits/website/shared.jsx", error: String((e && e.message) || e) }); }

__ds_ns.Mark = __ds_scope.Mark;

__ds_ns.Wordmark = __ds_scope.Wordmark;

__ds_ns.CodeFigure = __ds_scope.CodeFigure;

__ds_ns.DiffFigure = __ds_scope.DiffFigure;

__ds_ns.DocCard = __ds_scope.DocCard;

__ds_ns.LicenceSheet = __ds_scope.LicenceSheet;

__ds_ns.NumberedSteps = __ds_scope.NumberedSteps;

__ds_ns.PullQuote = __ds_scope.PullQuote;

__ds_ns.Badge = __ds_scope.Badge;

__ds_ns.Button = __ds_scope.Button;

__ds_ns.Card = __ds_scope.Card;

__ds_ns.CopyField = __ds_scope.CopyField;

__ds_ns.Eyebrow = __ds_scope.Eyebrow;

__ds_ns.PillTab = __ds_scope.PillTab;

__ds_ns.DocsSidebar = __ds_scope.DocsSidebar;

__ds_ns.PageToc = __ds_scope.PageToc;

__ds_ns.SiteFooter = __ds_scope.SiteFooter;

__ds_ns.SiteHeader = __ds_scope.SiteHeader;

__ds_ns.FileTree = __ds_scope.FileTree;

__ds_ns.PanelBar = __ds_scope.PanelBar;

__ds_ns.PaperPage = __ds_scope.PaperPage;

__ds_ns.PaperSection = __ds_scope.PaperSection;

__ds_ns.PaperParagraph = __ds_scope.PaperParagraph;

__ds_ns.PaperPlaceholder = __ds_scope.PaperPlaceholder;

__ds_ns.WordShell = __ds_scope.WordShell;

})();
