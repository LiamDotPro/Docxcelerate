/**
 * Builds every brand asset the site and the repository need, from one geometry.
 *
 * The mark is four sheared text lines with square terminals — the third one is
 * the generated line, so it takes the accent and runs longest. It is defined
 * once here, in `markGeometry()`, and everything else in this file is a framing
 * decision around it: which ground it sits on, how much air it gets, and which
 * of the palette's two accents applies. Nothing traces the shape twice.
 *
 * Two rasterisers, for two different reasons:
 *
 *   - Icons are pure geometry, so `sharp` renders them. That is portable: any
 *     machine with the site's dev dependencies installed reproduces them.
 *   - The share cards carry type, and the site's wordmark is a serif from a
 *     system stack rather than a webfont. Only a browser resolves that stack
 *     the way a reader's browser will, so headless Chrome screenshots them at
 *     2x and `sharp` downsamples. That is why the outputs are committed: a
 *     machine without Chrome should never need to rebuild them.
 *
 * Run it with `npm run brand`. It is deliberately *not* part of `npm run prep`
 * — the identity changes a few times a year, not every build.
 */

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import sharp from "sharp";

const here = dirname(fileURLToPath(import.meta.url));
const publicDir = join(here, "..", "public");
const brandDir = join(publicDir, "brand");
const repoRoot = join(here, "..", "..");

/**
 * The palette, lifted from src/styles/global.css rather than re-picked.
 *
 * A raster cannot follow a CSS variable, so every baked-in colour here has to
 * be the value the stylesheet would have resolved to. If Ledger Green ever
 * moves, these move with it — that is the only duplication in the identity and
 * it is one grep away from being caught.
 */
const C = {
  paperLight: "#ffffff",
  inkLight: "#16211c",
  accentLight: "#17624a",

  paperDark: "#0f1a15",
  inkDark: "#d7e0da",
  mutedDark: "#93a29a",
  accentDark: "#8ac5ae",

  onAccent: "#ffffff",
  surfaceDark: "#16281f",
  borderDark: "#2b3a33",
};

const SERIF =
  "'Iowan Old Style','Palatino Linotype',Palatino,'Book Antiqua','URW Palladio L',serif";
const SANS = "system-ui,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
const MONO = "ui-monospace,'Cascadia Code','Cascadia Mono',Consolas,'Courier New',monospace";

/* -------------------------------------------------------------------------- */
/* Geometry                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * The mark's ink, in its own tight coordinate space.
 *
 * The design canvas drew the four bars inside a 64x64 box and sheared them in
 * place, which left the ink sitting low and to the left of that box's centre —
 * invisible in a lockup, a plainly crooked favicon. So the shear is baked in
 * and the result translated to the origin: the viewBox below is exactly the
 * ink's bounding box, and every consumer centres it deliberately instead of
 * inheriting whatever padding the 64-box happened to leave.
 *
 * `translate` after `skewX` in the transform list means the points are sheared
 * first and shifted second, which is what puts the top-left bar corner at x=0.
 */
const INK_W = 47.14;
const INK_H = 43;

/**
 * @param {object} options
 * @param {string} options.ink Colour of the three quiet lines.
 * @param {string} options.accent Colour of the third, generated line.
 * @param {boolean} [options.mono] Collapse the opacity steps to one solid ink.
 * @returns {string} The `<g>` element, in the INK_W x INK_H space.
 */
const markGeometry = ({ ink, accent, mono = false }) => {
  const bars = [
    { y: 12, w: 38, fill: ink, opacity: mono ? 1 : 0.38 },
    { y: 24, w: 28, fill: ink, opacity: mono ? 1 : 0.55 },
    { y: 36, w: 42, fill: mono ? ink : accent, opacity: 1 },
    { y: 48, w: 21, fill: ink, opacity: mono ? 1 : 0.55 },
  ];
  const rects = bars
    .map(
      ({ y, w, fill, opacity }) =>
        `    <rect x="14" y="${y}" width="${w}" height="7" fill="${fill}"` +
        (opacity === 1 ? "" : ` opacity="${opacity}"`) +
        " />",
    )
    .join("\n");
  return `  <g transform="translate(-2.3094 -12) skewX(-12)">\n${rects}\n  </g>`;
};

/**
 * A standalone SVG of the mark, centred in a square with a little air.
 *
 * @param {object} options
 * @param {string} options.ink
 * @param {string} options.accent
 * @param {boolean} [options.mono]
 * @param {number} [options.pad] Horizontal air, in ink units, per side.
 * @param {string} [options.title] Accessible name; omitted for decorative use.
 * @returns {string}
 */
const squareMarkSvg = ({ ink, accent, mono = false, pad = 1, title }) => {
  const side = INK_W + pad * 2;
  const top = -(side - INK_H) / 2;
  const label = title
    ? `\n  <title>${title}</title>`
    : `\n  <!-- Decorative: label it at the point of use. -->`;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${-pad} ${top.toFixed(2)} ${side} ${side}" width="${side}" height="${side}" role="img">${label}
${markGeometry({ ink, accent, mono })}
</svg>
`;
};

/* -------------------------------------------------------------------------- */
/* SVG sources                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Marks that follow the reader's theme, for anywhere an `<img>` is all that is
 * available. Inside Astro the mark is an inline component painting in
 * `currentColor` and `var(--accent)`, which is strictly better — it follows the
 * site's own theme toggle, not just the OS preference. These files exist for
 * README-shaped places, where there is no stylesheet to inherit.
 */
const themedMarkSvg = () => {
  const side = INK_W + 2;
  const top = -(side - INK_H) / 2;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="-1 ${top.toFixed(2)} ${side} ${side}" width="${side}" height="${side}" role="img">
  <title>Docxcelerate</title>
  <style>
    .q { fill: ${C.inkLight} }
    .a { fill: ${C.accentLight} }
    @media (prefers-color-scheme: dark) {
      .q { fill: ${C.inkDark} }
      .a { fill: ${C.accentDark} }
    }
  </style>
  <g transform="translate(-2.3094 -12) skewX(-12)">
    <rect class="q" x="14" y="12" width="38" height="7" opacity="0.38" />
    <rect class="q" x="14" y="24" width="28" height="7" opacity="0.55" />
    <rect class="a" x="14" y="36" width="42" height="7" />
    <rect class="q" x="14" y="48" width="21" height="7" opacity="0.55" />
  </g>
</svg>
`;
};

/**
 * The horizontal lockup: mark, then wordmark, at the proportions the identity
 * sets — the mark's ink height matches the cap height of the type beside it.
 *
 * The wordmark stays live `<text>` in the site's serif stack rather than being
 * outlined, because there is no font here to outline it from. That is fine
 * wherever a browser renders the file and wrong wherever one does not, so
 * anything that must look identical everywhere uses banner.png instead.
 *
 * @param {object} options
 * @param {string} options.ink
 * @param {string} options.accent
 * @returns {string}
 */
const lockupSvg = ({ ink, accent }) => {
  const markH = 30;
  const markW = (markH * INK_W) / INK_H;
  const gap = 12;
  const fontSize = 40;
  const width = markW + gap + 236;
  const height = 48;
  const markY = (height - markH) / 2;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width.toFixed(2)} ${height}" width="${width.toFixed(2)}" height="${height}" role="img">
  <title>Docxcelerate</title>
  <g transform="translate(0 ${markY.toFixed(2)}) scale(${(markH / INK_H).toFixed(6)})">
${markGeometry({ ink, accent })}
  </g>
  <text x="${(markW + gap).toFixed(2)}" y="${height * 0.72}" font-family="${SERIF}" font-size="${fontSize}" font-weight="600" letter-spacing="-0.4" fill="${ink}">Docx<tspan fill="${accent}">celerate</tspan></text>
</svg>
`;
};

/* -------------------------------------------------------------------------- */
/* Raster icons                                                              */
/* -------------------------------------------------------------------------- */

/**
 * One app icon: a ground, and the mark centred on it at a stated width.
 *
 * `inkFraction` is the width of the *ink* as a share of the box, not the width
 * of some notional 64-unit slot the ink sits inside. Stating it that way is the
 * only way two icons at different sizes end up looking equally full.
 *
 * @param {object} options
 * @param {number} options.size Pixel side of the square.
 * @param {string|null} options.ground Background fill, or null for transparent.
 * @param {number} options.radius Corner radius in pixels.
 * @param {number} options.inkFraction Ink width / box width.
 * @param {string} options.ink
 * @param {string} options.accent
 * @param {boolean} [options.mono]
 * @returns {string}
 */
const iconSvg = ({ size, ground, radius, inkFraction, ink, accent, mono = false }) => {
  const w = size * inkFraction;
  const scale = w / INK_W;
  const h = INK_H * scale;
  const x = (size - w) / 2;
  const y = (size - h) / 2;
  const bg = ground
    ? `  <rect width="${size}" height="${size}" rx="${radius}" fill="${ground}" />\n`
    : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${size} ${size}" width="${size}" height="${size}">
${bg}  <g transform="translate(${x.toFixed(3)} ${y.toFixed(3)}) scale(${scale.toFixed(6)})">
${markGeometry({ ink, accent, mono })}
  </g>
</svg>
`;
};

/**
 * @param {string} svg
 * @param {number} size
 * @returns {Promise<Buffer>}
 */
const rasterise = (svg, size) =>
  sharp(Buffer.from(svg), { density: 384 })
    .resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png({ compressionLevel: 9 })
    .toBuffer();

/**
 * Packs PNGs into an ICO container.
 *
 * Every browser that still asks for favicon.ico reads PNG-in-ICO, so there is
 * no reason to encode BMP: the payloads go in untouched and only the 22-byte
 * header and per-image directory entries are assembled here.
 *
 * @param {Buffer[]} pngs Square PNGs, smallest first.
 * @param {number[]} sizes Their pixel sides, in the same order.
 * @returns {Buffer}
 */
const buildIco = (pngs, sizes) => {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0); // reserved
  header.writeUInt16LE(1, 2); // 1 = icon
  header.writeUInt16LE(pngs.length, 4);

  const directory = Buffer.alloc(16 * pngs.length);
  let offset = header.length + directory.length;
  pngs.forEach((png, i) => {
    const at = i * 16;
    // 256 is written as 0; nothing here is that large, but the rule is the rule.
    directory.writeUInt8(sizes[i] >= 256 ? 0 : sizes[i], at);
    directory.writeUInt8(sizes[i] >= 256 ? 0 : sizes[i], at + 1);
    directory.writeUInt8(0, at + 2); // palette size: not paletted
    directory.writeUInt8(0, at + 3); // reserved
    directory.writeUInt16LE(1, at + 4); // colour planes
    directory.writeUInt16LE(32, at + 6); // bits per pixel
    directory.writeUInt32LE(png.length, at + 8);
    directory.writeUInt32LE(offset, at + 12);
    offset += png.length;
  });

  return Buffer.concat([header, directory, ...pngs]);
};

/* -------------------------------------------------------------------------- */
/* Share cards                                                                */
/* -------------------------------------------------------------------------- */

const CHROME_CANDIDATES = [
  process.env.CHROME_PATH,
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/usr/bin/google-chrome",
  "/usr/bin/chromium",
].filter((path) => typeof path === "string");

const findChrome = () => {
  const found = CHROME_CANDIDATES.find((path) => existsSync(path));
  if (!found) {
    throw new Error(
      "Chrome is needed to render the share cards, because they carry type from a " +
        "system font stack that only a browser resolves faithfully. Install it, or " +
        "point CHROME_PATH at it. The committed PNGs are already correct, so if you " +
        "are not changing the identity you do not need this.",
    );
  }
  return found;
};

/**
 * Screenshots one HTML document at exactly `width` x `height`.
 *
 * Rendered at 2x and downsampled, which is what makes the serif look like the
 * site's serif rather than like a screenshot of it. The temporary profile keeps
 * the run away from whatever Chrome the author has open.
 *
 * @param {object} options
 * @param {string} options.html
 * @param {number} options.width
 * @param {number} options.height
 * @param {string} options.out Destination path.
 */
const shoot = async ({ html, width, height, out }) => {
  const chrome = findChrome();
  const work = mkdtempSync(join(tmpdir(), "dxcl-brand-"));
  try {
    const page = join(work, "card.html");
    const shot = join(work, "card.png");
    writeFileSync(page, html, "utf8");
    execFileSync(
      chrome,
      [
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--no-first-run",
        "--no-default-browser-check",
        `--user-data-dir=${join(work, "profile")}`,
        "--force-device-scale-factor=2",
        `--window-size=${width},${height}`,
        `--screenshot=${shot}`,
        pathToFileURL(page).href,
      ],
      { stdio: "pipe" },
    );
    const png = await sharp(readFileSync(shot))
      .resize(width, height, { kernel: "lanczos3" })
      .png({ compressionLevel: 9 })
      .toBuffer();
    write(out, png);
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
};

/**
 * The shell every card shares: no margin, no scrollbars, exact box.
 *
 * @param {object} options
 * @param {number} options.width
 * @param {number} options.height
 * @param {string} options.ground
 * @param {string} options.body
 * @returns {string}
 */
const cardHtml = ({ width, height, ground, body }) => `<!doctype html>
<meta charset="utf-8">
<style>
  html, body { margin: 0; padding: 0; width: ${width}px; height: ${height}px; overflow: hidden; }
  body { background: ${ground}; -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }
  * { box-sizing: border-box; }
</style>
${body}
`;

/**
 * The mark as inline SVG for use inside a card, sized by ink width.
 *
 * @param {object} options
 * @param {number} options.width Ink width in CSS pixels.
 * @param {string} options.ink
 * @param {string} options.accent
 * @param {boolean} [options.mono]
 * @param {number} [options.opacity]
 * @returns {string}
 */
const inlineMark = ({ width, ink, accent, mono = false, opacity = 1 }) => {
  const height = (width * INK_H) / INK_W;
  const style = opacity === 1 ? "" : ` style="opacity:${opacity}"`;
  return `<svg width="${width.toFixed(2)}" height="${height.toFixed(2)}" viewBox="0 0 ${INK_W} ${INK_H}"${style}>
${markGeometry({ ink, accent, mono })}
</svg>`;
};

/**
 * The wordmark, split the way the site splits it.
 *
 * The identity canvas accents the single `x`; the site has always accented the
 * whole `celerate`, in a serif rather than a grotesk. The site's reading wins
 * here, because a share card that does not look like the page it links to is a
 * worse card — and because switching would mean pulling a webfont into a site
 * that deliberately ships none.
 *
 * @param {object} options
 * @param {number} options.size Font size in pixels.
 * @param {string} options.ink
 * @param {string} options.accent
 * @returns {string}
 */
const wordmark = ({ size, ink, accent }) =>
  `<span style="font-family:${SERIF};font-size:${size}px;font-weight:600;letter-spacing:-0.02em;color:${ink}">Docx<span style="color:${accent}">celerate</span></span>`;

/* -------------------------------------------------------------------------- */
/* Copy                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * Read out of the English dictionary rather than retyped, so the card and the
 * page it links to cannot drift apart. A crude regex is enough: both keys are
 * plain string literals and parsing the TypeScript module would need a compiler.
 *
 * `meta` is the first key in the dictionary, so the first match of each name is
 * the one wanted. The description is authored as several literals joined with
 * `+` across lines, so the pattern takes that whole run and concatenates it —
 * anything that matched only the first literal would silently produce a card
 * with a truncated subtitle, which is worse than failing.
 */
const readCopy = () => {
  const source = readFileSync(join(here, "..", "src", "i18n", "ui", "en.ts"), "utf8");
  const tagline = /tagline:\s*"([^"]+)"/.exec(source);
  const description = /description:\s*((?:"[^"]*"\s*\+\s*)*"[^"]*")/.exec(source);
  if (!tagline) throw new Error("Could not find meta.tagline in en.ts");
  if (!description) throw new Error("Could not find meta.description in en.ts");
  const joined = [...description[1].matchAll(/"([^"]*)"/g)]
    .map((chunk) => chunk[1])
    .join("")
    .replace(/\s+/g, " ")
    .trim();
  if (!joined) throw new Error("meta.description in en.ts read as empty");
  return { tagline: tagline[1], description: joined };
};

const installCommand = () => {
  const source = readFileSync(join(here, "..", "src", "site.ts"), "utf8");
  const match = /INSTALL_COMMAND\s*=\s*"([^"]+)"/.exec(source);
  if (!match) throw new Error("Could not find INSTALL_COMMAND in site.ts");
  return match[1];
};

/* -------------------------------------------------------------------------- */
/* Build                                                                      */
/* -------------------------------------------------------------------------- */

const written = [];

/**
 * @param {string} path
 * @param {string|Buffer} contents
 */
const write = (path, contents) => {
  writeFileSync(path, contents);
  written.push(path.slice(repoRoot.length + 1).replace(/\\/g, "/"));
};

mkdirSync(brandDir, { recursive: true });

/* --- SVG sources ---------------------------------------------------------- */

write(join(brandDir, "mark.svg"), themedMarkSvg());
write(
  join(brandDir, "mark-light.svg"),
  squareMarkSvg({ ink: C.inkLight, accent: C.accentLight, title: "Docxcelerate" }),
);
write(
  join(brandDir, "mark-dark.svg"),
  squareMarkSvg({ ink: C.inkDark, accent: C.accentDark, title: "Docxcelerate" }),
);
write(
  join(brandDir, "mark-mono.svg"),
  squareMarkSvg({ ink: C.accentLight, accent: C.accentLight, mono: true, title: "Docxcelerate" }),
);
write(
  join(brandDir, "mark-reversed.svg"),
  squareMarkSvg({ ink: C.onAccent, accent: C.onAccent, mono: true, title: "Docxcelerate" }),
);
write(join(brandDir, "lockup-light.svg"), lockupSvg({ ink: C.inkLight, accent: C.accentLight }));
write(join(brandDir, "lockup-dark.svg"), lockupSvg({ ink: C.inkDark, accent: C.accentDark }));

/* --- Favicons ------------------------------------------------------------- */

// The tab icon follows the OS preference, because a mark that is legible on
// Chrome's light strip is not legible on its dark one. Transparent, so it is
// the mark in the tab rather than a tile in the tab.
write(join(publicDir, "favicon.svg"), themedMarkSvg());

// favicon.ico cannot ask about the theme, so it stops trying: reversed out of
// brand green, which holds against a light strip and a dark one alike. The
// opacity steps vanish below about 20px, so this is the monochrome cut — the
// rag of the four lines is what still identifies it that small.
const icoSizes = [16, 32, 48];
const icoPngs = await Promise.all(
  icoSizes.map((size) =>
    rasterise(
      iconSvg({
        size,
        ground: C.accentLight,
        radius: size <= 16 ? 2 : Math.round(size * 0.17),
        inkFraction: 0.74,
        ink: C.onAccent,
        accent: C.onAccent,
        mono: true,
      }),
      size,
    ),
  ),
);
write(join(publicDir, "favicon.ico"), buildIco(icoPngs, icoSizes));

/* --- App icons ------------------------------------------------------------ */

// iOS rounds and shadows this itself, so it goes out square and full-bleed on
// the site's own dark ground.
write(
  join(publicDir, "apple-touch-icon.png"),
  await rasterise(
    iconSvg({
      size: 180,
      ground: C.paperDark,
      radius: 0,
      inkFraction: 0.54,
      ink: C.inkDark,
      accent: C.accentDark,
    }),
    180,
  ),
);

// The installable icons take the light ground: a manifest icon is shown as-is
// on a launcher that may be any colour, and dark-on-white survives that better
// than light-on-dark does.
for (const size of [192, 512]) {
  write(
    join(publicDir, `icon-${size}.png`),
    await rasterise(
      iconSvg({
        size,
        ground: C.paperLight,
        radius: Math.round(size * 0.219),
        inkFraction: 0.54,
        ink: C.inkLight,
        accent: C.accentLight,
      }),
      size,
    ),
  );
}

// Maskable: the platform crops this to whatever shape it likes, so the ground
// bleeds to the edge with no rounding of its own and the ink stays inside the
// 80% safe circle. Half the width does that with room to spare — the ink's
// corner-to-corner span is 0.70 of the icon, against the 0.80 the mask allows —
// while still reading as full rather than as a mark lost in a field of green.
write(
  join(publicDir, "icon-maskable-512.png"),
  await rasterise(
    iconSvg({
      size: 512,
      ground: C.accentLight,
      radius: 0,
      inkFraction: 0.50,
      ink: C.onAccent,
      accent: C.onAccent,
      mono: true,
    }),
    512,
  ),
);

/* --- Share cards ---------------------------------------------------------- */

const { tagline, description } = readCopy();
const command = installCommand();
const [taglineFirst, ...taglineRest] = tagline.split(/(?<=\.)\s+/);

// Open Graph. 1200x630 is the ratio every platform crops least, and Seo.astro
// starts announcing the card the moment this file exists.
//
// The oversized mark behind the composition sits fully inside the frame rather
// than bleeding off a corner. Bled, the four bars lose their rag and read as
// three unrelated slabs — the one thing the mark must never do.
await shoot({
  width: 1200,
  height: 630,
  out: join(publicDir, "og.png"),
  html: cardHtml({
    width: 1200,
    height: 630,
    ground: C.paperDark,
    body: `<div style="position:relative;width:1200px;height:630px;display:flex;flex-direction:column;justify-content:space-between;padding:74px 84px;overflow:hidden">
  <div style="position:absolute;right:60px;bottom:52px;opacity:0.05">${inlineMark({ width: 320, ink: C.inkDark, accent: C.inkDark, mono: true })}</div>
  <div style="position:relative;display:flex;align-items:center;gap:15px">
    ${inlineMark({ width: 34, ink: C.inkDark, accent: C.accentDark })}
    ${wordmark({ size: 33, ink: C.inkDark, accent: C.accentDark })}
  </div>
  <div style="position:relative;display:flex;flex-direction:column;gap:24px">
    <div style="font-family:${SERIF};font-size:66px;font-weight:600;letter-spacing:-0.025em;line-height:1.1;color:${C.inkDark};max-width:900px">${taglineFirst}<br><span style="color:${C.accentDark}">${taglineRest.join(" ")}</span></div>
    <div style="font-family:${SANS};font-size:23px;line-height:1.45;color:${C.mutedDark};max-width:790px">${description}</div>
  </div>
  <div style="position:relative;display:flex;align-items:center;gap:20px">
    <span style="font-family:${MONO};font-size:19px;color:${C.accentDark};background:${C.surfaceDark};border:1px solid ${C.borderDark};border-radius:8px;padding:12px 16px">$ ${command}</span>
    <span style="font-family:${MONO};font-size:17px;color:${C.mutedDark}">MIT · docxcelerate.com</span>
  </div>
</div>`,
  }),
});

// The square crop, for anywhere that wants an avatar-shaped card. Reversed out
// of brand green so it never has to compete with a timeline's own ground.
await shoot({
  width: 1200,
  height: 1200,
  out: join(publicDir, "social-square.png"),
  html: cardHtml({
    width: 1200,
    height: 1200,
    ground: C.accentLight,
    body: `<div style="width:1200px;height:1200px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:80px;padding:140px">
  ${inlineMark({ width: 400, ink: C.onAccent, accent: C.onAccent, mono: true })}
  <div style="display:flex;flex-direction:column;align-items:center;gap:28px">
    <span style="font-family:${SERIF};font-size:104px;font-weight:600;letter-spacing:-0.025em;color:${C.onAccent}">Docxcelerate</span>
    <span style="font-family:${SANS};font-size:44px;line-height:1.4;text-align:center;color:rgba(255,255,255,0.74)">${tagline}</span>
  </div>
</div>`,
  }),
});

// The README's masthead. Wide and short, so it reads as a rule across the top
// of the page rather than as a picture the reader has to scroll past.
await shoot({
  width: 1280,
  height: 320,
  out: join(brandDir, "banner.png"),
  html: cardHtml({
    width: 1280,
    height: 320,
    ground: C.paperDark,
    body: `<div style="position:relative;width:1280px;height:320px;display:flex;align-items:center;gap:40px;padding:0 88px;overflow:hidden">
  <div style="position:absolute;right:60px;bottom:23px;opacity:0.05">${inlineMark({ width: 300, ink: C.inkDark, accent: C.inkDark, mono: true })}</div>
  <div style="position:relative;display:flex;align-items:center;gap:36px">
    ${inlineMark({ width: 104, ink: C.inkDark, accent: C.accentDark })}
    <div style="display:flex;flex-direction:column;gap:12px">
      ${wordmark({ size: 62, ink: C.inkDark, accent: C.accentDark })}
      <span style="font-family:${SANS};font-size:26px;color:${C.mutedDark}">${tagline}</span>
    </div>
  </div>
</div>`,
  }),
});

/* --- Report --------------------------------------------------------------- */

console.log(`brand: wrote ${written.length} files`);
for (const path of written) console.log(`  ${path}`);
