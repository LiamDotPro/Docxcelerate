/**
 * Reading what an image node points at.
 *
 * A document model is JSON that travels — to an engine, into an artifact, over
 * a wire — so the only picture a renderer can be sure of is one carried in the
 * model itself. That is what a `data:` URI is, and why it is the form both
 * renderers understand fully.
 *
 * A path or an http URL still renders on screen, where the browser can fetch
 * it. It cannot be packed into a Word file, because packing means embedding the
 * bytes and this module has no filesystem and does no fetching. Saying that
 * plainly is better than a renderer quietly reaching for a file that will not
 * be there when the document is written somewhere else.
 *
 * @module
 */

/** What an image node's path turned out to be. */
export type ImageSource =
  | {
    /** Bytes carried in the model itself. */
    readonly kind: "data";
    /** The media type, as written in the URI. */
    readonly mediaType: string;
    /** The decoded bytes. */
    readonly bytes: Uint8Array;
    /** The URI as written, for a renderer that can use it directly. */
    readonly uri: string;
  }
  | {
    /** Something a browser can fetch, but a packer cannot. */
    readonly kind: "url";
    /** The URL as written. */
    readonly uri: string;
  }
  | {
    /** Nothing to draw: no path, or one this cannot read. */
    readonly kind: "none";
  };

/** The raster types Word will embed without a fallback. */
export type RasterType = "png" | "jpg" | "gif" | "bmp";

const rasters: Record<string, RasterType> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/jpg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
};

/**
 * Works out what an image node points at.
 *
 * @param path The node's path, as the model carries it.
 * @returns What it is, and the bytes when it carries them.
 */
export function imageSourceOf(path: string | undefined): ImageSource {
  if (path === undefined || path === "") {
    return { kind: "none" };
  }

  if (!path.startsWith("data:")) {
    return { kind: "url", uri: path };
  }

  const comma = path.indexOf(",");

  if (comma === -1) {
    return { kind: "none" };
  }

  const head = path.slice(5, comma);
  const body = path.slice(comma + 1);
  // A data URI's head is a media type followed by any number of parameters,
  // of which `base64` is one. Reading the whole head as the type leaves the
  // parameters stuck to it — `image/svg+xml;utf8` is not a media type anything
  // recognises, so a perfectly good picture came back as nothing to draw.
  const parameters = head.split(";").map((part) => part.trim());
  const base64 = parameters.includes("base64");
  const mediaType = parameters[0] || "text/plain";

  try {
    const bytes = base64
      ? Uint8Array.from(atob(body), (character) => character.charCodeAt(0))
      : new TextEncoder().encode(decodeURIComponent(body));

    return { kind: "data", mediaType, bytes, uri: path };
  } catch {
    // A URI that will not decode is not a picture. Treated as nothing to draw,
    // so the node falls back to its placeholder rather than the build failing
    // on a string somebody pasted in wrong.
    return { kind: "none" };
  }
}

/**
 * The Word image type for a media type, when it is one Word takes directly.
 *
 * @param mediaType The media type from the URI.
 * @returns The type name, or `undefined` when Word needs more than the bytes.
 */
export function rasterTypeOf(mediaType: string): RasterType | undefined {
  return rasters[mediaType.toLowerCase()];
}

/** Whether a media type is SVG, which Word takes only with a raster fallback. */
export function isSvg(mediaType: string): boolean {
  return mediaType.toLowerCase() === "image/svg+xml";
}
