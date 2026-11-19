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

/**
 * The size a raster says it is, in pixels, read from its own header.
 *
 * PNG, JPEG, GIF and BMP all state their pixel size in the first bytes, so the
 * proportions a picture should keep can be known without decoding it or adding
 * a dependency. The bytes are sniffed by their signature, not by the media
 * type the URI claims, so a JPEG labelled `image/png` still reads.
 *
 * @param bytes The image's bytes, as {@linkcode imageSourceOf} decoded them.
 * @returns The width and height in pixels, or `undefined` when the bytes are
 * not one of those four formats, are cut short, or state a zero side.
 */
export function pixelSizeOf(bytes: Uint8Array): { width: number; height: number } | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const has = (end: number) => end <= bytes.byteLength;
  const startsWith = (...signature: number[]) =>
    has(signature.length) && signature.every((byte, at) => bytes[at] === byte);
  let size: { width: number; height: number } | undefined;

  if (startsWith(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a) && has(24)) {
    size = { width: view.getUint32(16), height: view.getUint32(20) };
  } else if (startsWith(0x47, 0x49, 0x46) && has(10)) {
    size = { width: view.getUint16(6, true), height: view.getUint16(8, true) };
  } else if (startsWith(0x42, 0x4d) && has(26)) {
    // A negative height is rows stored top-down, not a negative size.
    size = { width: view.getInt32(18, true), height: Math.abs(view.getInt32(22, true)) };
  } else if (startsWith(0xff, 0xd8)) {
    size = jpegSizeOf(view);
  }

  return size && size.width > 0 && size.height > 0 ? size : undefined;
}

/** Walks a JPEG's segments to its start-of-frame, which is where the size is. */
function jpegSizeOf(view: DataView): { width: number; height: number } | undefined {
  let at = 2;

  while (at + 4 <= view.byteLength && view.getUint8(at) === 0xff) {
    const marker = view.getUint8(at + 1);

    if (marker === 0xff) {
      at += 1; // fill byte
    } else if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      at += 2; // no length
    } else if (marker === 0xd9 || marker === 0xda) {
      return undefined; // the frame header never came before the image data
    } else if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return at + 9 <= view.byteLength
        ? { width: view.getUint16(at + 7), height: view.getUint16(at + 5) }
        : undefined;
    } else {
      at += 2 + view.getUint16(at + 2);
    }
  }

  return undefined;
}
