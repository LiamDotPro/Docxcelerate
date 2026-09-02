/**
 * The small filesystem layer the scaffolder and the registry installer share.
 *
 * Both write somebody else's project into place, and both had grown their own
 * copy of the same six helpers. The copies had already drifted: one `parentPath`
 * normalised separators and the other called `dirname`, which answers with
 * backslashes on Windows for a path that arrived with them.
 *
 * `node:fs/promises` is imported where it is used rather than at the top, so
 * importing a module that only *might* touch a disk does not pull the whole of
 * `node:fs` into a graph that never will.
 *
 * @module
 */

/**
 * The directory a path sits in.
 *
 * Always answers with forward slashes, whatever came in. Paths here are joined
 * with `/` and handed to `mkdir`, so a mixed-separator answer is one that only
 * breaks on one platform.
 *
 * @param path The path to take the parent of.
 * @returns The parent directory, or `.` when there is no separator.
 */
export function parentPath(path: string): string {
  const normalized = path.replaceAll("\\", "/");
  const index = normalized.lastIndexOf("/");

  return index === -1 ? "." : normalized.slice(0, index);
}

/**
 * Creates a directory and everything above it.
 *
 * An empty path and `.` are the current directory, which already exists; both
 * arrive here from `parentPath` on a bare filename.
 *
 * @param path The directory to create.
 */
export async function ensureDirectory(path: string): Promise<void> {
  if (path === "" || path === ".") {
    return;
  }

  const { mkdir } = await import("node:fs/promises");
  await mkdir(path, { recursive: true });
}

/**
 * Whether anything is at a path.
 *
 * @param path The path to look at.
 * @returns `true` when something is there.
 * @throws If the path cannot be read for any reason other than not being there.
 */
export async function exists(path: string): Promise<boolean> {
  const { stat } = await import("node:fs/promises");

  try {
    await stat(path);
    return true;
  } catch (error) {
    if (isNotFoundError(error)) {
      return false;
    }

    throw error;
  }
}

/**
 * Reads a file as UTF-8 text.
 *
 * @param path The file to read.
 * @returns Its contents.
 */
export async function readTextFile(path: string): Promise<string> {
  const { readFile } = await import("node:fs/promises");

  return await readFile(path, "utf8");
}

/**
 * Writes a file as UTF-8 text, replacing whatever was there.
 *
 * @param path The file to write.
 * @param contents What to write.
 */
export async function writeTextFile(path: string, contents: string): Promise<void> {
  const { writeFile } = await import("node:fs/promises");
  await writeFile(path, contents, "utf8");
}

/**
 * The names directly inside a directory.
 *
 * @param path The directory to read.
 * @returns The entry names, in whatever order the platform gives them.
 * @throws If the directory is not there.
 */
export async function readDirectoryNames(path: string): Promise<string[]> {
  const { readdir } = await import("node:fs/promises");

  return await readdir(path);
}

/**
 * Whether an error is a path that was not there.
 *
 * @param error The error to test.
 * @returns `true` when it is `ENOENT`.
 */
export function isNotFoundError(error: unknown): boolean {
  return Boolean(
    error && typeof error === "object" && "code" in error && error.code === "ENOENT",
  );
}
