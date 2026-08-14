import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { execFile } from "child_process";

export function getDownloadsDir(): string {
  const dir = path.join(os.homedir(), "Downloads");
  try {
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  } catch {
    return os.tmpdir();
  }
}

// Opens a local file with the OS-registered default application (e.g. double-click behavior).
// Uses execFile (no shell) so filePath cannot be used for command injection.
export function openWithDefaultApp(filePath: string) {
  const platform = process.platform;
  if (platform === "win32") {
    // explorer.exe opens a file with its registered default application, no shell involved.
    execFile("explorer.exe", [filePath], () => { /* best-effort */ });
  } else if (platform === "darwin") {
    execFile("open", [filePath], () => { /* best-effort */ });
  } else {
    execFile("xdg-open", [filePath], () => { /* best-effort */ });
  }
}

export function uniqueFilePath(dir: string, fileName: string): string {
  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext);
  let candidate = path.join(dir, fileName);
  let n = 1;
  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${base} (${n})${ext}`);
    n++;
  }
  return candidate;
}

export function generateToken(): string {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}
