import * as fs from "fs";
import * as os from "os";
import * as path from "path";

const WORKBOOK_NAME = "GrafanaErrorLog.xlsx";
const SUBDIR = "grafanalog";

function resolveOneDriveDir(): string | null {
  const fromEnv = process.env.ONEDRIVE_GRAFANA_DIR;
  if (fromEnv) return fromEnv;

  const cloudStorage = path.join(os.homedir(), "Library", "CloudStorage");
  if (fs.existsSync(cloudStorage)) {
    const match = fs
      .readdirSync(cloudStorage)
      .find((name) => name.startsWith("OneDrive"));
    if (match) return path.join(cloudStorage, match, SUBDIR);
  }

  const legacyOneDrive = path.join(os.homedir(), "OneDrive", SUBDIR);
  if (fs.existsSync(path.join(os.homedir(), "OneDrive"))) {
    return legacyOneDrive;
  }

  return null;
}

export function getOneDriveWorkbookPath(): string | null {
  const dir = resolveOneDriveDir();
  return dir ? path.join(dir, WORKBOOK_NAME) : null;
}

export function loadWorkbookFromOneDrive(localPath: string): void {
  const oneDrivePath = getOneDriveWorkbookPath();
  if (!oneDrivePath || fs.existsSync(localPath) || !fs.existsSync(oneDrivePath)) {
    return;
  }

  try {
    fs.copyFileSync(oneDrivePath, localPath);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.warn(`Could not load workbook from OneDrive: ${message}`);
  }
}

export function syncWorkbookToOneDrive(localPath: string): void {
  const oneDriveDir = resolveOneDriveDir();
  if (!oneDriveDir) {
    console.log("OneDrive sync skipped: no OneDrive folder found");
    return;
  }

  const oneDrivePath = path.join(oneDriveDir, WORKBOOK_NAME);

  try {
    if (!fs.existsSync(oneDriveDir)) {
      fs.mkdirSync(oneDriveDir, { recursive: true });
    }
    fs.copyFileSync(localPath, oneDrivePath);
    console.log(`Synced to OneDrive: ${oneDrivePath}`);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.warn(`OneDrive sync skipped: ${message}`);
  }
}
