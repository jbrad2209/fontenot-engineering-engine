import { mkdirSync, copyFileSync, rmSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const stagingRoot = join(root, ".grok", "engine-kit-staging");
const staging = join(stagingRoot, "fontenot-engineering-engine");
const outDir = join(root, "public", "downloads");
const zipPath = join(outDir, "fontenot-engineering-engine.zip");
const rolldown = join(root, "node_modules", ".bin", "rolldown");

rmSync(stagingRoot, { recursive: true, force: true });
mkdirSync(join(staging, "samples"), { recursive: true });
mkdirSync(outDir, { recursive: true });

const bundle = spawnSync(
  rolldown,
  [
    join(root, "src", "lib", "engine", "cli-entry.ts"),
    "--platform",
    "node",
    "--format",
    "cjs",
    "--external",
    "exceljs",
    "--file",
    join(staging, "report.cjs"),
    "--banner",
    "#!/usr/bin/env node",
  ],
  { encoding: "utf8" },
);
if (bundle.status !== 0) {
  process.stderr.write((bundle.stdout || "") + (bundle.stderr || ""));
  process.exit(bundle.status ?? 1);
}

copyFileSync(join(root, "engine-kit", "README.md"), join(staging, "README.md"));
copyFileSync(join(root, "engine-kit", "package.json"), join(staging, "package.json"));
copyFileSync(
  join(root, "public", "samples", "average_response_time.xlsx"),
  join(staging, "samples", "average_response_time.xlsx"),
);
copyFileSync(
  join(root, "public", "samples", "all_work_orders.xlsx"),
  join(staging, "samples", "all_work_orders.xlsx"),
);

if (existsSync(zipPath)) rmSync(zipPath);

const py = spawnSync(
  "python3",
  [
    "-c",
    `
import zipfile, os, sys
root = sys.argv[1]
zip_path = sys.argv[2]
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
    for dirpath, _, files in os.walk(root):
        for name in files:
            full = os.path.join(dirpath, name)
            arc = os.path.relpath(full, root)
            z.write(full, arc)
print(zip_path)
`,
    stagingRoot,
    zipPath,
  ],
  { encoding: "utf8" },
);
if (py.status !== 0) {
  process.stderr.write((py.stdout || "") + (py.stderr || "zip failed\n"));
  process.exit(py.status ?? 1);
}

process.stdout.write(`packed ${zipPath}\n`);
