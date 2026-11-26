import { execSync } from "node:child_process";
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(scriptDir, "..");
const addinRoot = join(repoRoot, "officejs", "apps", "addin");
const outDir = join(repoRoot, "dist-local");

function loadEnvFiles(paths) {
  for (const filePath of paths) {
    if (!existsSync(filePath)) continue;
    for (const rawLine of readFileSync(filePath, "utf8").split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) continue;
      const eq = line.indexOf("=");
      if (eq <= 0) continue;
      const key = line.slice(0, eq).trim();
      let value = line.slice(eq + 1).trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }
      if (!(key in process.env)) process.env[key] = value;
    }
  }
}

loadEnvFiles([
  join(addinRoot, ".env.local"),
  join(addinRoot, ".env"),
  join(repoRoot, ".env.local"),
  join(repoRoot, ".env"),
]);

function run(command, args, cwd = repoRoot) {
  execSync([command, ...args].join(" "), { cwd, stdio: "inherit" });
}

console.log("[1/3] Building Office.js frontend bundle...");
run("npm", ["--prefix", "officejs/apps/addin", "run", "bundle"]);

console.log("[2/3] Publishing Desktop Bridge...");
try {
  rmSync(outDir, { recursive: true, force: true });
} catch (error) {
  if (error?.code === "EPERM" || error?.code === "EBUSY") {
    console.error(`\nCannot clean ${outDir} (${error.code}) — stop the running Desktop Bridge (Ctrl+C on start:local / start:public) first.`);
    process.exit(1);
  }
  throw error;
}
mkdirSync(outDir, { recursive: true });
run("dotnet", [
  "publish",
  "src/WordOllama.DesktopBridge/WordOllama.DesktopBridge.csproj",
  "-c",
  "Release",
  "-o",
  outDir,
]);

console.log("[3/3] Copying frontend bundle and assets into wwwroot...");
const wwwroot = join(outDir, "wwwroot");
mkdirSync(wwwroot, { recursive: true });
cpSync(join(addinRoot, "dist"), wwwroot, { recursive: true });
if (existsSync(join(addinRoot, "assets"))) {
  cpSync(join(addinRoot, "assets"), join(wwwroot, "assets"), { recursive: true });
  const icon32 = join(addinRoot, "assets", "icon-32.png");
  if (existsSync(icon32)) {
    cpSync(icon32, join(wwwroot, "favicon.ico"));
  }
}

const listenPublic =
  process.argv.includes("public") ||
  process.argv.includes("--public") ||
  process.env.WORDOLLAMA_ALLOW_REMOTE === "1";
const listenUrl = (process.env.WORDOLLAMA_LISTEN_URL || (listenPublic ? "http://0.0.0.0:37421" : "")).trim();
const addinOrigin = (process.env.WORDOLLAMA_ADDIN_URL || "").trim().replace(/\/+$/, "");
const appSettingsPath = join(outDir, "appsettings.json");
if (existsSync(appSettingsPath)) {
  const settings = JSON.parse(readFileSync(appSettingsPath, "utf8"));
  settings.Bridge ??= {};
  if (listenUrl) {
    settings.Bridge.Urls = listenUrl;
  }
  if (listenPublic || addinOrigin) {
    settings.Bridge.AllowRemoteAccess = true;
  }
  if (addinOrigin) {
    const allowed = new Set(Array.isArray(settings.Bridge.AllowedOrigins) ? settings.Bridge.AllowedOrigins : []);
    allowed.add(addinOrigin);
    settings.Bridge.AllowedOrigins = [...allowed];
  }
  writeFileSync(appSettingsPath, `${JSON.stringify(settings, null, 2)}\n`, "utf8");
  console.log(
    `Configured dist-local/appsettings.json (Urls=${settings.Bridge.Urls}, AllowRemoteAccess=${Boolean(settings.Bridge.AllowRemoteAccess)})`,
  );
}

console.log(`\nDone. Run "npm run start:local" to serve both UI and Bridge on http://127.0.0.1:37421`);
