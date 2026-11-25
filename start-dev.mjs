import { spawn, spawnSync } from "node:child_process";

const isWin = process.platform === "win32";

function prefixStream(stream, label, color) {
  let buffer = "";
  stream.on("data", (chunk) => {
    buffer += chunk.toString();
    const lines = buffer.split(/\r?\n/);
    buffer = lines.pop() ?? "";
    for (const line of lines) {
      if (line.length > 0) {
        process.stdout.write(`${color}[${label}]\x1b[0m ${line}\n`);
      }
    }
  });
  stream.on("end", () => {
    if (buffer.length > 0) {
      process.stdout.write(`${color}[${label}]\x1b[0m ${buffer}\n`);
    }
  });
}

const children = [];
let shuttingDown = false;

function stopAll(exitCode = 0) {
  if (shuttingDown) return;
  shuttingDown = true;
  for (const child of children) {
    if (child.pid && !child.killed) {
      if (isWin) {
        spawnSync("taskkill", ["/PID", String(child.pid), "/T", "/F"], { stdio: "ignore" });
      } else {
        child.kill("SIGTERM");
      }
    }
  }
  process.exit(exitCode);
}

function launch(label, color, command, args) {
  const child = spawn(command, args, {
    stdio: ["ignore", "pipe", "pipe"],
  });
  children.push(child);
  prefixStream(child.stdout, label, color);
  prefixStream(child.stderr, label, color);
  child.on("exit", (code) => {
    if (!shuttingDown) {
      process.stdout.write(`${color}[${label}]\x1b[0m exited with code ${code ?? 0}\n`);
      stopAll(code ?? 0);
    }
  });
  return child;
}

process.on("SIGINT", () => stopAll(0));
process.on("SIGTERM", () => stopAll(0));

launch("bridge", "\x1b[36m", "dotnet", [
  "run",
  "--project",
  "./src/WordOllama.DesktopBridge/WordOllama.DesktopBridge.csproj",
]);

launch("addin", "\x1b[35m", isWin ? "cmd.exe" : "npm", [
  ...(isWin ? ["/c", "npm.cmd"] : []),
  "--prefix",
  "officejs/apps/addin",
  "run",
  "dev",
]);
