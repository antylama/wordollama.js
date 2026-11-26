import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { crc32 } from "node:zlib";
import { Resvg } from "@resvg/resvg-js";

const scriptDir = dirname(fileURLToPath(import.meta.url));
const addinRoot = resolve(scriptDir, "..");
const repoRoot = resolve(addinRoot, "../..");

async function loadEnvFiles(paths) {
  for (const filePath of paths) {
    try {
      const content = await readFile(filePath, "utf8");
      for (const rawLine of content.split(/\r?\n/)) {
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
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
    }
  }
}

await loadEnvFiles([
  join(addinRoot, ".env.local"),
  join(addinRoot, ".env"),
  join(repoRoot, ".env.local"),
  join(repoRoot, ".env"),
]);

function readFlag(name, fallback) {
  const prefix = `--${name}=`;
  const match = process.argv.find((arg) => arg.startsWith(prefix));
  return match ? match.slice(prefix.length) : fallback;
}

const rawBaseUrl = readFlag("url", process.env.WORDOLLAMA_ADDIN_URL || "");
if (!rawBaseUrl) {
  throw new Error("Missing WORDOLLAMA_ADDIN_URL. Set it in officejs/apps/addin/.env or pass --url=https://...");
}
const baseUri = new URL(rawBaseUrl);
if (baseUri.protocol !== "https:") {
  throw new Error(`Base URL must use https:// (got ${rawBaseUrl})`);
}
const origin = baseUri.origin;
const domain = baseUri.hostname;
const appId = readFlag("id", process.env.WORDOLLAMA_APP_ID || "9f8e6b40-3c1f-4e7b-9a2d-5e1c8a7b4f01");
const developerName = readFlag("developer", process.env.WORDOLLAMA_DEVELOPER_NAME || "WordOllama.JS");
const outDir = resolve(addinRoot, readFlag("out", "dist-manifest"));

const shortVersion = readFlag("version", "1.2.1");
const fourPartVersion = /^\d+\.\d+\.\d+\.\d+$/.test(shortVersion)
  ? shortVersion
  : `${shortVersion}.4`;

const colorSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 192 192">
  <rect width="192" height="192" rx="36" fill="#1d4ed8"/>
  <g fill="none" stroke="#ffffff" stroke-width="11" stroke-linecap="round" stroke-linejoin="round">
    <path d="M96 64V36H68"/>
    <rect width="112" height="84" x="40" y="64" rx="16"/>
    <path d="M24 106h16"/>
    <path d="M152 106h16"/>
    <path d="M118 98v16"/>
    <path d="M74 98v16"/>
  </g>
</svg>`;

const outlineSvg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  <g fill="none" stroke="#ffffff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M16 10V5h-5"/>
    <rect width="20" height="15" x="6" y="10" rx="3"/>
    <path d="M3 17.5h3"/>
    <path d="M26 17.5h3"/>
    <path d="M20 16v3"/>
    <path d="M12 16v3"/>
  </g>
</svg>`;

function renderPng(svg, width) {
  const resvg = new Resvg(svg, { fitTo: { mode: "width", value: width } });
  return resvg.render().asPng();
}

function createZip(entries) {
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  for (const { name, data } of entries) {
    const nameBytes = Buffer.from(name, "utf8");
    const crc = crc32(data) >>> 0;

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0x0800, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(nameBytes.length, 26);
    localHeader.writeUInt16LE(0, 28);

    localChunks.push(localHeader, nameBytes, data);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(nameBytes.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);

    centralChunks.push(centralHeader, nameBytes);
    offset += localHeader.length + nameBytes.length + data.length;
  }

  const centralDir = Buffer.concat(centralChunks);
  const endRecord = Buffer.alloc(22);
  endRecord.writeUInt32LE(0x06054b50, 0);
  endRecord.writeUInt16LE(0, 4);
  endRecord.writeUInt16LE(0, 6);
  endRecord.writeUInt16LE(entries.length, 8);
  endRecord.writeUInt16LE(entries.length, 10);
  endRecord.writeUInt32LE(centralDir.length, 12);
  endRecord.writeUInt32LE(offset, 16);
  endRecord.writeUInt16LE(0, 20);

  return Buffer.concat([...localChunks, centralDir, endRecord]);
}

const unifiedManifest = {
  $schema: "https://developer.microsoft.com/json-schemas/teams/v1.25/MicrosoftTeams.schema.json",
  manifestVersion: "1.25",
  id: appId,
  version: shortVersion,
  developer: {
    name: developerName,
    websiteUrl: origin,
    privacyUrl: `${origin}/privacy`,
    termsOfUseUrl: `${origin}/terms`,
  },
  icons: {
    color: "color-192.png",
    outline: "outline-32.png",
  },
  name: {
    short: "WordOllama.JS",
    full: "WordOllama.JS AI Assistant for Word",
  },
  description: {
    short: "AI assistant and MCP bridge for Microsoft Word.",
    full: "WordOllama.JS brings an AI Agent, custom instructions, skills, and Word MCP integration directly into Microsoft Word.",
  },
  accentColor: "#1d4ed8",
  validDomains: [domain],
  authorization: {
    permissions: {
      resourceSpecific: [
        {
          name: "Document.ReadWrite.User",
          type: "Delegated",
        },
      ],
    },
  },
  extensions: [
    {
      requirements: {
        scopes: ["document"],
        capabilities: [
          {
            name: "WordApi",
            minVersion: "1.1",
          },
        ],
      },
      runtimes: [
        {
          id: "WordOllamaRuntime",
          type: "general",
          code: {
            page: `${origin}/index.html`,
          },
          lifetime: "long",
          requirements: {
            capabilities: [
              {
                name: "SharedRuntime",
                minVersion: "1.1",
              },
            ],
            formFactors: ["desktop"],
          },
          actions: [
            {
              id: "OpenWordOllamaTaskpane",
              type: "openPage",
              displayName: "WordOllama.JS",
              pinnable: false,
              view: "WordOllamaTaskpane",
              multiselect: false,
              supportsNoItemContext: false,
            },
          ],
        },
      ],
      ribbons: [
        {
          contexts: ["default"],
          tabs: [
            {
              builtInTabId: "TabHome",
              groups: [
                {
                  id: "WordOllama.JS.HomeGroup",
                  label: "WordOllama.JS",
                  icons: [
                    { size: 16, url: `${origin}/assets/ribbon/agent-16.png` },
                    { size: 32, url: `${origin}/assets/ribbon/agent-32.png` },
                    { size: 80, url: `${origin}/assets/ribbon/agent-80.png` },
                  ],
                  controls: [
                    {
                      id: "WordOllama.JS.OpenTaskpane",
                      type: "button",
                      label: "WordOllama.JS",
                      supertip: {
                        title: "WordOllama.JS",
                        description: "Open WordOllama.JS assistant in Word.",
                      },
                      icons: [
                        { size: 16, url: `${origin}/assets/ribbon/agent-16.png` },
                        { size: 32, url: `${origin}/assets/ribbon/agent-32.png` },
                        { size: 80, url: `${origin}/assets/ribbon/agent-80.png` },
                      ],
                      actionId: "OpenWordOllamaTaskpane",
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
  ],
};

await mkdir(outDir, { recursive: true });

const colorPng = renderPng(colorSvg, 192);
const outlinePng = renderPng(outlineSvg, 32);
const manifestJsonBytes = Buffer.from(`${JSON.stringify(unifiedManifest, null, 2)}\n`, "utf8");

const sourceXml = await readFile(join(addinRoot, "manifest.xml"), "utf8");
const prodXml = sourceXml
  .replaceAll("https://localhost:3000", origin)
  .replace("<Id>4d2a7c5e-2d2a-4a1a-8b72-6a1cf4f7b701</Id>", `<Id>${appId}</Id>`)
  .replace(/<Version>[^<]+<\/Version>/, `<Version>${fourPartVersion}</Version>`);

const zipBytes = createZip([
  { name: "manifest.json", data: manifestJsonBytes },
  { name: "color-192.png", data: colorPng },
  { name: "outline-32.png", data: outlinePng },
]);

await Promise.all([
  writeFile(join(outDir, "manifest.json"), manifestJsonBytes),
  writeFile(join(outDir, "color-192.png"), colorPng),
  writeFile(join(outDir, "outline-32.png"), outlinePng),
  writeFile(join(outDir, "manifest.xml"), prodXml, "utf8"),
  writeFile(join(outDir, "wordollama-m365-manifest.zip"), zipBytes),
]);

console.log(`Generated M365 manifest package for ${origin} (id: ${appId}):`);
console.log(`  - ${join(outDir, "wordollama-m365-manifest.zip")}`);
console.log(`  - ${join(outDir, "manifest.json")}`);
console.log(`  - ${join(outDir, "manifest.xml")}`);
