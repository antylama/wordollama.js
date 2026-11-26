# AGENTS.md

Guidance for coding agents working in `wordollama.js`.

## Architecture

- **Desktop Bridge (`src/WordOllama.DesktopBridge`)**: ASP.NET Core (.NET 8) host serving the add-in bundle (`wwwroot`), model providers (`src/WordOllama.Core`), internal Agent runtime (`AgentSession.cs`, `AgentContextWindow.cs`), and external Word MCP server (`WordMcpEndpoint.cs`, `OfficeToolBroker.cs`).
- **Office.js / WPS Add-in (`officejs/apps/addin`)**: React + TypeScript + Vite frontend (`src/officejs-word-adapter.ts`, `src/officejs-tool-registry.ts`, `src/runtime-client.ts`, `src/main.ts`, `src/taskpane/*`, `src/settings/*`).
- **Contracts & MCP (`src/WordOllama.Contracts/BridgeProtocol.cs`, `src/WordOllama.Mcp`)**: Shared DTOs and external MCP client manager.

## Build, Run, and Test Commands

- **TypeScript check:**
  `npm --prefix officejs/apps/addin run build` (`tsc --noEmit`)
- **Add-in smoke tests:**
  - `npm --prefix officejs/apps/addin run test:mock` (`officejs/test-fixtures/officejs-registry-smoke.ts`)
  - `npm --prefix officejs/apps/addin run test:matrix` (`officejs/test-fixtures/officejs-host-matrix.ts`)
  - `npm --prefix officejs/apps/addin run test:golden` (`officejs/test-fixtures/officejs-golden-runner-smoke.ts`)
- **Update frontend UI in place (no Bridge restart needed):**
  `npm run build:ui`
- **Full local / public build (rebuilds `dist-local/WordOllama.DesktopBridge.dll` and `dist-local/wwwroot`):**
  `npm run build:local` or `npm run build:public`
  *(Stop any running `DesktopBridge` process first; otherwise cleaning `dist-local` fails with `EPERM`/`EBUSY` on Windows.)*
- **Run Desktop Bridge:**
  `npm run start:local` or `npm run start:public`
- **.NET build & core smoke test:**
  - `dotnet build src/WordOllama.DesktopBridge/WordOllama.DesktopBridge.csproj -c Release`
  - `dotnet run --project tools/unified-core-smoke/WordOllama.UnifiedCoreSmoke.csproj -c Release`
  *(On Windows, if `DOTNET_ROOT` or `MSBuildSDKsPath` in your shell points to a different SDK than `where.exe dotnet`, clear both env vars before running `dotnet`.)*

## Word / Office.js & MCP Conventions

- **English tool schemas & 1-based indexing:** Keep all tool descriptions in `officejs/apps/addin/src/officejs-tool-registry.ts` inline in English (do not route tool schemas through `i18n.t`). Always document 1-based indices (`paragraph_index`, `table_index`, `comment_index`, `section_index`).
- **Capability gating & test sync:** Gate new tools in `OfficeJsWordAdapter.supportsTool` using `WordApi` (`1.1`–`1.5`) or `WordApiDesktop` (`1.2`–`1.4`) requirement sets. Whenever tools are added or removed, update `officejs/test-fixtures/officejs-registry-smoke.ts` and `officejs/test-fixtures/officejs-host-matrix.ts`.
- **Selection-hash guard:** Tools that mutate or insert at the current selection (`insert_at_cursor`, `insert_field`, `insert_hyperlink`, `insert_footnote`, `format_text`, `insert_page_break`, `format_list`, `add_comment`, `apply_precise_revision`, and `clear_formatting` with `scope: "selection"`) must require `expected_selection_hash`.
- **Tables with merged cells or mixed widths:** Never access `table.columns` or `table.getCell(row, column)` when reading tables — Word throws if cells have varying widths or spans. Iterate `table.rows.items` -> `row.cells.items` instead.
- **Localized style names:** `paragraph.style` and `apply_style` / `edit_style` use Word's localized style names (e.g. `Nagłówek 1` vs `Heading 1`). Use `styleBuiltIn` when targeting built-in styles locale-independently (e.g. `Word.BuiltInStyleName.normal`).
- **Visual page previews (`__wordollamaImage`):** `get_page_preview` exports PDF via `Office.context.document.getFileAsync(Office.FileType.Pdf)` and rasterizes pages lazily with `pdfjs-dist` + `<canvas>`. Payloads with `__wordollamaImage: true` are converted into MCP `image` blocks in `WordMcpEndpoint` and forwarded as `ImageDataUrl` in `AgentSession`.
- **MCP security & multi-window targeting:** `/mcp/word` requires Bearer or route token auth (`WordMcpEndpoint.TryAuthorize`). Connected Word panes share the machine's persisted MCP token and report `documentName` (filename only, never full path), `platform`, `isFocused`, `lastFocusedAt`, and `isMostRecentlyActive` via `/capabilities` and `/office-tools/hosts/{hostId}/activity`.
- **Graceful Bridge shutdown:** Long-lived polling/streaming endpoints (`/office-tools/hosts/{hostId}/requests/next`, `/agent/sessions/{id}/events`) must link `IHostApplicationLifetime.ApplicationStopping` with `HttpContext.RequestAborted` so `Ctrl+C` exits promptly.
