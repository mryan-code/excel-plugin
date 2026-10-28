# excel-plugin — Market Pulse (learning project)

A deliberately small **Excel task-pane add-in** built with TypeScript and the
Office JS API. It exists to learn the platform's core patterns by doing:

- **Insert sample data** — writes a small table starting at the selected cell
  using `Excel.run` with **one batched `context.sync()`**.
- **Live updates** — a 2-second timer random-walks the prices and writes the
  whole block back per tick (throttled, batched — never one API call per cell),
  with per-tick error handling so a bad tick doesn't kill the stream.

## Run it (Mac)

Prereqs: Node.js, Excel for Mac.

```bash
npm install
npm start        # starts the HTTPS dev server (port 3000) and sideloads the
                 # add-in into Excel via office-addin-debugging
```

The first run installs a local dev certificate (office-addin-dev-certs) and
asks you to trust it — that's what lets Excel load the add-in from
`https://localhost:3000`. Excel opens with a **Market Pulse** group on the
Home tab; click **Open Market Pulse**.

Stop everything with `npm stop`.

Manual alternative: `npm run dev-server`, then copy `manifest.xml` into
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` and insert the
add-in from Excel's Add-ins menu.

## Layout

- `manifest.xml` — the add-in manifest (identity, permissions, ribbon button)
- `src/taskpane/` — task pane UI + logic (`taskpane.ts` is the interesting file)
- `src/commands/` — FunctionFile required by the manifest (currently a stub)

## Concepts this exercises

- Host-bridged API: web code talks to Excel through a request context, not directly
- Batching + `context.sync()` round-trips (the #1 performance lever in Office JS)
- Requirement sets / capability detection across Excel on Mac, Windows, and web
- Sideloading, manifests, and local HTTPS dev certs — the deployment basics
