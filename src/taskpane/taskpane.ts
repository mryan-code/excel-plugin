import "./taskpane.css";

/*
 * Market Pulse — a deliberately small Excel add-in.
 *
 * The two things it demonstrates are the two patterns every real Excel
 * add-in lives on:
 *   1. Excel.run + batched operations: queue up everything, one context.sync().
 *   2. Streaming into cells: a timer produces new values locally, and each
 *      tick writes the whole block in a single batched call — never one
 *      API call per cell (that is what kills add-ins in production).
 */

interface Row {
  ticker: string;
  price: number;
}

const seed: Row[] = [
  { ticker: "ACME", price: 182.4 },
  { ticker: "GLOBEX", price: 94.15 },
  { ticker: "INITECH", price: 41.02 },
  { ticker: "UMBRELLA", price: 210.77 },
];

let tableAddress: string | null = null;
let timer: number | null = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.");
    return;
  }
  document.getElementById("insert")!.addEventListener("click", insertSampleData);
  document.getElementById("toggle-live")!.addEventListener("click", toggleLive);
  setStatus("Ready. Select a cell, then insert the sample data.");
});

function setStatus(text: string) {
  document.getElementById("status")!.textContent = text;
}

async function insertSampleData() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const start = context.workbook.getSelectedRange();
      start.load("rowIndex,columnIndex");
      await context.sync();

      const rows: (string | number)[][] = [
        ["Ticker", "Price", "Change %"],
        ...seed.map((r) => [r.ticker, r.price, 0]),
      ];
      const target = sheet.getRangeByIndexes(start.rowIndex, start.columnIndex, rows.length, 3);
      target.values = rows;

      const header = sheet.getRangeByIndexes(start.rowIndex, start.columnIndex, 1, 3);
      header.format.font.bold = true;
      target.format.autofitColumns();

      target.load("address");
      await context.sync(); // <-- the one round-trip that does all of the above

      tableAddress = target.address;
      (document.getElementById("toggle-live") as HTMLButtonElement).disabled = false;
      setStatus(`Inserted at ${tableAddress}. Live updates are ready to start.`);
    });
  } catch (error) {
    setStatus(`Insert failed: ${String(error)}`);
  }
}

function toggleLive() {
  const button = document.getElementById("toggle-live") as HTMLButtonElement;
  if (timer !== null) {
    window.clearInterval(timer);
    timer = null;
    button.textContent = "Start live updates";
    setStatus("Live updates stopped.");
    return;
  }
  button.textContent = "Stop live updates";
  setStatus("Streaming… watch the Price / Change % columns.");
  timer = window.setInterval(tick, 2000);
}

async function tick() {
  if (!tableAddress) return;
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(tableAddress!);
      range.load("values");
      await context.sync();

      // Random-walk the prices locally, then write the whole block back
      // in a single batched assignment.
      const values = range.values;
      for (let i = 1; i < values.length; i++) {
        const prev = Number(values[i][1]);
        const next = Math.max(0.01, prev * (1 + (Math.random() - 0.5) * 0.02));
        values[i][1] = Math.round(next * 100) / 100;
        values[i][2] = Math.round(((next - prev) / prev) * 10000) / 100;
      }
      range.values = values;
      await context.sync();
    });
  } catch (error) {
    // A streaming add-in must survive a bad tick (sheet closed, range
    // deleted) instead of dying silently.
    setStatus(`Tick skipped: ${String(error)}`);
  }
}
