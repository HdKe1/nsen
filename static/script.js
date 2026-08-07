const form = document.getElementById("report-form");
const endDateInput = document.getElementById("end-date");
const nDaysInput = document.getElementById("n-days");
const symbolsTextarea = document.getElementById("symbols");
const symbolCountEl = document.getElementById("symbol-count");
const runBtn = document.getElementById("run-btn");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");

const upTableBody = document.querySelector("#up-table tbody");
const downTableBody = document.querySelector("#down-table tbody");
const upCountEl = document.getElementById("up-count");
const downCountEl = document.getElementById("down-count");
const upEmptyEl = document.getElementById("up-empty");
const downEmptyEl = document.getElementById("down-empty");
const downloadBtn = document.getElementById("download-btn");

let lastRequestBody = null;
let lastReportData = null;
let symbolsCustomized = false;

// Default end date = today
endDateInput.value = new Date().toISOString().slice(0, 10);

function loadDefaultSymbols(endDate) {
  const url = endDate
    ? `/api/default-symbols?end_date=${encodeURIComponent(endDate)}`
    : "/api/default-symbols";
  return fetch(url)
    .then((r) => r.json())
    .then((data) => {
      symbolsTextarea.value = data.symbols.join("\n");
      symbolCountEl.textContent = data.symbols.length;
    })
    .catch(() => {
      symbolCountEl.textContent = "0";
    });
}

// Preload default symbol list into the (collapsed) textarea, filtered to
// stocks already listed as of today's default date
loadDefaultSymbols(endDateInput.value);

// If the person hasn't customized the symbol list, automatically refresh
// it whenever the final date changes -- so a date from a decade ago
// won't include companies that hadn't listed yet.
endDateInput.addEventListener("change", () => {
  if (!symbolsCustomized) {
    loadDefaultSymbols(endDateInput.value);
  }
});

symbolsTextarea.addEventListener("input", () => {
  symbolsCustomized = true;
  const count = symbolsTextarea.value
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean).length;
  symbolCountEl.textContent = count;
});

function fmtPct(value) {
  if (value === null || value === undefined) return "—";
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(2)}%`;
}

function pctClass(value) {
  return value >= 0 ? "pos" : "neg";
}

function renderRow(stock) {
  const tr = document.createElement("tr");

  const symbolTd = document.createElement("td");
  symbolTd.className = "symbol-cell";
  symbolTd.textContent = stock.symbol;
  tr.appendChild(symbolTd);

  const daysTd = document.createElement("td");
  stock.days.forEach((d) => {
    const chip = document.createElement("span");
    chip.className = `day-chip ${pctClass(d.pct_change)}`;
    chip.textContent = fmtPct(d.pct_change);
    chip.title = `${d.label} · Volume ${d.volume.toLocaleString("en-IN")}`;
    daysTd.appendChild(chip);
  });
  tr.appendChild(daysTd);

  const netTd = document.createElement("td");
  netTd.className = `net-cell ${pctClass(stock.net_pct_change)}`;
  netTd.textContent = fmtPct(stock.net_pct_change);
  tr.appendChild(netTd);

  const nextTd = document.createElement("td");
  nextTd.className = "next-cell";
  if (stock.next_day_pct_change !== null && stock.next_day_pct_change !== undefined) {
    const pctSpan = document.createElement("span");
    pctSpan.className = `pct ${pctClass(stock.next_day_pct_change)}`;
    pctSpan.textContent = fmtPct(stock.next_day_pct_change);
    pctSpan.title = stock.next_day_volume !== null
      ? `Volume ${stock.next_day_volume.toLocaleString("en-IN")}`
      : "";
    nextTd.appendChild(pctSpan);
    const labelSpan = document.createElement("span");
    labelSpan.className = "label";
    labelSpan.textContent = stock.next_day_label || "";
    nextTd.appendChild(labelSpan);
  } else {
    nextTd.textContent = "—";
  }
  tr.appendChild(nextTd);

  const verdictTd = document.createElement("td");
  verdictTd.className = "verdict-cell" + (stock.verdict.includes("broken") || stock.verdict.includes("Broken") ? " broken" : "");
  verdictTd.textContent = stock.verdict;
  tr.appendChild(verdictTd);

  return tr;
}

function setStatus(message, isError = false) {
  statusEl.hidden = false;
  statusEl.className = "status" + (isError ? " error" : "");
  statusEl.innerHTML = isError
    ? message
    : `<span class="tick"></span>${message}`;
}

function clearStatus() {
  statusEl.hidden = true;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  const endDate = endDateInput.value;
  const nDays = parseInt(nDaysInput.value, 10);
  const symbols = symbolsTextarea.value
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean);

  if (!endDate || !nDays) return;

  runBtn.disabled = true;
  resultsEl.hidden = true;
  setStatus(`Checking ${symbols.length || "default"} symbols for a ${nDays}-day streak through ${endDate}…`);

  const requestBody = {
    end_date: endDate,
    n_days: nDays,
    symbols: symbols.length ? symbols : null,
  };

  try {
    const res = await fetch("/api/report", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(requestBody),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `Request failed (${res.status})`);
    }

    const data = await res.json();

    upTableBody.innerHTML = "";
    downTableBody.innerHTML = "";

    data.up_streaks.forEach((s) => upTableBody.appendChild(renderRow(s)));
    data.down_streaks.forEach((s) => downTableBody.appendChild(renderRow(s)));

    upCountEl.textContent = data.up_streaks.length;
    downCountEl.textContent = data.down_streaks.length;
    upEmptyEl.hidden = data.up_streaks.length > 0;
    downEmptyEl.hidden = data.down_streaks.length > 0;

    resultsEl.hidden = false;
    clearStatus();
    lastRequestBody = requestBody;
    lastReportData = data;
  } catch (err) {
    setStatus(`Could not complete the run — ${err.message}`, true);
  } finally {
    runBtn.disabled = false;
  }
});

function buildStockRow(stock, nDays) {
  const row = [stock.symbol, `${nDays}-Day ${stock.direction} Streak`, stock.net_pct_change];
  stock.days.forEach((d) => {
    row.push(d.pct_change);
    row.push(d.volume);
  });
  row.push(stock.next_day_pct_change ?? "N/A");
  row.push(stock.next_day_volume ?? "N/A");
  row.push(stock.verdict);
  return row;
}

function addStreakSheet(workbook, sheetName, stocks, nDays) {
  const sheet = workbook.addWorksheet(sheetName);

  const headers = ["Symbol", "Streak Type", "Net % Change"];
  for (let i = 1; i <= nDays; i++) {
    headers.push(`Day ${i} % Change`, `Day ${i} Volume`);
  }
  headers.push("Next Day % Change", "Next Day Volume", "Verdict");

  const headerRow = sheet.addRow(headers);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
  });

  const upFill = "FFC6EFCE";
  const downFill = "FFFFC7CE";
  const brokenFill = "FFFFEB9C";

  stocks.forEach((stock) => {
    const dataRow = sheet.addRow(buildStockRow(stock, nDays));
    const rowFill = stock.direction === "Up" ? upFill : downFill;
    dataRow.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowFill } };
    });
    if (stock.verdict && stock.verdict.toLowerCase().includes("broken")) {
      const verdictCell = dataRow.getCell(headers.length);
      verdictCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: brokenFill } };
      verdictCell.font = { bold: true };
    }
  });

  if (stocks.length === 0) {
    sheet.getCell("A2").value = "No qualifying stocks for this window.";
  }

  sheet.columns.forEach((col) => {
    let maxLen = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const len = cell.value ? String(cell.value).length : 0;
      if (len > maxLen) maxLen = len;
    });
    col.width = maxLen + 3;
  });
}

downloadBtn.addEventListener("click", async () => {
  if (!lastReportData || !lastRequestBody) return;

  downloadBtn.disabled = true;
  const originalLabel = downloadBtn.textContent;
  downloadBtn.textContent = "Preparing…";

  try {
    const nDays = lastRequestBody.n_days;
    const workbook = new ExcelJS.Workbook();

    addStreakSheet(workbook, `${nDays}-Day Up Streak`, lastReportData.up_streaks, nDays);
    addStreakSheet(workbook, `${nDays}-Day Down Streak`, lastReportData.down_streaks, nDays);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `nse_streak_report_${lastRequestBody.end_date}_n${nDays}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  } catch (err) {
    setStatus(`Could not generate the Excel file — ${err.message}`, true);
  } finally {
    downloadBtn.disabled = false;
    downloadBtn.textContent = originalLabel;
  }
});
