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

// Default end date = today
endDateInput.value = new Date().toISOString().slice(0, 10);

// Preload default symbol list into the (collapsed) textarea
fetch("/api/default-symbols")
  .then((r) => r.json())
  .then((data) => {
    symbolsTextarea.value = data.symbols.join("\n");
    symbolCountEl.textContent = data.symbols.length;
  })
  .catch(() => {
    symbolCountEl.textContent = "0";
  });

symbolsTextarea.addEventListener("input", () => {
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

  // If Render's free tier has spun the server down from inactivity, the
  // first request can take 30–60s to wake it back up. Let the person know
  // rather than leaving them staring at a stuck spinner.
  const coldStartTimer = setTimeout(() => {
    setStatus("Still working — if this is the first request in a while, the server may be waking up from sleep (can take up to a minute)…");
  }, 6000);

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
  } catch (err) {
    setStatus(`Could not complete the run — ${err.message}`, true);
  } finally {
    clearTimeout(coldStartTimer);
    runBtn.disabled = false;
  }
});

downloadBtn.addEventListener("click", async () => {
  if (!lastRequestBody) return;

  downloadBtn.disabled = true;
  const originalLabel = downloadBtn.textContent;
  downloadBtn.textContent = "Preparing…";

  try {
    const res = await fetch("/api/report/excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(lastRequestBody),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `Request failed (${res.status})`);
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `nse_streak_report_${lastRequestBody.end_date}_n${lastRequestBody.n_days}.xlsx`;
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
