(() => {
  "use strict";

  const PRESET_SOURCES = window.PRESET_SOURCES || {};

  const state = {
    token: null,
    columns: [],          // original columns
    available: {},        // { A: bool, B: bool }
    presets: {},          // { A: bool, B: bool }
    custom: [],           // [{ id, name, columns:[...], sep }]
    last: null,           // last preview result { columns, rows, layout }
  };

  let customSeq = 0;

  // ---- Elements ----
  const pastezone = document.getElementById("pastezone");
  const hasHeader = document.getElementById("has-header");
  const fileLink = document.getElementById("file-link");
  const fileInput = document.getElementById("file-input");
  const uploadStage = document.getElementById("upload-stage");
  const uploadError = document.getElementById("upload-error");
  const workArea = document.getElementById("work-area");
  const warningBanner = document.getElementById("warning-banner");
  const strategyList = document.getElementById("strategy-list");
  const customList = document.getElementById("custom-list");
  const builder = document.getElementById("builder");
  const builderName = document.getElementById("builder-name");
  const builderSep = document.getElementById("builder-sep");
  const builderCols = document.getElementById("builder-cols");
  const builderError = document.getElementById("builder-error");
  const builderAdd = document.getElementById("builder-add");
  const table = document.getElementById("output-table");
  const fitIndicator = document.getElementById("fit-indicator");
  const fitIcon = document.getElementById("fit-icon");
  const fitText = document.getElementById("fit-text");
  const metaName = document.getElementById("meta-name");
  const metaRows = document.getElementById("meta-rows");
  const metaCols = document.getElementById("meta-cols");
  const resetBtn = document.getElementById("reset-btn");
  const selectBtn = document.getElementById("select-btn");
  const copyBtn = document.getElementById("copy-btn");
  const copyFeedback = document.getElementById("copy-feedback");

  // ---- Helpers ----
  function showError(msg) {
    uploadError.textContent = msg;
    uploadError.hidden = false;
  }
  function clearError() {
    uploadError.hidden = true;
    uploadError.textContent = "";
  }
  function flash(msg) {
    copyFeedback.textContent = msg;
    copyFeedback.hidden = false;
    clearTimeout(flash._t);
    flash._t = setTimeout(() => (copyFeedback.hidden = true), 2500);
  }

  // Columns already consumed by an active preset or a custom merge.
  function consumedColumns() {
    const set = new Set();
    Object.keys(state.presets).forEach((key) => {
      if (state.presets[key] && state.available[key]) {
        (PRESET_SOURCES[key] || []).forEach((c) => set.add(c));
      }
    });
    state.custom.forEach((m) => m.columns.forEach((c) => set.add(c)));
    return set;
  }

  // ---- Paste ingestion ----
  // Parse an HTML <table> from the clipboard into a grid of trimmed strings.
  function parseHtmlTable(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (!table) return null;
    const grid = [];
    table.querySelectorAll("tr").forEach((tr) => {
      const cells = [...tr.querySelectorAll("th,td")].map((c) =>
        c.textContent.replace(/\s+/g, " ").trim()
      );
      if (cells.length) grid.push(cells);
    });
    return grid.length ? grid : null;
  }

  // Parse delimited plain text (tab from spreadsheets, comma as fallback).
  function parseTextTable(text) {
    const lines = text.replace(/\r\n/g, "\n").replace(/\s+$/, "").split("\n");
    if (!lines.length) return null;
    const delim = lines[0].indexOf("\t") !== -1 ? "\t" : ",";
    const grid = lines
      .filter((l) => l.trim() !== "")
      .map((l) => l.split(delim).map((c) => c.trim()));
    return grid.length ? grid : null;
  }

  function gridToColumnsRows(grid) {
    // Normalize ragged rows to the widest row.
    const width = grid.reduce((m, r) => Math.max(m, r.length), 0);
    const padded = grid.map((r) => {
      const copy = r.slice();
      while (copy.length < width) copy.push("");
      return copy;
    });

    let columns, rows;
    if (hasHeader.checked) {
      columns = padded[0];
      rows = padded.slice(1);
    } else {
      columns = padded[0].map((_, i) => `Column ${i + 1}`);
      rows = padded;
    }
    return { columns, rows };
  }

  async function handlePaste(e) {
    e.preventDefault();
    clearError();
    const cd = e.clipboardData || window.clipboardData;
    if (!cd) return;

    const html = cd.getData("text/html");
    const text = cd.getData("text/plain");

    let grid = (html && parseHtmlTable(html)) || (text && parseTextTable(text));
    if (!grid || grid.length === 0) {
      showError("Could not read a table from the clipboard. Copy a cell range from your spreadsheet and try again.");
      return;
    }

    const { columns, rows } = gridToColumnsRows(grid);
    if (!columns.length || !rows.length) {
      showError("The pasted table needs at least a header row and one data row.");
      return;
    }

    try {
      const res = await fetch("/ingest", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ columns, rows }),
      });
      const data = await res.json();
      if (!res.ok) {
        showError(data.error || "Could not read the pasted table.");
        return;
      }
      enterWorkArea(data, "Pasted table");
    } catch (err) {
      showError("Something went wrong reading the pasted table. Please try again.");
    }
  }

  // Shared setup once a grid is ingested (from paste OR file upload).
  function enterWorkArea(data, sourceName) {
    state.token = data.token;
    state.columns = data.columns || [];
    state.available = data.available || {};
    state.presets = {};
    state.custom = [];

    metaName.textContent = sourceName;
    metaRows.textContent = data.row_count;
    metaCols.textContent = state.columns.length;

    configureStrategies(state.available);
    renderCustomList();
    renderBuilderColumns();

    if (data.warning) {
      warningBanner.textContent = data.warning;
      warningBanner.hidden = false;
    } else {
      warningBanner.hidden = true;
    }

    uploadStage.hidden = true;
    workArea.hidden = false;
    refreshPreview();
  }

  // ---- Upload (secondary path) ----
  async function uploadFile(file) {
    clearError();
    if (!file) return;

    const name = file.name.toLowerCase();
    if (!name.endsWith(".xlsx") && !name.endsWith(".xlsm") && !name.endsWith(".csv")) {
      showError("Unsupported file type. Please upload a .xlsx or .csv file.");
      return;
    }

    const form = new FormData();
    form.append("file", file);

    try {
      const res = await fetch("/upload", { method: "POST", body: form });
      const data = await res.json();
      if (!res.ok) {
        showError(data.error || "Upload failed.");
        return;
      }
      enterWorkArea(data, file.name);
    } catch (err) {
      showError("Something went wrong while uploading. Please try again.");
    }
  }

  // Enable/disable preset toggles based on header availability.
  function configureStrategies(available) {
    strategyList.querySelectorAll(".strategy").forEach((label) => {
      const key = label.dataset.strategy;
      const toggle = label.querySelector(".strategy-toggle");
      const status = label.querySelector("[data-status]");
      const ok = !!available[key];

      toggle.checked = false;
      toggle.disabled = !ok;
      label.classList.toggle("disabled", !ok);

      if (ok) {
        status.textContent = "";
        status.classList.remove("unavailable");
      } else {
        status.textContent = "Required columns not found in this file.";
        status.classList.add("unavailable");
      }
    });
  }

  // ---- Custom merge builder ----
  function renderCustomList() {
    customList.innerHTML = "";
    if (state.custom.length === 0) {
      const empty = document.createElement("p");
      empty.className = "config-note";
      empty.style.margin = "0";
      empty.textContent = "No custom merges yet.";
      customList.appendChild(empty);
      return;
    }
    state.custom.forEach((m) => {
      const chip = document.createElement("div");
      chip.className = "custom-chip";

      const body = document.createElement("div");
      body.className = "custom-chip-body";
      const nm = document.createElement("span");
      nm.className = "custom-chip-name";
      nm.textContent = m.name || m.columns.join(" / ");
      const desc = document.createElement("span");
      desc.className = "custom-chip-desc";
      desc.textContent = m.columns.join(` ${m.sep.trim() || "·"} `);
      body.appendChild(nm);
      body.appendChild(desc);

      const remove = document.createElement("button");
      remove.className = "custom-chip-remove";
      remove.type = "button";
      remove.setAttribute("aria-label", "Remove merge");
      remove.textContent = "\u00d7";
      remove.addEventListener("click", () => {
        state.custom = state.custom.filter((x) => x.id !== m.id);
        renderCustomList();
        renderBuilderColumns();
        refreshPreview();
      });

      chip.appendChild(body);
      chip.appendChild(remove);
      customList.appendChild(chip);
    });
  }

  function renderBuilderColumns() {
    const taken = consumedColumns();
    builderCols.innerHTML = "";
    state.columns.forEach((col) => {
      const isTaken = taken.has(col);
      const row = document.createElement("label");
      row.className = "builder-col" + (isTaken ? " taken" : "");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.value = col;
      cb.disabled = isTaken;
      const span = document.createElement("span");
      span.textContent = col + (isTaken ? " (already merged)" : "");
      row.appendChild(cb);
      row.appendChild(span);
      builderCols.appendChild(row);
    });
  }

  function addCustomMerge() {
    builderError.hidden = true;
    const picked = [...builderCols.querySelectorAll("input:checked")].map((c) => c.value);
    if (picked.length < 2) {
      builderError.textContent = "Select at least two columns to merge.";
      builderError.hidden = false;
      return;
    }
    const sep = builderSep.value === "" ? " - " : builderSep.value;
    const name = builderName.value.trim();
    state.custom.push({ id: ++customSeq, name, columns: picked, sep });

    builderName.value = "";
    builderSep.value = " - ";
    builder.open = false;

    renderCustomList();
    renderBuilderColumns();
    refreshPreview();
  }

  // ---- Live preview ----
  async function refreshPreview() {
    if (!state.token) return;
    try {
      const res = await fetch("/preview", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          token: state.token,
          options: { presets: state.presets, custom: state.custom },
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        showError(data.error || "Could not refresh the preview.");
        return;
      }
      state.last = { columns: data.columns, rows: data.rows, layout: data.layout };
      renderTable(data.columns, data.rows, data.layout);
      updateFit(data.layout, data.columns.length);
    } catch (err) {
      showError("Could not refresh the preview.");
    }
  }

  function renderTable(columns, rows, layout) {
    const widths = (layout && layout.widths_in) || [];
    const printable = (layout && layout.printable_in) || 7.5;

    // Pin the table width inline (in inches) so Word honors the 7.5" footprint.
    table.style.width = printable + "in";

    const colgroup = document.createElement("colgroup");
    columns.forEach((_, i) => {
      const col = document.createElement("col");
      if (widths[i]) col.style.width = widths[i] + "in";
      colgroup.appendChild(col);
    });

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    columns.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);

    const tbody = document.createElement("tbody");
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = cell;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.innerHTML = "";
    table.appendChild(colgroup);
    table.appendChild(thead);
    table.appendChild(tbody);
  }

  // Golden-rule indicator: is the table comfortable within 7.5", or squeezed?
  function updateFit(layout, colCount) {
    if (!layout) {
      fitIndicator.hidden = true;
      return;
    }
    fitIndicator.hidden = false;
    const printable = layout.printable_in;
    const needed = layout.needed_in;

    if (layout.fits) {
      fitIndicator.classList.add("ok");
      fitIndicator.classList.remove("tight");
      fitIcon.textContent = "\u2713";
      fitText.innerHTML =
        `<strong>Word-ready.</strong> ${colCount} columns fit cleanly and readably ` +
        `inside the ${printable}&Prime; printable width of an 8.5&times;11&Prime; page. Copy it over.`;
    } else {
      fitIndicator.classList.add("tight");
      fitIndicator.classList.remove("ok");
      fitIcon.textContent = "\u26a0";
      fitText.innerHTML =
        `<strong>Too cramped for Word.</strong> ${colCount} columns need about ` +
        `${needed}&Prime; to stay readable, but a page only has ${printable}&Prime;. ` +
        `Merge some columns on the left (presets or a custom merge) until this turns green.`;
    }
  }

  // ---- Selection & copy ----
  function selectTable() {
    const range = document.createRange();
    range.selectNode(table);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    return range;
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // Build a fully INLINE-styled table for the clipboard. Word ignores external
  // CSS/classes, so (just like Excel) every border/width/font must be inline on
  // the elements themselves for the formatting to survive the paste.
  function buildWordHtml() {
    if (!state.last) return "";
    const { columns, rows, layout } = state.last;
    const widths = (layout && layout.widths_in) || [];
    const printable = (layout && layout.printable_in) || 7.5;

    const cellBase =
      "border:1px solid #000000;padding:4px 7px;vertical-align:top;" +
      "font-family:Calibri,Arial,sans-serif;font-size:11pt;" +
      "word-wrap:break-word;overflow-wrap:break-word;";

    const colWidth = (i) => (widths[i] ? `width:${widths[i]}in;` : "");

    let html =
      `<table border="1" cellspacing="0" cellpadding="0" ` +
      `style="border-collapse:collapse;table-layout:fixed;width:${printable}in;">`;

    html += "<colgroup>";
    columns.forEach((_, i) => {
      html += `<col style="${colWidth(i)}">`;
    });
    html += "</colgroup>";

    html += "<thead><tr>";
    columns.forEach((c, i) => {
      html +=
        `<th style="${cellBase}${colWidth(i)}` +
        `background-color:#D9E2F3;font-weight:bold;text-align:left;">${esc(c)}</th>`;
    });
    html += "</tr></thead>";

    html += "<tbody>";
    rows.forEach((r) => {
      html += "<tr>";
      r.forEach((cell, i) => {
        html += `<td style="${cellBase}${colWidth(i)}">${esc(cell)}</td>`;
      });
      html += "</tr>";
    });
    html += "</tbody></table>";

    return html;
  }

  // Tab-separated plain-text fallback (Word can convert text->table from this).
  function buildTsv() {
    if (!state.last) return "";
    const { columns, rows } = state.last;
    return [columns.join("\t"), ...rows.map((r) => r.join("\t"))].join("\n");
  }

  // Fallback copy: drop the inline-styled HTML into an offscreen node, select it,
  // and use execCommand so the rich formatting still lands on the clipboard.
  function copyViaSelection(html) {
    const holder = document.createElement("div");
    holder.setAttribute("contenteditable", "true");
    holder.style.position = "fixed";
    holder.style.left = "-99999px";
    holder.style.top = "0";
    holder.innerHTML = html;
    document.body.appendChild(holder);

    const range = document.createRange();
    range.selectNodeContents(holder);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);

    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch (_) {
      ok = false;
    }
    sel.removeAllRanges();
    document.body.removeChild(holder);
    return ok;
  }

  async function copyTable() {
    const html = buildWordHtml();
    const text = buildTsv();
    if (!html) {
      flash("Nothing to copy yet.");
      return;
    }
    let ok = false;

    // Preferred: rich HTML clipboard with fully inline styles so Word keeps the
    // borders, header shading, and the 7.5" column widths.
    if (navigator.clipboard && window.ClipboardItem) {
      try {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([html], { type: "text/html" }),
            "text/plain": new Blob([text], { type: "text/plain" }),
          }),
        ]);
        ok = true;
      } catch (_) {
        ok = false;
      }
    }

    if (!ok) {
      ok = copyViaSelection(html);
    }

    flash(ok
      ? "Table copied. Paste into Word with Ctrl/Cmd+V."
      : "Table selected. Press Ctrl/Cmd+C to copy.");
  }

  // ---- Reset ----
  function reset() {
    state.token = null;
    state.columns = [];
    state.presets = {};
    state.custom = [];
    state.available = {};
    state.last = null;
    fileInput.value = "";
    table.innerHTML = "<tbody></tbody>";
    fitIndicator.hidden = true;
    workArea.hidden = true;
    uploadStage.hidden = false;
    warningBanner.hidden = true;
    clearError();
  }

  // ---- Wire up events ----
  pastezone.addEventListener("paste", handlePaste);
  // Keep the placeholder clean: contenteditable should not retain typed text.
  pastezone.addEventListener("keydown", (e) => {
    // Allow copy/select/navigation; block free typing so it stays a paste target.
    const allowed = e.metaKey || e.ctrlKey || ["Tab", "Escape"].includes(e.key);
    if (!allowed) e.preventDefault();
  });

  fileLink.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => uploadFile(e.target.files[0]));

  strategyList.addEventListener("change", (e) => {
    const toggle = e.target.closest(".strategy-toggle");
    if (!toggle) return;
    state.presets[toggle.value] = toggle.checked;
    renderBuilderColumns();
    refreshPreview();
  });

  builderAdd.addEventListener("click", addCustomMerge);
  resetBtn.addEventListener("click", reset);
  selectBtn.addEventListener("click", () => {
    selectTable();
    flash("Table selected. Press Ctrl/Cmd+C to copy.");
  });
  copyBtn.addEventListener("click", copyTable);
})();
