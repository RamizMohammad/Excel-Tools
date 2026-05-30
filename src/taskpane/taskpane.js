/* global Office */
import { readSelection, writeDataFrame } from "./excel-writer.js";
import { KernelClient } from "./kernel-client.js";

const DEFAULT_KERNEL_URL =
  location.hostname === "localhost" || location.hostname === "127.0.0.1"
    ? "ws://localhost:8008/ws"
    : `wss://${location.hostname}/ws`;

const state = {
  client: null,
  cells: [],          // {id, code, el, outputEl, lastDataframe}
  counter: 0,
  busyCell: null,
};

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    document.getElementById("app").innerHTML =
      "<p class='hint'>JupyXL runs inside Excel. Open it from the Excel ribbon.</p>";
    return;
  }
  wireToolbar();
  loadKernelUrl();
  addCell("import pandas as pd\n\ndf = pd.DataFrame({\n    \"region\": [\"North\", \"South\", \"East\"],\n    \"revenue\": [12500.5, 9800.0, 14300.75],\n    \"orders\": [120, 95, 143],\n})\nxl.write(df, anchor=\"A1\")");
  connectKernel();
});

// ---------------------------------------------------------------------------
// Kernel connection
// ---------------------------------------------------------------------------
function loadKernelUrl() {
  const saved = Office.context.document.settings.get("jupyxl_kernel_url");
  document.getElementById("kernel-url").value = saved || DEFAULT_KERNEL_URL;
}

function saveKernelUrl(url) {
  Office.context.document.settings.set("jupyxl_kernel_url", url);
  Office.context.document.settings.saveAsync();
}

function setStatus(text, cls) {
  const el = document.getElementById("kernel-status");
  el.textContent = text;
  el.className = "status " + (cls || "");
}

async function connectKernel() {
  const url = document.getElementById("kernel-url").value.trim() || DEFAULT_KERNEL_URL;
  saveKernelUrl(url);

  if (state.client) state.client.close();
  setStatus("connecting…", "pending");

  const client = new KernelClient(url);
  state.client = client;

  client
    .on("open", () => setStatus("handshaking…", "pending"))
    .on("close", () => setStatus("disconnected", "error"))
    .on("ws_error", () => setStatus("no kernel — is the server running?", "error"))
    .on("fatal", (m) => setStatus(m.message || "kernel error", "error"))
    .on("status", (m) => onKernelStatus(m))
    .on("stream", (m) => appendStream(m))
    .on("result", (m) => appendResult(m))
    .on("html", (m) => appendHtml(m))
    .on("image", (m) => appendImage(m))
    .on("dataframe", (m) => onDataframe(m))
    .on("error", (m) => appendError(m))
    .on("kernel_restarted", () => setStatus("kernel restarted", "ok"));

  try {
    await client.connect();
    setStatus("kernel ready", "ok");
  } catch (e) {
    setStatus("no kernel — is the server running?", "error");
  }
}

function onKernelStatus(m) {
  if (m.state === "busy") {
    setStatus("running…", "pending");
  } else if (m.state === "idle") {
    setStatus("kernel ready", "ok");
    if (state.busyCell) {
      state.busyCell.el.classList.remove("running");
      state.busyCell = null;
    }
  }
}

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------
function addCell(initialCode = "", afterCell = null) {
  const id = "cell-" + ++state.counter;
  const cell = { id, code: initialCode, el: null, outputEl: null, lastDataframe: null };

  const wrapper = document.createElement("div");
  wrapper.className = "cell";
  wrapper.dataset.id = id;
  wrapper.innerHTML = `
    <div class="cell-gutter">
      <button class="run-btn" title="Run (Ctrl/Cmd+Enter)">▶</button>
    </div>
    <div class="cell-main">
      <textarea class="code" spellcheck="false" rows="3"></textarea>
      <div class="output"></div>
    </div>
    <div class="cell-actions">
      <button class="add-below" title="Add cell below">+</button>
      <button class="del-cell" title="Delete cell">×</button>
    </div>`;

  const textarea = wrapper.querySelector(".code");
  textarea.value = initialCode;

  textarea.addEventListener("input", () => {
    cell.code = textarea.value;
    autoSize(textarea);
  });
  textarea.addEventListener("keydown", (e) => handleKeydown(e, cell, textarea));

  wrapper.querySelector(".run-btn").addEventListener("click", () => runCell(cell));
  wrapper.querySelector(".add-below").addEventListener("click", () => addCell("", cell));
  wrapper.querySelector(".del-cell").addEventListener("click", () => deleteCell(cell));

  cell.el = wrapper;
  cell.outputEl = wrapper.querySelector(".output");

  const list = document.getElementById("cells");
  if (afterCell && afterCell.el.nextSibling) {
    list.insertBefore(wrapper, afterCell.el.nextSibling);
    const idx = state.cells.indexOf(afterCell);
    state.cells.splice(idx + 1, 0, cell);
  } else {
    list.appendChild(wrapper);
    state.cells.push(cell);
  }
  
  // Auto-size after DOM insertion to ensure proper calculations
  requestAnimationFrame(() => autoSize(textarea));
  textarea.focus();
  return cell;
}

function deleteCell(cell) {
  cell.el.remove();
  state.cells = state.cells.filter((c) => c !== cell);
  if (state.cells.length === 0) addCell();
}

function handleKeydown(e, cell, textarea) {
  // Tab inserts spaces instead of moving focus
  if (e.key === "Tab") {
    e.preventDefault();
    const s = textarea.selectionStart;
    const eN = textarea.selectionEnd;
    textarea.value = textarea.value.slice(0, s) + "    " + textarea.value.slice(eN);
    textarea.selectionStart = textarea.selectionEnd = s + 4;
    cell.code = textarea.value;
    autoSize(textarea);
    return;
  }
  const run = (e.ctrlKey || e.metaKey) && e.key === "Enter";
  const runNext = e.shiftKey && e.key === "Enter";
  if (run) {
    e.preventDefault();
    runCell(cell);
  } else if (runNext) {
    e.preventDefault();
    runCell(cell);
    const idx = state.cells.indexOf(cell);
    if (idx === state.cells.length - 1) addCell("", cell);
    else state.cells[idx + 1].el.querySelector(".code").focus();
  }
}

function autoSize(ta) {
  ta.style.height = "auto";
  ta.style.height = Math.min(ta.scrollHeight + 2, 480) + "px";
}

function runCell(cell) {
  if (!state.client || !state.client.ready) {
    appendError({ cell_id: cell.id, ename: "NoKernel", evalue: "Kernel not connected. Check the endpoint in Settings.", traceback: [] });
    return;
  }
  cell.outputEl.innerHTML = "";
  cell.lastDataframe = null;
  cell.el.classList.add("running");
  state.busyCell = cell;
  state.client.execute(cell.code, cell.id);
}

// ---------------------------------------------------------------------------
// Output rendering  (each looks up its owning cell by cell_id)
// ---------------------------------------------------------------------------
function cellById(id) {
  return state.cells.find((c) => c.id === id);
}

function appendStream(m) {
  const cell = cellById(m.cell_id);
  if (!cell) return;
  const pre = document.createElement("pre");
  pre.className = "stream " + (m.name === "stderr" ? "stderr" : "stdout");
  pre.textContent = m.text;
  cell.outputEl.appendChild(pre);
}

function appendResult(m) {
  const cell = cellById(m.cell_id);
  if (!cell) return;
  const pre = document.createElement("pre");
  pre.className = "result";
  pre.textContent = m.data;
  cell.outputEl.appendChild(pre);
}

function appendHtml(m) {
  const cell = cellById(m.cell_id);
  if (!cell) return;
  const div = document.createElement("div");
  div.className = "html-out";
  div.innerHTML = m.data; // pandas .to_html etc. Trusted: it's the user's own kernel.
  cell.outputEl.appendChild(div);
}

function appendImage(m) {
  const cell = cellById(m.cell_id);
  if (!cell) return;
  const img = document.createElement("img");
  img.className = "img-out";
  img.src = `data:${m.mime};base64,${m.data}`;
  cell.outputEl.appendChild(img);
}

function appendError(m) {
  const cell = cellById(m.cell_id);
  const target = cell ? cell.outputEl : null;
  const pre = document.createElement("pre");
  pre.className = "error";
  const tb = (m.traceback && m.traceback.length)
    ? stripAnsi(m.traceback.join("\n"))
    : `${m.ename}: ${m.evalue}`;
  pre.textContent = tb;
  if (target) target.appendChild(pre);
}

function stripAnsi(s) {
  // eslint-disable-next-line no-control-regex
  return s.replace(/\x1b\[[0-9;]*m/g, "");
}

// ---------------------------------------------------------------------------
// DataFrame -> grid
// ---------------------------------------------------------------------------
async function onDataframe(m) {
  const cell = cellById(m.cell_id);
  if (!cell) return;
  cell.lastDataframe = m.payload;

  const hasTarget = m.payload.target && m.payload.target.anchor;
  const preview = renderDfPreview(m.payload);
  cell.outputEl.appendChild(preview.el);

  if (hasTarget) {
    // xl.write(...) was called explicitly -> push immediately.
    try {
      const res = await writeDataFrame(m.payload);
      preview.setNote(`written to ${res.address}`);
    } catch (e) {
      preview.setNote("write failed: " + (e.message || e), true);
    }
  } else {
    // Bare DataFrame -> offer a button.
    preview.addPushButton(async (opts) => {
      try {
        const res = await writeDataFrame(m.payload, opts);
        preview.setNote(`written to ${res.address}`);
      } catch (e) {
        preview.setNote("write failed: " + (e.message || e), true);
      }
    });
  }
}

function renderDfPreview(payload) {
  const { columns, rows, nrows, ncols, truncated_rows } = payload;
  const wrap = document.createElement("div");
  wrap.className = "df-out";

  const meta = document.createElement("div");
  meta.className = "df-meta";
  meta.textContent = `DataFrame · ${nrows} × ${ncols}` + (truncated_rows ? " (preview truncated)" : "");
  wrap.appendChild(meta);

  // small HTML preview (first 8 rows)
  const table = document.createElement("table");
  table.className = "df-table";
  const thead = document.createElement("thead");
  const htr = document.createElement("tr");
  columns.forEach((c) => {
    const th = document.createElement("th");
    th.textContent = c;
    htr.appendChild(th);
  });
  thead.appendChild(htr);
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  rows.slice(0, 8).forEach((r) => {
    const tr = document.createElement("tr");
    r.forEach((v) => {
      const td = document.createElement("td");
      td.textContent = v === null ? "" : String(v);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrap.appendChild(table);

  const note = document.createElement("div");
  note.className = "df-note";
  wrap.appendChild(note);

  return {
    el: wrap,
    setNote(text, isError) {
      note.textContent = text;
      note.classList.toggle("err", !!isError);
    },
    addPushButton(onPush) {
      const bar = document.createElement("div");
      bar.className = "df-actions";
      const anchorInput = document.createElement("input");
      anchorInput.className = "anchor-input";
      anchorInput.value = "A1";
      anchorInput.title = "Anchor cell";
      const btn = document.createElement("button");
      btn.className = "push-btn";
      btn.textContent = "Send to grid →";
      btn.addEventListener("click", () =>
        onPush({ anchor: anchorInput.value.trim() || "A1", table: true })
      );
      bar.appendChild(anchorInput);
      bar.appendChild(btn);
      wrap.insertBefore(bar, note);
    },
  };
}

// ---------------------------------------------------------------------------
// Toolbar
// ---------------------------------------------------------------------------
function wireToolbar() {
  document.getElementById("btn-add").addEventListener("click", () => addCell());
  document.getElementById("btn-run-all").addEventListener("click", runAll);
  document.getElementById("btn-restart").addEventListener("click", () => {
    if (state.client) state.client.restart();
  });
  document.getElementById("btn-interrupt").addEventListener("click", () => {
    if (state.client) state.client.interrupt();
  });
  document.getElementById("btn-settings").addEventListener("click", () => {
    document.getElementById("settings").classList.toggle("open");
  });
  document.getElementById("btn-connect").addEventListener("click", connectKernel);
  document.getElementById("btn-read-selection").addEventListener("click", insertSelectionCell);
}

async function runAll() {
  for (const cell of state.cells) {
    runCell(cell);
    await waitIdle();
  }
}

function waitIdle() {
  return new Promise((resolve) => {
    const check = () => {
      if (!state.busyCell) resolve();
      else setTimeout(check, 80);
    };
    setTimeout(check, 80);
  });
}

/** Pull the current Excel selection into a new cell as a DataFrame. */
async function insertSelectionCell() {
  try {
    const sel = await readSelection();
    const json = JSON.stringify(sel.values);
    const code =
      "import pandas as pd\n" +
      "_vals = " + json + "\n" +
      "df = pd.DataFrame(_vals[1:], columns=_vals[0])\n" +
      "df";
    addCell(code);
  } catch (e) {
    addCell("# Could not read selection: " + (e.message || e));
  }
}
