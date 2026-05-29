# JupyXL — a live Jupyter kernel inside Excel

Run real Python in a notebook panel docked inside Excel, and push **formatted
DataFrames straight into the grid** — headers styled, number formats inferred
from dtypes, dates written as real Excel dates, optionally wrapped as a Table.

It is an **Office.js add-in** (so the same build runs on Excel for Windows, Mac
and the web) talking over a WebSocket to a **local FastAPI server** that hosts a
genuine IPython kernel via `jupyter_client`. The kernel endpoint is a single
setting, so you can swap in a hosted `wss://` kernel later without touching the
front end.

```
┌────────────────────────┐        WebSocket         ┌─────────────────────────┐
│  Excel task pane         │  ws(s)://host/ws         │  FastAPI kernel server   │
│  (Office.js, this repo)  │ ───────────────────────▶ │  jupyter_client + ipykernel
│                          │ ◀─────────────────────── │                         │
│  • notebook cells        │   stream / result /      │  • one kernel / session │
│  • DataFrame → Range API │   dataframe / error      │  • DataFrame serializer │
└────────────────────────┘                          │  • xl.write() helper    │
                                                      └─────────────────────────┘
```

## Why the kernel is local (and how the web version works)

An Office add-in is a web app loaded into a webview Excel hosts — WebView2 on
Windows, WKWebView on Mac, an `<iframe>` on Excel for the web. A **local**
kernel is ideal for a dev tool: full access to your installed packages, private
data never leaves your machine, zero server cost. On desktop Excel the task pane
reaches `localhost` directly (this is how xlwings works too).

For **Excel on the web**, the page is served over HTTPS, so calling
`ws://localhost` is technically *mixed content*. Browsers (Chrome/Edge) treat
`localhost`/`127.0.0.1` as a trustworthy secure origin, so it is permitted **if**
the local server sets CORS headers — but the user still has to be running the
local kernel on the same machine, and it won't work on a tablet. For a true
zero-install web experience, run the kernel server on a host you control and put
its `wss://` URL in **Settings → Kernel endpoint**. Everything else is identical.

> Tip: to avoid the mixed-content edge case entirely, run the kernel server with
> `--tls` so it serves `wss://` using the same dev certificate as the task pane.

## Setup

### 1. Kernel server
```bash
cd server
python -m venv .venv && source .venv/bin/activate   # optional
pip install -r requirements.txt
python main.py            # ws://localhost:8765/ws
# or, matching the HTTPS task pane cleanly:
python main.py --tls      # wss://localhost:8765/ws   (needs step 2's certs)
```

### 2. Task pane (dev)
```bash
npm install
npm run certs             # trust the localhost dev certificate (once)
npm run dev               # serves the add-in at https://localhost:3000
npm start                 # sideloads manifest.xml into Excel and opens it
```
Then in Excel: **Home → JupyXL → Notebook**. If you used `--tls` on the server,
set the endpoint in Settings to `wss://localhost:8765/ws`.

### Build-free alternative
The task pane uses native ES modules, so you can skip webpack entirely and serve
`src/taskpane/` over HTTPS, pointing the manifest's `SourceLocation` there.

## Using it

```python
import pandas as pd
df = pd.read_csv("sales.csv").groupby("region").revenue.sum().reset_index()

xl.write(df, anchor="A1")              # explicit: write here, as a Table
xl.write(df, anchor="B2", sheet="Summary", table=False)
df                                      # bare DataFrame → preview + "Send to grid" button
```

- `xl.write(df, anchor, sheet, table)` writes immediately to the chosen spot.
- Leaving a bare DataFrame as the last expression shows a preview with a
  **Send to grid →** button and an anchor field.
- **⤓ Selection** pulls your current Excel selection into a new cell as a DataFrame.
- Plots (`matplotlib`) render inline as images; `print()` and reprs stream in.

### How formatting is applied
When a DataFrame lands in the grid the add-in:
- writes the header row and bolds it (or wraps the whole block as an Excel Table),
- infers a number format per column from the pandas dtype
  (`int → #,##0`, `float → #,##0.00`, `datetime → yyyy-mm-dd hh:mm`),
- converts datetime values to real Excel date serials (not text),
- blanks `NaN`/`None`, and autofits columns.

## Files
```
manifest.xml                Office add-in manifest (task pane + ribbon button)
package.json / webpack      dev server, sideload, dev-cert tooling
server/
  main.py                   FastAPI WebSocket endpoint + wire protocol
  kernel_manager.py         async jupyter_client session (1 kernel / connection)
  kernel_bootstrap.py       injected DataFrame serializer + xl.write helper
src/taskpane/
  taskpane.html/.css        the notebook UI
  taskpane.js               cells, output rendering, kernel wiring
  kernel-client.js          WebSocket client (clean message protocol, reconnect)
  excel-writer.js           DataFrame → Range with dtype-aware formatting
src/commands/               ribbon command runtime
```

## Roadmap / known gaps
- **Auth** for hosted kernels (token in the WS URL or a header) — not in the MVP.
- **Sandboxing** if you host the kernel: run per-user containers with CPU/memory
  limits; never expose an unsandboxed kernel to the internet.
- Syntax highlighting (drop in CodeMirror/Monaco in place of the textarea cells).
- Two-way binding (named ranges that re-read on change) and `=PY()`-style custom
  functions are natural next steps.
- Replace the placeholder icons in `assets/`.
