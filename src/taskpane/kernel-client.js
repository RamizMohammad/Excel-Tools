/**
 * Talks to the JupyXL kernel server over a WebSocket.
 *
 * Usage:
 *   const k = new KernelClient("ws://localhost:8765/ws");
 *   k.on("dataframe", m => ...);
 *   await k.connect();
 *   k.execute("import pandas as pd; pd.DataFrame({'a':[1,2]})", "cell-1");
 */
export class KernelClient {
  constructor(url) {
    this.url = url;
    this.ws = null;
    this.ready = false;
    this.handlers = {}; // type -> [fn]
    this._reconnectDelay = 1000;
    this._shouldReconnect = true;
  }

  on(type, fn) {
    (this.handlers[type] ||= []).push(fn);
    return this;
  }

  _emit(type, msg) {
    (this.handlers[type] || []).forEach((fn) => fn(msg));
    (this.handlers["*"] || []).forEach((fn) => fn(msg));
  }

  connect() {
    return new Promise((resolve, reject) => {
      try {
        this.ws = new WebSocket(this.url);
      } catch (e) {
        reject(e);
        return;
      }

      this.ws.onopen = () => {
        this._reconnectDelay = 1000;
        this._emit("open", {});
      };

      this.ws.onmessage = (ev) => {
        let msg;
        try {
          msg = JSON.parse(ev.data);
        } catch {
          return;
        }
        if (msg.type === "kernel_ready") {
          this.ready = true;
          resolve();
        }
        this._emit(msg.type, msg);
      };

      this.ws.onerror = (err) => {
        this._emit("ws_error", err);
        if (!this.ready) reject(new Error("Could not reach kernel at " + this.url));
      };

      this.ws.onclose = () => {
        this.ready = false;
        this._emit("close", {});
        if (this._shouldReconnect) {
          setTimeout(() => this.connect().catch(() => {}), this._reconnectDelay);
          this._reconnectDelay = Math.min(this._reconnectDelay * 2, 15000);
        }
      };
    });
  }

  _send(obj) {
    if (this.ws && this.ws.readyState === WebSocket.OPEN) {
      this.ws.send(JSON.stringify(obj));
    } else {
      this._emit("ws_error", new Error("Socket not open"));
    }
  }

  execute(code, cellId) {
    this._send({ type: "execute", code, cell_id: cellId });
  }

  interrupt() {
    this._send({ type: "interrupt" });
  }

  restart() {
    this._send({ type: "restart" });
  }

  close() {
    this._shouldReconnect = false;
    if (this.ws) this.ws.close();
  }
}
