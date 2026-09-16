(function (root) {
  'use strict';
  const PREFIX = 'phone-workbench-session-';
  const VERSION = '20260916-import-hours-v1';
  const abortError = () => new DOMException('已取消處理。', 'AbortError');
  function deleteDatabase(name) {
    return new Promise((resolve, reject) => {
      const request = indexedDB.deleteDatabase(name);
      request.onsuccess = resolve;
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(new Error('暫存仍被使用中，請關閉舊分頁後重試。'));
    });
  }
  class Client {
    constructor(onProgress) {
      this.sessionId = crypto.randomUUID(); this.pending = new Map(); this.sequence = 0;
      this.onProgress = onProgress; this.worker = null; this.ready = null; this.closed = false;
    }
    async initialize() {
      if (!root.Worker || !root.indexedDB || !navigator.locks || !indexedDB.databases) throw new Error('請使用最新版 64 位元 Edge 或 Chrome，並允許此網站使用本機儲存空間。');
      await new Promise((resolve, reject) => {
        navigator.locks.request(PREFIX + this.sessionId, () => new Promise(release => { this.release = release; resolve(); })).catch(reject);
      });
      for (const { name } of await indexedDB.databases()) {
        if (!name?.startsWith(PREFIX) || name === PREFIX + this.sessionId) continue;
        await navigator.locks.request(name, { ifAvailable: true }, async lock => { if (lock) await deleteDatabase(name); });
      }
      await this.start();
      this.pagehide = () => { this.closed = true; this.worker?.terminate(); this.release?.(); deleteDatabase(PREFIX + this.sessionId).catch(() => {}); };
      root.addEventListener('pagehide', this.pagehide);
      return this;
    }
    start() {
      this.worker = new Worker('./dataset-worker.js?v=' + VERSION);
      this.worker.onmessage = ({ data }) => {
        const pending = this.pending.get(data.id);
        if (!pending) return;
        if (data.progress) { this.onProgress?.(data.progress); return; }
        this.pending.delete(data.id);
        if (data.ok) pending.resolve(data.result);
        else { const error = new Error(data.message || '本機處理失敗，請重新匯入。'); error.name = data.name || 'Error'; pending.reject(error); }
      };
      this.failed = false;
      this.worker.onerror = () => { this.failed = true; this.worker.terminate(); this.fail(new Error('背景處理程序意外停止，請重新匯入。')); };
      this.ready = this.send('init', { sessionId: this.sessionId, retainId: this.activeDatasetId }, 'init');
      return this.ready;
    }
    activate(id) { this.activeDatasetId = id; }
    fail(error) { for (const p of this.pending.values()) p.reject(error); this.pending.clear(); }
    send(op, payload, channel) {
      const id = ++this.sequence;
      return new Promise((resolve, reject) => {
        this.pending.set(id, { resolve, reject, channel });
        this.worker.postMessage({ id, op, payload, channel });
      });
    }
    async request(op, payload = {}, { channel = 'default' } = {}) {
      if (this.closed) throw new Error('本次工作階段已結束，請重新整理並匯入。');
      if (this.failed) await this.start();
      await this.ready;
      return this.send(op, payload, channel);
    }
    cancel(channel) {
      const ids = [];
      for (const [id, p] of this.pending) if (p.channel === channel) { ids.push(id); p.reject(abortError()); this.pending.delete(id); }
      if (ids.length) this.worker.postMessage({ op: 'cancel', ids });
      if (['import', 'export'].includes(channel) && ids.length) {
        this.worker.terminate(); this.fail(abortError()); this.start().catch(() => {});
      }
    }
    async dispose() {
      this.closed = true; this.worker?.terminate(); this.fail(abortError());
      root.removeEventListener('pagehide', this.pagehide);
      try { await deleteDatabase(PREFIX + this.sessionId); } finally { this.release?.(); }
    }
  }
  root.PhoneDatasetClient = { create: async options => {
    const client = new Client(options?.onProgress);
    try { return await client.initialize(); }
    catch (error) { client.worker?.terminate(); client.release?.(); throw error; }
  } };
})(globalThis);
