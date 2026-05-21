import type { Movement, SyncStatus } from "./types";
import { getWebhookUrl, postMovement } from "./webhook";

type SyncListener = (movements: Movement[]) => void;

const BACKOFF = [1500, 4000, 10000, 30000];

class SyncEngine {
  private movements: Movement[] = [];
  private listeners = new Set<SyncListener>();
  private running = false;
  private online = typeof navigator === "undefined" ? true : navigator.onLine;

  constructor() {
    if (typeof window !== "undefined") {
      window.addEventListener("online", () => {
        this.online = true;
        this.run();
      });
      window.addEventListener("offline", () => {
        this.online = false;
      });
    }
  }

  setMovements(list: Movement[]) {
    this.movements = list;
  }

  subscribe(fn: SyncListener) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  private emit() {
    this.listeners.forEach((fn) => fn([...this.movements]));
  }

  private patch(id: string, patch: Partial<Movement>) {
    this.movements = this.movements.map((m) =>
      m.id === id ? { ...m, ...patch } : m,
    );
    this.emit();
  }

  enqueue(m: Movement) {
    this.movements = [m, ...this.movements];
    this.emit();
    this.run();
  }

  async run() {
    if (this.running) return;
    const url = getWebhookUrl();
    if (!url) {
      this.movements = this.movements.map((m) =>
        m._sync === "queued" || m._sync === "sending"
          ? { ...m, _sync: "local" as SyncStatus }
          : m,
      );
      this.emit();
      return;
    }
    if (!this.online) return;

    this.running = true;
    try {
      while (true) {
        const next = this.movements.find(
          (m) => m._sync === "queued" || m._sync === "local",
        );
        if (!next) break;

        this.patch(next.id, { _sync: "sending" });
        const result = await postMovement(url, next);
        if (result.ok) {
          this.patch(next.id, { _sync: "ok", _error: undefined });
          continue;
        }

        const attempts = next._attempts + 1;
        const delay = BACKOFF[Math.min(attempts - 1, BACKOFF.length - 1)];
        this.patch(next.id, {
          _sync: "error",
          _attempts: attempts,
          _error: result.error,
        });
        if (attempts >= 4) break;
        await new Promise((r) => setTimeout(r, delay));
        this.patch(next.id, { _sync: "queued" });
      }
    } finally {
      this.running = false;
    }
  }

  retry(id: string) {
    this.patch(id, { _sync: "queued", _attempts: 0, _error: undefined });
    this.run();
  }

  retryAll() {
    this.movements = this.movements.map((m) =>
      m._sync === "error"
        ? { ...m, _sync: "queued" as SyncStatus, _attempts: 0, _error: undefined }
        : m,
    );
    this.emit();
    this.run();
  }
}

export const syncEngine = new SyncEngine();
