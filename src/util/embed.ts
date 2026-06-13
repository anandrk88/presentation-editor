/**
 * Host-application integration (the iframe embed + document bridge surface).
 *
 * Two ways to integrate:
 *  1. URL parameters — point an iframe (or tab) at the editor:
 *       /?file=<encoded pptx url>&title=My+Deck&saveUrl=<presigned PUT url>
 *     The editor fetches the file, and Save uploads back to saveUrl
 *     (falls back to a local download when no saveUrl is given).
 *  2. postMessage bridge — add &embed=1&parentOrigin=<https://your-app> and
 *     drive the editor from the host page:
 *       editor.contentWindow.postMessage({ type: "pe:load", url, title }, EDITOR_ORIGIN)
 *       editor.contentWindow.postMessage({ type: "pe:save" }, EDITOR_ORIGIN)
 *     Events posted back to the host (always tagged with source: "presentation-editor"):
 *       pe:ready                       — bridge is listening
 *       pe:loaded  { title, slideCount }
 *       pe:dirty   { dirty }           — unsaved-changes indicator
 *       pe:document{ data: ArrayBuffer, fileName } — reply to pe:save (transferable)
 *       pe:saved   { via: "upload" | "message" }
 *       pe:error   { message }
 *
 * Security: inbound commands are accepted ONLY from parentOrigin, and all
 * outbound messages target parentOrigin explicitly. Without parentOrigin the
 * bridge stays disabled even if embed=1.
 */

export interface EmbedConfig {
  fileUrl?: string;
  title?: string;
  saveUrl?: string;
  embed: boolean;
  parentOrigin?: string;
}

/** Pure parser (testable): read an embed config from a query string + base URL. */
export function parseEmbedQuery(search: string, baseHref: string): EmbedConfig {
  const q = new URLSearchParams(search);
  // only http(s) sources/sinks survive; anything else (javascript:, ftp:, …) is dropped
  const httpish = (u: string | null): string | undefined => {
    if (!u) return undefined;
    try {
      const abs = new URL(u, baseHref).href;
      return /^https?:\/\//i.test(abs) ? abs : undefined;
    } catch { return undefined; }
  };
  return {
    fileUrl: httpish(q.get("file")),
    title: q.get("title") ?? undefined,
    saveUrl: httpish(q.get("saveUrl")),
    embed: q.get("embed") === "1",
    parentOrigin: q.get("parentOrigin") ?? undefined,
  };
}

export function embedConfig(): EmbedConfig {
  if (typeof location === "undefined") return { embed: false };
  return parseEmbedQuery(location.search, location.href);
}

export interface BridgeHandlers {
  /** Open a presentation from raw bytes. */
  load: (buf: ArrayBuffer, title: string) => Promise<void>;
  /** Export the current presentation as a .pptx blob. */
  exportPptx: () => Promise<Blob>;
  /** Current document title (for the export file name). */
  title: () => string;
}

let post: ((msg: Record<string, unknown>, transfer?: Transferable[]) => void) | null = null;

/** Send an event to the host page (no-op when not embedded). */
export function notifyHost(msg: Record<string, unknown>, transfer?: Transferable[]) {
  post?.(msg, transfer);
}

/** Upload a blob to a presigned PUT url (R2 / S3 style). */
export async function uploadTo(saveUrl: string, blob: Blob): Promise<void> {
  const res = await fetch(saveUrl, {
    method: "PUT",
    body: blob,
    headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
  });
  if (!res.ok) throw new Error(`Upload failed: HTTP ${res.status}`);
}

export function initEmbedBridge(cfg: EmbedConfig, handlers: BridgeHandlers) {
  if (!cfg.embed || !cfg.parentOrigin || typeof window === "undefined" || window.parent === window) return;
  const origin = cfg.parentOrigin;
  post = (msg, transfer) =>
    window.parent.postMessage({ source: "presentation-editor", ...msg }, origin, transfer ?? []);

  window.addEventListener("message", async e => {
    if (e.origin !== origin || !e.data || typeof e.data.type !== "string") return;
    try {
      switch (e.data.type) {
        case "pe:load": {
          let buf: ArrayBuffer | undefined = e.data.data instanceof ArrayBuffer ? e.data.data : undefined;
          if (!buf && typeof e.data.url === "string") {
            const res = await fetch(e.data.url);
            if (!res.ok) throw new Error(`Could not fetch document: HTTP ${res.status}`);
            buf = await res.arrayBuffer();
          }
          if (!buf) throw new Error("pe:load needs a url or an ArrayBuffer in data");
          await handlers.load(buf, typeof e.data.title === "string" ? e.data.title : "Presentation");
          break;
        }
        case "pe:save": {
          const blob = await handlers.exportPptx();
          const data = await blob.arrayBuffer();
          post?.({ type: "pe:document", data, fileName: `${handlers.title() || "Presentation"}.pptx` }, [data]);
          if (cfg.saveUrl) {
            await uploadTo(cfg.saveUrl, blob);
            post?.({ type: "pe:saved", via: "upload" });
          } else {
            post?.({ type: "pe:saved", via: "message" });
          }
          break;
        }
      }
    } catch (err) {
      post?.({ type: "pe:error", message: (err as Error).message });
    }
  });

  post({ type: "pe:ready", version: typeof __APP_VERSION__ !== "undefined" ? __APP_VERSION__ : "dev" });
}
