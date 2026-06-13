/** Injected at build time from package.json (vite define). */
declare const __APP_VERSION__: string;

interface Window {
  /** Public scripting API for host apps (same-origin). See src/util/api.ts. */
  presentationEditor?: import("./util/api").PresentationEditorApi;
}
