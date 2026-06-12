import { useSyncExternalStore } from "react";
import { store } from "./store";
import type { EditorState } from "./store";

export function useEditorState(): EditorState {
  return useSyncExternalStore(store.subscribe, store.getState);
}

/** Ribbon <-> text-edit-overlay coordination (commit before model-level paragraph ops). */
export const editorBus: {
  commitTextEdit: (() => void) | null;
  exec: (cmd: string, value?: string) => void;
} = {
  commitTextEdit: null,
  exec(cmd: string, value?: string) {
    document.execCommand(cmd, false, value);
  },
};
