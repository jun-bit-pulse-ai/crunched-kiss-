import type { CellValue } from "./excelPolicy";

/** Snapshot of a range's values before a write, so it can be restored on undo. */
export type Snapshot = {
  sheet: string;
  address: string;
  values: CellValue[][];
};

export const MAX_UNDO_DEPTH = 10;

let stack: Snapshot[] = [];
const listeners = new Set<() => void>();

function notify(): void {
  for (const listener of listeners) {
    listener();
  }
}

/**
 * Subscribe to changes so React can stay in sync via `useSyncExternalStore`
 * instead of a manually-bumped re-render counter. Returns an unsubscribe function.
 */
export function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** Push a snapshot, evicting the oldest one past MAX_UNDO_DEPTH. */
export function pushSnapshot(snapshot: Snapshot): void {
  stack.push(snapshot);
  if (stack.length > MAX_UNDO_DEPTH) {
    stack.shift();
  }
  notify();
}

/** Pop and return the most recent snapshot, or undefined if the stack is empty. */
export function popSnapshot(): Snapshot | undefined {
  const snapshot = stack.pop();
  notify();
  return snapshot;
}

export function canUndo(): boolean {
  return stack.length > 0;
}

export function undoDepth(): number {
  return stack.length;
}

export function clearUndoStack(): void {
  stack = [];
  notify();
}
