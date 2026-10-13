// What is being dragged, shared between the field list (drag source) and the areas
// (drag source and drop target). Kept in a module variable because the browser hides
// the dragged data from `dragover` handlers. Drag-and-drop is a shortcut only: every
// move it can make is also in the chip menu and the field checkboxes.

import type { DragEvent } from "react";

export interface DragPayload {
  kind: "measure" | "hierarchy";
  /** Unique name of the measure or hierarchy. */
  name: string;
  caption: string;
  /** For a hierarchy: its All member. */
  allMember: string | null;
}

let current: DragPayload | null = null;

export function startDrag(event: DragEvent, payload: DragPayload): void {
  current = payload;
  try {
    // Safari does not start a drag unless some data is set.
    event.dataTransfer.setData("text/plain", payload.caption);
    event.dataTransfer.effectAllowed = "move";
  } catch {
    // Some web views restrict dataTransfer; the module variable is enough.
  }
}

export function endDrag(): void {
  current = null;
}

export function dragged(): DragPayload | null {
  return current;
}
