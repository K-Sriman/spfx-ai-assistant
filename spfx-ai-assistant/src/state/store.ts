import { create } from "zustand";

// ─── Panel State Store (Zustand) ──────────────────────────────────────────────
// Chat history lives in AssistantPanel local state + sessionStorage.
// This store only tracks the panel open/close status globally.

interface PanelState {
  isPanelOpen: boolean;
  openPanel: () => void;
  closePanel: () => void;
  togglePanel: () => void;
}

export const usePanelStore = create<PanelState>()((set) => ({
  isPanelOpen: false,
  openPanel: () => set({ isPanelOpen: true }),
  closePanel: () => set({ isPanelOpen: false }),
  togglePanel: () => set((s) => ({ isPanelOpen: !s.isPanelOpen })),
}));
