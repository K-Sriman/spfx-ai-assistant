import * as React from "react";
import type { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { usePanelStore } from "../../../state/store";
import AssistantPanel from "./AssistantPanel";
import { trackEvent } from "../../../utils/telemetry";
import styles from "./styles.module.scss";

// ─── AssistantLauncher ────────────────────────────────────────────────────────
// Renders the FAB and the slide-in panel host.
// Panel is kept mounted (not unmounted on close) so sessionStorage chat
// history is preserved across opens without re-loading.

interface AssistantLauncherProps {
  spContext: ApplicationCustomizerContext;
}

const PANEL_W = 440;

const AssistantLauncher: React.FC<AssistantLauncherProps> = ({ spContext }) => {
  const { isPanelOpen, openPanel, closePanel } = usePanelStore();
  const [everOpened, setEverOpened] = React.useState(false);

  function open(): void {
    setEverOpened(true);
    openPanel();
    trackEvent({ name: "panel_opened" });
  }

  function close(): void {
    closePanel();
    trackEvent({ name: "panel_closed" });
  }

  function toggle(): void {
    isPanelOpen ? close() : open();
  }

  // Keyboard shortcut: Escape to close
  React.useEffect(() => {
    function onKey(e: KeyboardEvent): void {
      if (e.key === "Escape" && isPanelOpen) close();
    }
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [isPanelOpen]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <>
      {/* ── Floating Action Button ── */}
      <button
        className={styles.fab}
        onClick={toggle}
        aria-label={isPanelOpen ? "Close AI assistant" : "Open AI assistant"}
        aria-expanded={isPanelOpen}
        title="AI Assistant"
      >
        {isPanelOpen ? (
          // Close icon
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none"
            stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
            <line x1="18" y1="6" x2="6" y2="18" />
            <line x1="6" y1="6" x2="18" y2="18" />
          </svg>
        ) : (
          // Spark icon
          <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor">
            <path d="M12 2L13.8 9.2L21 10.5L13.8 11.8L12 19L10.2 11.8L3 10.5L10.2 9.2Z" />
            <circle cx="20" cy="4" r="1.8" opacity="0.7" />
            <circle cx="4" cy="19" r="1.2" opacity="0.5" />
          </svg>
        )}
      </button>

      {/* ── Backdrop ── */}
      {isPanelOpen && (
        <div
          onClick={close}
          aria-hidden="true"
          style={{
            position: "fixed",
            inset: 0,
            zIndex: 9998,
            background: "rgba(0,0,0,0.18)",
            backdropFilter: "blur(2px)",
            WebkitBackdropFilter: "blur(2px)",
          }}
        />
      )}

      {/* ── Slide-in panel ── */}
      <div
        style={{
          position: "fixed",
          top: 0,
          right: 0,
          bottom: 0,
          width: `${PANEL_W}px`,
          maxWidth: "100vw",
          zIndex: 9999,
          transform: isPanelOpen ? "translateX(0)" : `translateX(${PANEL_W}px)`,
          transition: "transform 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
        }}
        aria-hidden={!isPanelOpen}
      >
        {/* Only render panel once it has been opened (lazy) */}
        {everOpened && (
          <AssistantPanel spContext={spContext} onClose={close} />
        )}
      </div>
    </>
  );
};

export default AssistantLauncher;
