import React, { useEffect, useLayoutEffect, useRef, useState } from "react";

export function Dropdown({
  button, children, className, panelClassName, title, disabled, fixed,
}: {
  button: React.ReactNode;
  children: React.ReactNode | ((close: () => void) => React.ReactNode);
  className?: string;
  panelClassName?: string;
  title?: string;
  disabled?: boolean;
  /** Position the panel with fixed coordinates — escapes scrollable ancestors (settings pane). */
  fixed?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [flip, setFlip] = useState({ x: false, y: false });
  const [fixedPos, setFixedPos] = useState<{ top: number; right: number } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: PointerEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") setOpen(false); };
    document.addEventListener("pointerdown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("pointerdown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  // keep the panel inside the viewport: flip up/left when it would overflow
  useLayoutEffect(() => {
    if (!open || !panelRef.current) return;
    const r = panelRef.current.getBoundingClientRect();
    if (fixed) {
      // nudge a fixed panel back on screen instead of flipping
      if (r.bottom > window.innerHeight - 6 && fixedPos) {
        setFixedPos({ ...fixedPos, top: Math.max(8, window.innerHeight - 6 - r.height) });
      }
      return;
    }
    setFlip({
      x: r.right > window.innerWidth - 4,
      y: r.bottom > window.innerHeight - 4 && r.top - r.height > 4,
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const close = () => { setOpen(false); setFlip({ x: false, y: false }); setFixedPos(null); };

  const openPanel = () => {
    if (fixed && ref.current) {
      const r = ref.current.getBoundingClientRect();
      setFixedPos({ top: r.bottom + 3, right: Math.max(8, window.innerWidth - r.right) });
    }
    setOpen(true);
  };

  return (
    <div className={`dd ${className ?? ""}`} ref={ref}>
      <button
        type="button"
        className={`dd-btn ${open ? "active" : ""}`}
        title={title}
        disabled={disabled}
        onClick={() => (open ? close() : openPanel())}
      >
        {button}
      </button>
      {open && (
        <div
          ref={panelRef}
          className={`dd-panel ${panelClassName ?? ""} ${flip.x ? "flip-x" : ""} ${flip.y ? "up" : ""}`}
          style={fixed && fixedPos ? { position: "fixed", top: fixedPos.top, right: fixedPos.right, left: "auto", zIndex: 300 } : undefined}
        >
          {typeof children === "function" ? children(close) : children}
        </div>
      )}
    </div>
  );
}
