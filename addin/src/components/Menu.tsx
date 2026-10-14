// A small pop-up menu anchored to a button. Used for the chip menus and the report
// menu: the click path that drag-and-drop is only a shortcut for. An entry with a
// submenu ("Show Values As ▸") opens it in place of the menu, with a way back: in a
// pane 320 px wide there is no room to fly out sideways.

import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { cx } from "../util";

export type MenuEntry =
  | { label: string; run: () => void; disabled?: boolean; danger?: boolean; checked?: boolean }
  | { label: string; submenu: MenuEntry[]; disabled?: boolean }
  | "separator";

interface MenuProps {
  /** The element the menu hangs from. */
  anchor: HTMLElement;
  items: MenuEntry[];
  onClose: () => void;
  label: string;
}

const MARGIN = 6;

export function Menu({ anchor, items: topItems, onClose, label }: MenuProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState<{ left: number; top: number } | null>(null);
  /** The submenu shown instead of the top level, if any. */
  const [open, setOpen] = useState<{ label: string; items: MenuEntry[] } | null>(null);
  const items = open ? open.items : topItems;

  // Place it under the anchor, inside the pane; above when there is no room below.
  useLayoutEffect(() => {
    const menu = ref.current;
    if (!menu) return;
    const at = anchor.getBoundingClientRect();
    const pane = anchor.closest(".pane");
    const bounds = pane
      ? pane.getBoundingClientRect()
      : { left: 0, right: window.innerWidth, top: 0, bottom: window.innerHeight };
    const width = menu.offsetWidth;
    const height = menu.offsetHeight;
    let left = Math.min(at.left, bounds.right - width - MARGIN);
    left = Math.max(left, bounds.left + MARGIN);
    let top = at.bottom + 2;
    if (top + height > window.innerHeight - MARGIN) top = Math.max(MARGIN, at.top - height - 2);
    setPosition({ left, top });
  }, [anchor, items.length]);

  useEffect(() => {
    const first = ref.current?.querySelector<HTMLButtonElement>("button:not(:disabled)");
    if (first) first.focus();
  }, [position === null, open]);

  useEffect(() => {
    const outside = (event: Event) => {
      const target = event.target as Node | null;
      if (ref.current && target && ref.current.contains(target)) return;
      if (target && anchor.contains(target)) return; // the anchor toggles the menu itself
      onClose();
    };
    const close = () => onClose();
    document.addEventListener("mousedown", outside, true);
    document.addEventListener("scroll", outside, true);
    window.addEventListener("resize", close);
    window.addEventListener("blur", close);
    return () => {
      document.removeEventListener("mousedown", outside, true);
      document.removeEventListener("scroll", outside, true);
      window.removeEventListener("resize", close);
      window.removeEventListener("blur", close);
    };
  }, [anchor, onClose]);

  const move = (delta: number) => {
    const buttons = Array.from(
      ref.current?.querySelectorAll<HTMLButtonElement>("button:not(:disabled)") ?? [],
    );
    if (buttons.length === 0) return;
    const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const next = (index + delta + buttons.length) % buttons.length;
    buttons[next].focus();
  };

  return createPortal(
    <div
      ref={ref}
      className="menu"
      role="menu"
      aria-label={label}
      style={position ? { left: position.left, top: position.top } : { left: 0, top: 0, visibility: "hidden" }}
      onKeyDown={(event) => {
        if ((event.key === "Escape" || event.key === "ArrowLeft") && open) {
          event.preventDefault();
          setOpen(null);
        } else if (event.key === "Escape" || event.key === "Tab") {
          event.preventDefault();
          onClose();
          anchor.focus();
        } else if (event.key === "ArrowDown") {
          event.preventDefault();
          move(1);
        } else if (event.key === "ArrowUp") {
          event.preventDefault();
          move(-1);
        }
      }}
    >
      {open && (
        <>
          <button type="button" role="menuitem" className="menu-item menu-back" onClick={() => setOpen(null)}>
            <span className="menu-check" aria-hidden="true">
              ‹
            </span>
            {open.label}
          </button>
          <div className="menu-separator" role="separator" />
        </>
      )}
      {items.map((item, index) => {
        if (item === "separator") return <div key={`s${index}`} className="menu-separator" role="separator" />;
        if ("submenu" in item) {
          return (
            <button
              key={item.label}
              type="button"
              role="menuitem"
              aria-haspopup="menu"
              className="menu-item"
              disabled={item.disabled}
              onClick={() => setOpen({ label: item.label, items: item.submenu })}
              onKeyDown={(event) => {
                if (event.key === "ArrowRight") {
                  event.preventDefault();
                  event.stopPropagation();
                  setOpen({ label: item.label, items: item.submenu });
                }
              }}
            >
              <span className="menu-check" aria-hidden="true" />
              <span className="menu-text">{item.label}</span>
              <span className="menu-more" aria-hidden="true">
                ▸
              </span>
            </button>
          );
        }
        return (
          <button
            key={item.label}
            type="button"
            role={item.checked === undefined ? "menuitem" : "menuitemradio"}
            aria-checked={item.checked === undefined ? undefined : item.checked}
            className={cx("menu-item", item.danger && "danger")}
            disabled={item.disabled}
            onClick={() => {
              onClose();
              item.run();
            }}
          >
            <span className="menu-check" aria-hidden="true">
              {item.checked ? "✓" : ""}
            </span>
            {item.label}
          </button>
        );
      })}
    </div>,
    document.body,
  );
}
