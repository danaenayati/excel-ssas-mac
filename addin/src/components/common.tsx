// Small building blocks: icons, spinner, error note, copy button, in-pane dialog.
// The pane runs in a web view where window.alert/confirm/prompt are blocked, so every
// question and every error is drawn here, inside the pane.

import { useEffect, useRef, useState, type ReactNode } from "react";
import type { UiError } from "../state/runner";
import { RESTART_COMMAND, copyText, cx } from "../util";

type IconName =
  | "chevron-right"
  | "chevron-down"
  | "chevron-left"
  | "measure"
  | "hierarchy"
  | "attribute"
  | "table"
  | "folder"
  | "funnel"
  | "close"
  | "search"
  | "more"
  | "sort"
  | "trash";

const PATHS: Record<IconName, ReactNode> = {
  "chevron-right": <path d="M6 3.5 10.5 8 6 12.5" />,
  "chevron-down": <path d="M3.5 6 8 10.5 12.5 6" />,
  "chevron-left": <path d="M10 3.5 5.5 8 10 12.5" />,
  measure: <path d="M11.5 4.5v-1h-7L8.5 8l-4 4.5h7v-1" />,
  hierarchy: (
    <>
      <rect x="2.5" y="2.5" width="5" height="3" rx="0.5" />
      <rect x="8.5" y="6.5" width="5" height="3" rx="0.5" />
      <rect x="8.5" y="10.5" width="5" height="3" rx="0.5" />
      <path d="M5 5.5V12h3.5M5 8h3.5" />
    </>
  ),
  attribute: (
    <>
      <rect x="2.5" y="3.5" width="11" height="9" rx="1" />
      <path d="M2.5 6.5h11M6.5 6.5v6" />
    </>
  ),
  table: (
    <>
      <rect x="2.5" y="2.5" width="11" height="11" rx="1" />
      <path d="M2.5 6h11M2.5 9.5h11M6.5 6v7.5" />
    </>
  ),
  folder: <path d="M2.5 4.5a1 1 0 0 1 1-1h3l1.5 1.5h4.5a1 1 0 0 1 1 1v5.5a1 1 0 0 1-1 1h-9a1 1 0 0 1-1-1z" />,
  funnel: <path d="M2.5 3.5h11L9.5 8.5v4l-3 1v-5z" />,
  close: <path d="M4 4l8 8M12 4l-8 8" />,
  search: (
    <>
      <circle cx="7" cy="7" r="4" />
      <path d="M10 10l3.5 3.5" />
    </>
  ),
  more: (
    <>
      <circle cx="3.5" cy="8" r="0.6" />
      <circle cx="8" cy="8" r="0.6" />
      <circle cx="12.5" cy="8" r="0.6" />
    </>
  ),
  trash: <path d="M3 4.5h10M6.5 4.5v-2h3v2M4.5 4.5l.5 9h6l.5-9M6.75 7v4M9.25 7v4" />,
  sort: <path d="M5 3v10M2.5 10.5 5 13l2.5-2.5M11 13V3M8.5 5.5 11 3l2.5 2.5" />,
};

export function Icon({ name, className }: { name: IconName; className?: string }) {
  return (
    <svg
      className={cx("icon", className)}
      viewBox="0 0 16 16"
      width="14"
      height="14"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.3"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      {PATHS[name]}
    </svg>
  );
}

export function Spinner({ label }: { label?: string }) {
  return <span className="spinner" role="status" aria-label={label ?? "Working"} />;
}

export function CopyButton({ text, label = "Copy" }: { text: string; label?: string }) {
  const [state, setState] = useState<"idle" | "copied" | "failed">("idle");
  const timer = useRef<number | undefined>(undefined);
  useEffect(() => () => window.clearTimeout(timer.current), []);
  return (
    <button
      type="button"
      className="button small"
      onClick={async () => {
        const ok = await copyText(text);
        setState(ok ? "copied" : "failed");
        window.clearTimeout(timer.current);
        timer.current = window.setTimeout(() => setState("idle"), 2000);
      }}
    >
      {state === "copied" ? "Copied" : state === "failed" ? "Select and press ⌘C" : label}
    </button>
  );
}

/**
 * A command to copy. It wraps only at its spaces (never after a hyphen, which would
 * read as a different command), and selecting it by hand gives exactly `text`.
 */
export function CodeBlock({ text }: { text: string }) {
  const words = text.split(" ");
  return (
    <div className="code-block">
      <code>
        {words.map((word, index) => (
          <span key={index}>
            <span className="nowrap">{word}</span>
            {index < words.length - 1 ? " " : ""}
          </span>
        ))}
      </code>
      <CopyButton text={text} />
    </div>
  );
}

interface ErrorNoteProps {
  error: UiError;
  /** Shown as a button next to the message, e.g. "Try again". */
  action?: { label: string; run: () => void };
  onDismiss?: () => void;
}

/**
 * An error in plain words. Technical detail stays behind a "Details" disclosure. When
 * the helper itself cannot be reached, shows the command that restarts it.
 */
export function ErrorNote({ error, action, onDismiss }: ErrorNoteProps) {
  return (
    <div className="note error" role="alert">
      <div className="note-head">
        <span className="note-text">{error.message}</span>
        {onDismiss && (
          <button type="button" className="icon-button" onClick={onDismiss} aria-label="Dismiss">
            <Icon name="close" />
          </button>
        )}
      </div>
      {error.helperDown && (
        <div className="note-body">
          <p>To start it, open Terminal and run:</p>
          <CodeBlock text={RESTART_COMMAND} />
        </div>
      )}
      {error.detail && !error.helperDown && (
        <details className="details">
          <summary>Details</summary>
          <pre>{error.detail}</pre>
        </details>
      )}
      {action && (
        <div className="note-actions">
          <button type="button" className="button small" onClick={action.run}>
            {action.label}
          </button>
        </div>
      )}
    </div>
  );
}

interface DialogProps {
  title: string;
  children: ReactNode;
  /** Buttons, left to right. The last one is where focus starts unless `focus` says otherwise. */
  buttons: { label: string; run: () => void; primary?: boolean; danger?: boolean; focus?: boolean }[];
  /** Escape, or a click on the backdrop. */
  onCancel: () => void;
  wide?: boolean;
}

/** A question asked inside the pane, over everything else. */
export function Dialog({ title, children, buttons, onCancel, wide }: DialogProps) {
  const box = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const root = box.current;
    if (!root) return;
    const preferred = root.querySelector<HTMLElement>("[data-autofocus]");
    (preferred ?? root).focus();
  }, []);
  return (
    <div
      className="backdrop"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onCancel();
      }}
    >
      <div
        ref={box}
        className={cx("dialog", wide && "wide")}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        tabIndex={-1}
        onKeyDown={(event) => {
          if (event.key === "Escape") {
            event.stopPropagation();
            onCancel();
          }
        }}
      >
        <h2 className="dialog-title">{title}</h2>
        <div className="dialog-body">{children}</div>
        <div className="dialog-buttons">
          {buttons.map((button, index) => {
            const focus = button.focus ?? (index === buttons.length - 1 && !buttons.some((b) => b.focus));
            return (
              <button
                key={button.label}
                type="button"
                className={cx("button", button.primary && "primary", button.danger && "danger")}
                onClick={button.run}
                {...(focus ? { "data-autofocus": true } : {})}
              >
                {button.label}
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/** A checkbox that can also show the "some of the children" state. */
export function TriCheckbox(props: {
  state: "on" | "off" | "partial";
  onChange: () => void;
  label: string;
  disabled?: boolean;
}) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    if (ref.current) ref.current.indeterminate = props.state === "partial";
  }, [props.state]);
  return (
    <input
      ref={ref}
      type="checkbox"
      checked={props.state === "on"}
      onChange={props.onChange}
      aria-label={props.label}
      aria-checked={props.state === "partial" ? "mixed" : props.state === "on"}
      disabled={props.disabled}
    />
  );
}
