// Screen 1: the saved connections. Choosing one opens the Explorer on it. The Sample
// cube is always there and needs no server.

import { useState } from "react";
import type { Api } from "../api";
import { toUiError, type UiError } from "../state/runner";
import { signInTag } from "../state/signIn";
import type { Connection } from "../types";
import { cx } from "../util";
import { ConnectionForm } from "./ConnectionForm";
import { ErrorNote, Icon, Spinner } from "./common";

interface Props {
  api: Api;
  connections: Connection[];
  /** Connection of the report being edited, marked in the list. */
  activeId: string | null;
  onChoose: (connection: Connection) => void;
  /** Called after the list changed on the helper (added or deleted). */
  onChanged: () => Promise<void>;
  /** Present when there is an Explorer to go back to. */
  onBack?: () => void;
}

export function describe(connection: Connection): string {
  if (connection.kind === "demo") return "Built-in sample data. No server needed.";
  return [connection.server, connection.username, signInTag(connection.signIn)].filter(Boolean).join(" · ");
}

export function Connections({ api, connections, activeId, onChoose, onChanged, onBack }: Props) {
  const [adding, setAdding] = useState(false);
  const [confirming, setConfirming] = useState<string | null>(null);
  const [deleting, setDeleting] = useState<string | null>(null);
  const [error, setError] = useState<UiError | null>(null);

  const remove = async (connection: Connection) => {
    setDeleting(connection.id);
    setError(null);
    try {
      await api.deleteConnection(connection.id);
      await onChanged();
    } catch (failure) {
      setError(toUiError(failure));
    } finally {
      setDeleting(null);
      setConfirming(null);
    }
  };

  return (
    <div className="screen">
      <header className="bar">
        {onBack && (
          <button type="button" className="icon-button" onClick={onBack} aria-label="Back to the report">
            <Icon name="chevron-left" />
          </button>
        )}
        <h1 className="bar-title">Connections</h1>
      </header>

      <div className="scroll">
        <ul className="conn-list">
          {connections.map((connection) => (
            <li key={connection.id} className={cx("conn", connection.id === activeId && "active")}>
              {confirming === connection.id ? (
                <div className="conn-confirm">
                  <span>Delete &ldquo;{connection.name}&rdquo;? Reports that use it will ask for another connection.</span>
                  <span className="conn-confirm-buttons">
                    {deleting === connection.id && <Spinner label="Deleting" />}
                    <button
                      type="button"
                      className="button small"
                      onClick={() => setConfirming(null)}
                      disabled={deleting !== null}
                      autoFocus
                    >
                      Cancel
                    </button>
                    <button
                      type="button"
                      className="button small danger"
                      onClick={() => remove(connection)}
                      disabled={deleting !== null}
                    >
                      Delete
                    </button>
                  </span>
                </div>
              ) : (
                <>
                  <button type="button" className="conn-main" onClick={() => onChoose(connection)}>
                    <span className="conn-name">{connection.name}</span>
                    <span className="conn-sub">{describe(connection)}</span>
                  </button>
                  {connection.kind !== "demo" && (
                    <button
                      type="button"
                      className="icon-button"
                      onClick={() => setConfirming(connection.id)}
                      aria-label={`Delete ${connection.name}`}
                      title="Delete"
                    >
                      <Icon name="trash" />
                    </button>
                  )}
                </>
              )}
            </li>
          ))}
        </ul>

        {error && <ErrorNote error={error} onDismiss={() => setError(null)} />}

        {adding ? (
          <section className="section">
            <h2 className="section-title">Add connection</h2>
            <ConnectionForm
              api={api}
              onCancel={() => setAdding(false)}
              onAdded={async (connection) => {
                setAdding(false);
                await onChanged();
                onChoose(connection);
              }}
            />
          </section>
        ) : (
          <div className="section">
            <button type="button" className="button" onClick={() => setAdding(true)}>
              Add connection…
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
