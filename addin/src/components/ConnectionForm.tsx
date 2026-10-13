// "Add connection": name, server, username, password. The helper signs in to the
// server before it saves anything, so a wrong password is reported here, at once. The
// password goes to the helper (which keeps it in the macOS Keychain) and is not kept
// by the pane.

import { useState, type FormEvent } from "react";
import type { Api } from "../api";
import { toUiError, type UiError } from "../state/runner";
import type { Connection } from "../types";
import { ErrorNote, Spinner } from "./common";

interface Props {
  api: Api;
  onAdded: (connection: Connection) => void;
  onCancel: () => void;
}

export function ConnectionForm({ api, onAdded, onCancel }: Props) {
  const [name, setName] = useState("");
  const [server, setServer] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<UiError | null>(null);

  const ready = server.trim() !== "" && username.trim() !== "" && password !== "";

  const submit = async (event: FormEvent) => {
    event.preventDefault();
    if (!ready || busy) return;
    setBusy(true);
    setError(null);
    try {
      const connection = await api.addConnection({
        name: name.trim() || undefined,
        server: server.trim(),
        username: username.trim(),
        password,
      });
      setPassword("");
      onAdded(connection);
    } catch (failure) {
      setError(toUiError(failure));
      setBusy(false);
    }
  };

  // Server names and usernames must reach the helper exactly as typed.
  const plain = {
    autoCapitalize: "off",
    autoCorrect: "off",
    spellCheck: false,
    autoComplete: "off",
  } as const;

  return (
    <form className="form" onSubmit={submit} aria-label="Add connection">
      <label className="field">
        <span className="field-label">Server</span>
        <input
          className="input"
          value={server}
          onChange={(e) => setServer(e.target.value)}
          placeholder="server, server:port or server\instance"
          disabled={busy}
          autoFocus
          {...plain}
        />
      </label>
      <label className="field">
        <span className="field-label">Username</span>
        <input
          className="input"
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          placeholder="DOMAIN\user"
          disabled={busy}
          {...plain}
        />
      </label>
      <label className="field">
        <span className="field-label">Password</span>
        <input
          className="input"
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          disabled={busy}
          {...plain}
        />
      </label>
      <label className="field">
        <span className="field-label">Name</span>
        <input
          className="input"
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Optional, shown in the list"
          disabled={busy}
          {...plain}
        />
      </label>
      <p className="hint">
        Use the same server name and sign-in as in Excel on Windows. The password is kept in
        your Mac&rsquo;s Keychain, never in the workbook.
      </p>
      {error && <ErrorNote error={error} onDismiss={() => setError(null)} />}
      <div className="form-buttons">
        {busy && (
          <span className="busy-inline">
            <Spinner label="Testing the connection" /> Signing in to the server…
          </span>
        )}
        <button type="button" className="button" onClick={onCancel} disabled={busy}>
          Cancel
        </button>
        <button type="submit" className="button primary" disabled={!ready || busy}>
          Connect
        </button>
      </div>
    </form>
  );
}
