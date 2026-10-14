// "Add connection": server, how to sign in, username, password, name. The helper signs
// in to the server before it saves anything, so a wrong password is reported here, at
// once. The password goes to the helper (which keeps it in the macOS Keychain) and is
// not kept by the pane.

import { useState, type FormEvent } from "react";
import type { Api } from "../api";
import { toUiError, type UiError } from "../state/runner";
import { defaultSignIn, keepSignIn, signInOptions, signInProblem } from "../state/signIn";
import type { Connection, NewConnection, SignIn } from "../types";
import { ErrorNote, Spinner } from "./common";

interface Props {
  api: Api;
  onAdded: (connection: Connection) => void;
  onCancel: () => void;
}

export interface FormFields {
  name: string;
  server: string;
  signIn: SignIn;
  username: string;
  password: string;
}

/** What the helper is sent: trimmed, without an empty name. The password as typed. */
export function newConnection(fields: FormFields): NewConnection {
  return {
    name: fields.name.trim() || undefined,
    server: fields.server.trim(),
    username: fields.username.trim(),
    password: fields.password,
    signIn: fields.signIn,
  };
}

export function formReady(fields: FormFields): boolean {
  return (
    fields.server.trim() !== "" &&
    fields.username.trim() !== "" &&
    fields.password !== "" &&
    signInProblem(fields.server, fields.signIn) === null
  );
}

export function ConnectionForm({ api, onAdded, onCancel }: Props) {
  const [name, setName] = useState("");
  const [server, setServer] = useState("");
  const [signIn, setSignIn] = useState<SignIn>(defaultSignIn());
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<UiError | null>(null);

  const fields: FormFields = { name, server, signIn, username, password };
  const ready = formReady(fields);
  const options = signInOptions(server);
  const chosen = options.find((option) => option.value === signIn) ?? options[0];

  const changeServer = (value: string) => {
    setServer(value);
    // Basic is offered for https:// addresses only; leaving one takes the choice back.
    setSignIn((current) => keepSignIn(value, current));
  };

  const submit = async (event: FormEvent) => {
    event.preventDefault();
    if (!ready || busy) return;
    setBusy(true);
    setError(null);
    try {
      const connection = await api.addConnection(newConnection(fields));
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
          onChange={(e) => changeServer(e.target.value)}
          placeholder="server, server\instance or https://…/msmdpump.dll"
          disabled={busy}
          autoFocus
          {...plain}
        />
      </label>
      <fieldset className="field choice-group" disabled={busy}>
        <legend className="field-label">Sign in with</legend>
        {options.map((option) => (
          <label key={option.value} className="choice">
            <input
              type="radio"
              name="sign-in"
              value={option.value}
              checked={option.value === chosen.value}
              onChange={() => setSignIn(option.value)}
            />
            <span>{option.label}</span>
          </label>
        ))}
        <span className="hint">{chosen.hint}</span>
      </fieldset>
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
