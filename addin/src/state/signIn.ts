// "Sign in with" in the connection form: which kinds a server allows, the default, and
// the words for each. The helper applies the same rules (helper/src/cubebridge/ssas/
// auth.py) and names these same labels in its error messages, so an error can say
// exactly which choice to pick.
//
// Windows Excel's wording, for comparison: "Use Windows Authentication" is what both
// NTLM and Kerberos do here (Windows sends the signed-in user; a Mac uses the saved
// user name and password, or a Kerberos ticket it already has), and "Use the following
// User Name and Password" with a web address is Basic.

import type { SignIn } from "../types";

export type ServerKind = "name" | "http" | "https";

export function serverKind(server: string): ServerKind {
  const text = server.trim().toLowerCase();
  if (text.startsWith("https://")) return "https";
  if (text.startsWith("http://")) return "http";
  return "name";
}

export const SIGN_IN_LABELS: Record<SignIn, string> = {
  ntlm: "User name and password (NTLM)",
  kerberos: "Windows Authentication (Kerberos)",
  basic: "User name and password over HTTPS (Basic)",
};

const SIGN_IN_HINTS: Record<SignIn, string> = {
  ntlm: "Works with most company servers. The password itself is never sent.",
  kerberos:
    "For servers that accept only Kerberos. Uses this Mac’s Kerberos ticket, or gets one with the password below.",
  basic: "For web addresses set up for Basic sign-in. The password is sent, protected by https.",
};

export interface SignInOption {
  value: SignIn;
  label: string;
  /** One sentence on when to choose it. */
  hint: string;
}

/** The kinds this server allows, the default first. Basic only for https:// addresses. */
export function signInOptions(server: string): SignInOption[] {
  const kinds: SignIn[] = serverKind(server) === "https" ? ["ntlm", "kerberos", "basic"] : ["ntlm", "kerberos"];
  return kinds.map((value) => ({ value, label: SIGN_IN_LABELS[value], hint: SIGN_IN_HINTS[value] }));
}

/**
 * NTLM for every kind of server: it works with IIS's Windows Authentication (the set-up
 * Microsoft recommends for msmdpump.dll) as with the TCP port, and it never sends the
 * password itself. A server that accepts only Basic says so, naming that choice.
 */
export function defaultSignIn(): SignIn {
  return "ntlm";
}

/** The choice to keep after the server changed: one the new server does not allow goes back to the default. */
export function keepSignIn(server: string, current: SignIn): SignIn {
  return signInOptions(server).some((option) => option.value === current) ? current : defaultSignIn();
}

/** Why this server and sign-in cannot go together, or null. */
export function signInProblem(server: string, signIn: SignIn): string | null {
  if (signIn !== "basic") return null;
  const kind = serverKind(server);
  if (kind === "https") return null;
  if (kind === "http") return "Basic sends the password itself, so it needs an https:// address.";
  return "Basic is only for a web address (https://…/msmdpump.dll).";
}

/** A word for the connection list; nothing for NTLM, the usual kind. */
export function signInTag(signIn: SignIn | undefined): string {
  if (signIn === "kerberos") return "Kerberos";
  if (signIn === "basic") return "Basic";
  return "";
}
