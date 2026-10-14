import { describe, expect, it } from "vitest";
import {
  SIGN_IN_LABELS,
  defaultSignIn,
  keepSignIn,
  serverKind,
  signInOptions,
  signInProblem,
  signInTag,
} from "./signIn";

describe("serverKind", () => {
  it("tells server names from web addresses", () => {
    expect(serverKind("bi-server\\TAB")).toBe("name");
    expect(serverKind("bi-server:2383")).toBe("name");
    expect(serverKind("  HTTPS://bi/olap/msmdpump.dll")).toBe("https");
    expect(serverKind("http://bi/olap/msmdpump.dll")).toBe("http");
    expect(serverKind("")).toBe("name");
  });
});

describe("signInOptions", () => {
  it("offers NTLM first, then Kerberos, for a server name and for http://", () => {
    for (const server of ["", "bi-server", "http://bi/olap/msmdpump.dll"]) {
      expect(signInOptions(server).map((o) => o.value)).toEqual(["ntlm", "kerberos"]);
    }
  });

  it("adds Basic for https:// only", () => {
    expect(signInOptions("https://bi/olap/msmdpump.dll").map((o) => o.value)).toEqual(["ntlm", "kerberos", "basic"]);
  });

  it("labels each choice in the helper's words, with a hint", () => {
    const [ntlm, kerberos, basic] = signInOptions("https://bi/x");
    expect(ntlm.label).toBe("User name and password (NTLM)");
    expect(kerberos.label).toBe("Windows Authentication (Kerberos)");
    expect(basic.label).toBe("User name and password over HTTPS (Basic)");
    expect(kerberos.hint).toMatch(/Kerberos ticket/);
    expect(basic.hint).toMatch(/https/);
    expect(Object.keys(SIGN_IN_LABELS).sort()).toEqual(["basic", "kerberos", "ntlm"]);
  });
});

describe("defaults and changes", () => {
  it("defaults to NTLM", () => {
    expect(defaultSignIn()).toBe("ntlm");
  });

  it("keeps a choice the new server allows and drops Basic when the address is no longer https", () => {
    expect(keepSignIn("bi-server", "kerberos")).toBe("kerberos");
    expect(keepSignIn("https://bi/x", "basic")).toBe("basic");
    expect(keepSignIn("https://bi/x", "kerberos")).toBe("kerberos");
    expect(keepSignIn("http://bi/x", "basic")).toBe("ntlm");
    expect(keepSignIn("bi-server", "basic")).toBe("ntlm");
  });

  it("explains why Basic cannot go with anything but https", () => {
    expect(signInProblem("https://bi/x", "basic")).toBeNull();
    expect(signInProblem("bi-server", "ntlm")).toBeNull();
    expect(signInProblem("http://bi/x", "basic")).toMatch(/https:\/\//);
    expect(signInProblem("bi-server", "basic")).toMatch(/web address/);
  });

  it("tags the connection list for the less usual kinds only", () => {
    expect(signInTag(undefined)).toBe("");
    expect(signInTag("ntlm")).toBe("");
    expect(signInTag("kerberos")).toBe("Kerberos");
    expect(signInTag("basic")).toBe("Basic");
  });
});
