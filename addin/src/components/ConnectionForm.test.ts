import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import type { Api } from "../api";
import { createMockApi } from "../dev/mockApi";
import { ApiFailure } from "../apiErrors";
import { ConnectionForm, formReady, newConnection, type FormFields } from "./ConnectionForm";
import { describe as describeConnection } from "./Connections";

const FIELDS: FormFields = {
  name: "  ",
  server: "  https://bi.corp.local/olap/msmdpump.dll ",
  signIn: "basic",
  username: " CORP\\dana ",
  password: " pw with spaces ",
};

describe("the connection form", () => {
  it("shows Sign in with, NTLM chosen, and no Basic for an empty server", () => {
    const html = renderToStaticMarkup(
      createElement(ConnectionForm, { api: {} as Api, onAdded: () => {}, onCancel: () => {} }),
    );
    const radio = (value: string) => html.match(new RegExp(`<input[^>]*value="${value}"[^>]*/>`))?.[0] ?? "";
    expect(html).toContain("<legend");
    expect(html).toContain("Sign in with");
    expect(radio("ntlm")).toContain('type="radio"');
    expect(radio("ntlm")).toContain("checked");
    expect(radio("kerberos")).toContain('type="radio"');
    expect(radio("kerberos")).not.toContain("checked");
    expect(radio("basic")).toBe("");
    expect(html).toContain("The password itself is never sent.");
    // The choice comes between the server and the user name, as in Windows' dialog.
    expect(html.indexOf("Server")).toBeLessThan(html.indexOf("Sign in with"));
    expect(html.indexOf("Sign in with")).toBeLessThan(html.indexOf("Username"));
  });

  it("sends the sign-in kind with the trimmed fields and the password as typed", () => {
    expect(newConnection(FIELDS)).toEqual({
      name: undefined,
      server: "https://bi.corp.local/olap/msmdpump.dll",
      username: "CORP\\dana",
      password: " pw with spaces ",
      signIn: "basic",
    });
  });

  it("is ready only with every field and a sign-in the server allows", () => {
    expect(formReady(FIELDS)).toBe(true);
    expect(formReady({ ...FIELDS, password: "" })).toBe(false);
    expect(formReady({ ...FIELDS, server: "http://bi/olap/msmdpump.dll" })).toBe(false);
    expect(formReady({ ...FIELDS, server: "bi-server", signIn: "kerberos" })).toBe(true);
  });
});

describe("the connection list", () => {
  it("adds Kerberos or Basic after the user name, and nothing for NTLM", () => {
    const base = { id: "a", name: "BI", kind: "ssas" as const, server: "bi", username: "CORP\\dana" };
    expect(describeConnection(base)).toBe("bi · CORP\\dana");
    expect(describeConnection({ ...base, signIn: "ntlm" })).toBe("bi · CORP\\dana");
    expect(describeConnection({ ...base, signIn: "kerberos" })).toBe("bi · CORP\\dana · Kerberos");
    expect(describeConnection({ ...base, server: "https://bi/x", signIn: "basic" })).toBe("https://bi/x · CORP\\dana · Basic");
  });
});

describe("the development stand-in for the helper", () => {
  it("keeps the sign-in kind and refuses Basic without https, as the helper does", async () => {
    const api = createMockApi({ delayMs: 0 });
    const made = await api.addConnection({ server: "bi", username: "CORP\\dana", password: "pw", signIn: "kerberos" });
    expect(made.signIn).toBe("kerberos");
    expect((await api.addConnection({ server: "bi", username: "u", password: "pw" })).signIn).toBe("ntlm");
    const refused = await api.addConnection({ server: "bi", username: "u", password: "pw", signIn: "basic" }).catch((e) => e);
    expect(refused).toBeInstanceOf(ApiFailure);
    expect((refused as ApiFailure).error.category).toBe("bad-request");
  });
});
