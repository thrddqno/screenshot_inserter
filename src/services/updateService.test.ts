import { afterEach, describe, expect, it, vi } from "vitest";
import {
  buildUpdateUrl,
  checkForAppUpdate,
  CURRENT_APP_VERSION,
} from "./updateService";

describe("app update service", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("requests same-origin version without cache", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ version: CURRENT_APP_VERSION }),
    });
    vi.stubGlobal("fetch", fetchMock);

    await expect(checkForAppUpdate()).resolves.toBeNull();
    const [url, options] = fetchMock.mock.calls[0];
    expect(url).toBeInstanceOf(URL);
    expect((url as URL).origin).toBe(window.location.origin);
    expect((url as URL).pathname).toMatch(/\/version\.json$/);
    expect(options).toMatchObject({
      cache: "no-store",
      credentials: "same-origin",
      redirect: "error",
    });
  });

  it("returns different valid deployed version", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ version: "abcdef123456" }),
    }));
    const expected = CURRENT_APP_VERSION === "abcdef123456" ? null : "abcdef123456";
    await expect(checkForAppUpdate()).resolves.toBe(expected);
  });

  it("rejects malformed and unsuccessful responses", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ version: "<script>" }),
    }));
    await expect(checkForAppUpdate()).rejects.toThrow("Update response is invalid");

    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: false, status: 503 }));
    await expect(checkForAppUpdate()).rejects.toThrow("HTTP 503");
  });

  it("builds same-page cache-busting URL and rejects invalid versions", () => {
    expect(buildUpdateUrl("abcdef123456", "https://example.com/taskpane.html?host=excel#old"))
      .toBe("https://example.com/taskpane.html?host=excel&version=abcdef123456");
    expect(() => buildUpdateUrl("../other-page")).toThrow("Update version is invalid");
  });
});
