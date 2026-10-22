import { afterEach, describe, expect, it, vi } from "vitest";
import { IMAGE_LIMITS } from "../config/layoutProfiles";
import {
  detectImageMimeType,
  getImageFromSystemClipboard,
  ImageValidationError,
  validateImageBytes,
} from "./imageValidation";

const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]);
const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0xff, 0xd9]);

describe("image validation", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("detects PNG and JPEG signatures", () => {
    expect(detectImageMimeType(png)).toBe("image/png");
    expect(detectImageMimeType(jpeg)).toBe("image/jpeg");
  });

  it("rejects unsupported MIME types", () => {
    expect(() => validateImageBytes("image/gif", png.byteLength, png)).toThrow(ImageValidationError);
  });

  it("rejects spoofed MIME types", () => {
    expect(() => validateImageBytes("image/jpeg", png.byteLength, png)).toThrow(/does not match/);
  });

  it("rejects images over configured maximum", () => {
    expect(() => validateImageBytes("image/png", IMAGE_LIMITS.maximumBytes + 1, png)).toThrow(/10 MB/);
  });

  it("explains when copied clipboard content has no browser-readable image", async () => {
    vi.stubGlobal("navigator", {
      clipboard: {
        read: vi.fn(async () => [{ types: ["text/plain"], getType: vi.fn() }]),
      },
    });

    await expect(getImageFromSystemClipboard()).rejects.toThrow(
      "Clipboard contains no PNG or JPEG image",
    );
  });
});
