import { describe, expect, it } from "vitest";
import { IMAGE_LIMITS } from "../config/layoutProfiles";
import {
  clipboardContainsFiles,
  detectImageMimeType,
  ImageValidationError,
  normalizeImageMimeType,
  validateImageBytes,
} from "./imageValidation";

const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00]);
const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0xff, 0xd9]);

describe("image validation", () => {
  it("detects PNG and JPEG signatures", () => {
    expect(detectImageMimeType(png)).toBe("image/png");
    expect(detectImageMimeType(jpeg)).toBe("image/jpeg");
  });

  it("normalizes safe JPEG aliases and MIME parameters", () => {
    expect(normalizeImageMimeType(" image/JPG ")).toBe("image/jpeg");
    expect(normalizeImageMimeType("image/pjpeg; charset=binary")).toBe("image/jpeg");
    expect(normalizeImageMimeType("image/webp")).toBeNull();
  });

  it("accepts JPEG signature when valid containers have trailing bytes", () => {
    expect(detectImageMimeType(new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00]))).toBe("image/jpeg");
  });

  it("detects clipboard file items without treating text as an image", () => {
    expect(clipboardContainsFiles({ items: [{ kind: "file" }], files: [] } as unknown as DataTransfer)).toBe(true);
    expect(clipboardContainsFiles({ items: [{ kind: "string" }], files: [] } as unknown as DataTransfer)).toBe(false);
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

});
