import { IMAGE_LIMITS } from "../config/layoutProfiles";
import type { PendingImage } from "../types/evidence";

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] as const;
const MIME_ALIASES: Readonly<Record<string, "image/png" | "image/jpeg">> = {
  "image/png": "image/png",
  "image/jpeg": "image/jpeg",
  "image/jpg": "image/jpeg",
  "image/pjpeg": "image/jpeg",
};

export class ImageValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ImageValidationError";
  }
}

export const normalizeImageMimeType = (value: string): "image/png" | "image/jpeg" | null => {
  const canonical = MIME_ALIASES[value.split(";", 1)[0].trim().toLowerCase()];
  return canonical && (IMAGE_LIMITS.acceptedMimeTypes as readonly string[]).includes(canonical)
    ? canonical
    : null;
};

export const detectImageMimeType = (bytes: Uint8Array): "image/png" | "image/jpeg" | null => {
  const isPng =
    bytes.length >= PNG_SIGNATURE.length &&
    PNG_SIGNATURE.every((value, index) => bytes[index] === value);
  if (isPng) return "image/png";

  const isJpeg =
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff;
  return isJpeg ? "image/jpeg" : null;
};

export const validateImageBytes = (
  declaredMimeType: string,
  size: number,
  bytes: Uint8Array,
): "image/png" | "image/jpeg" => {
  const normalizedMimeType = normalizeImageMimeType(declaredMimeType);
  if (!normalizedMimeType) {
    throw new ImageValidationError("Only PNG and JPEG screenshots are accepted.");
  }
  if (size <= 0) throw new ImageValidationError("Screenshot is empty.");
  if (size > IMAGE_LIMITS.maximumBytes) {
    throw new ImageValidationError("Screenshot exceeds 10 MB limit.");
  }

  const detected = detectImageMimeType(bytes);
  if (!detected || detected !== normalizedMimeType) {
    throw new ImageValidationError("Screenshot content does not match its PNG/JPEG MIME type.");
  }
  return detected;
};

const readAsDataUrl = (blob: Blob): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new ImageValidationError("Could not read screenshot."));
    reader.onload = () => resolve(String(reader.result));
    reader.readAsDataURL(blob);
  });

const decodeDimensions = (url: string): Promise<{ width: number; height: number }> =>
  new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      const dimensions = { width: image.naturalWidth, height: image.naturalHeight };
      if (dimensions.width <= 0 || dimensions.height <= 0) {
        reject(new ImageValidationError("Screenshot has invalid dimensions."));
      } else {
        resolve(dimensions);
      }
    };
    image.onerror = () => {
      reject(new ImageValidationError("Screenshot could not be decoded."));
    };
    image.src = url;
  });

export const createPendingImage = async (blob: Blob): Promise<PendingImage> => {
  const mimeType = normalizeImageMimeType(blob.type);
  if (!mimeType) throw new ImageValidationError("Only PNG and JPEG screenshots are accepted.");
  if (blob.size <= 0) throw new ImageValidationError("Screenshot is empty.");
  if (blob.size > IMAGE_LIMITS.maximumBytes) {
    throw new ImageValidationError("Screenshot exceeds 10 MB limit.");
  }
  const bytes = new Uint8Array(await blob.arrayBuffer());
  validateImageBytes(mimeType, blob.size, bytes);
  const previewUrl = URL.createObjectURL(blob);
  try {
    const [{ width, height }, dataUrl] = await Promise.all([decodeDimensions(previewUrl), readAsDataUrl(blob)]);
    const marker = "base64,";
    const markerIndex = dataUrl.indexOf(marker);
    if (markerIndex < 0) throw new ImageValidationError("Screenshot encoding is invalid.");

    return {
      base64: dataUrl.slice(markerIndex + marker.length),
      mimeType,
      size: blob.size,
      width,
      height,
      previewUrl,
    };
  } catch (error) {
    URL.revokeObjectURL(previewUrl);
    throw error;
  }
};

const clipboardFiles = (clipboardData: DataTransfer): File[] => {
  const fromItems = [...clipboardData.items]
    .filter((item) => item.kind === "file")
    .map((item) => item.getAsFile())
    .filter((file): file is File => file !== null);
  return fromItems.length > 0 ? fromItems : [...(clipboardData.files ?? [])];
};

export const clipboardContainsFiles = (clipboardData: DataTransfer | null): boolean =>
  Boolean(clipboardData && (
    [...clipboardData.items].some((item) => item.kind === "file") || (clipboardData.files?.length ?? 0) > 0
  ));

export const getImageFromClipboard = async (clipboardData: DataTransfer | null): Promise<PendingImage> => {
  if (!clipboardData) throw new ImageValidationError("Clipboard data is unavailable.");
  const files = clipboardFiles(clipboardData);
  const file = files.find((candidate) => normalizeImageMimeType(candidate.type));
  if (!file) throw new ImageValidationError("Clipboard contains no browser-accessible PNG or JPEG screenshot.");
  return createPendingImage(file);
};
