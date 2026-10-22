import { IMAGE_LIMITS } from "../config/layoutProfiles";
import type { PendingImage } from "../types/evidence";

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] as const;

export class ImageValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ImageValidationError";
  }
}

export const detectImageMimeType = (bytes: Uint8Array): "image/png" | "image/jpeg" | null => {
  const isPng =
    bytes.length >= PNG_SIGNATURE.length &&
    PNG_SIGNATURE.every((value, index) => bytes[index] === value);
  if (isPng) return "image/png";

  const isJpeg =
    bytes.length >= 4 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff &&
    bytes.at(-2) === 0xff &&
    bytes.at(-1) === 0xd9;
  return isJpeg ? "image/jpeg" : null;
};

export const validateImageBytes = (
  declaredMimeType: string,
  size: number,
  bytes: Uint8Array,
): "image/png" | "image/jpeg" => {
  if (!(IMAGE_LIMITS.acceptedMimeTypes as readonly string[]).includes(declaredMimeType)) {
    throw new ImageValidationError("Only PNG and JPEG screenshots are accepted.");
  }
  if (size <= 0) throw new ImageValidationError("Screenshot is empty.");
  if (size > IMAGE_LIMITS.maximumBytes) {
    throw new ImageValidationError("Screenshot exceeds 10 MB limit.");
  }

  const detected = detectImageMimeType(bytes);
  if (!detected || detected !== declaredMimeType) {
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

const decodeDimensions = (blob: Blob): Promise<{ width: number; height: number }> =>
  new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      const dimensions = { width: image.naturalWidth, height: image.naturalHeight };
      URL.revokeObjectURL(url);
      if (dimensions.width <= 0 || dimensions.height <= 0) {
        reject(new ImageValidationError("Screenshot has invalid dimensions."));
      } else {
        resolve(dimensions);
      }
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new ImageValidationError("Screenshot could not be decoded."));
    };
    image.src = url;
  });

export const createPendingImage = async (blob: Blob): Promise<PendingImage> => {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const mimeType = validateImageBytes(blob.type, blob.size, bytes);
  const [{ width, height }, dataUrl] = await Promise.all([decodeDimensions(blob), readAsDataUrl(blob)]);
  const marker = "base64,";
  const markerIndex = dataUrl.indexOf(marker);
  if (markerIndex < 0) throw new ImageValidationError("Screenshot encoding is invalid.");

  return {
    base64: dataUrl.slice(markerIndex + marker.length),
    mimeType,
    size: blob.size,
    width,
    height,
    previewUrl: URL.createObjectURL(blob),
  };
};

export const getImageFromClipboard = async (clipboardData: DataTransfer | null): Promise<PendingImage> => {
  if (!clipboardData) throw new ImageValidationError("Clipboard data is unavailable.");
  const imageItems = [...clipboardData.items].filter((item) => item.kind === "file");
  const item = imageItems.find((candidate) =>
    (IMAGE_LIMITS.acceptedMimeTypes as readonly string[]).includes(candidate.type),
  );
  if (!item) {
    throw new ImageValidationError(
      "Clipboard contains no PNG or JPEG image. Select the Excel picture itself, copy it, then paste here.",
    );
  }
  const file = item.getAsFile();
  if (!file) throw new ImageValidationError("Clipboard screenshot could not be read.");
  return createPendingImage(file);
};

export const getImageFromSystemClipboard = async (): Promise<PendingImage> => {
  if (!navigator.clipboard?.read) {
    throw new ImageValidationError("Clipboard reading is unavailable. Focus the task pane and press Ctrl+V instead.");
  }

  let clipboardItems: ClipboardItems;
  try {
    clipboardItems = await navigator.clipboard.read();
  } catch {
    throw new ImageValidationError("Clipboard access was denied. Focus the task pane and press Ctrl+V instead.");
  }

  for (const clipboardItem of clipboardItems) {
    const mimeType = clipboardItem.types.find((type) =>
      (IMAGE_LIMITS.acceptedMimeTypes as readonly string[]).includes(type),
    );
    if (mimeType) return createPendingImage(await clipboardItem.getType(mimeType));
  }

  throw new ImageValidationError(
    "Clipboard contains no PNG or JPEG image. Select the Excel picture itself, copy it, then try again.",
  );
};
