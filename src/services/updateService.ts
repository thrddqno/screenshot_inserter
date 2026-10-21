export const CURRENT_APP_VERSION = __APP_VERSION__;
export const UPDATE_CHECK_INTERVAL_MS = 15 * 60 * 1000;

const VERSION_PATTERN = /^[a-f0-9]{12}$/;

const isVersionDocument = (value: unknown): value is { version: string } => {
  if (!value || typeof value !== "object") return false;
  const version = (value as { version?: unknown }).version;
  return typeof version === "string" && VERSION_PATTERN.test(version);
};

export const checkForAppUpdate = async (signal?: AbortSignal): Promise<string | null> => {
  const requestUrl = new URL(`${import.meta.env.BASE_URL}version.json`, window.location.origin);
  requestUrl.searchParams.set("check", Date.now().toString(36));
  const response = await fetch(requestUrl, {
    method: "GET",
    cache: "no-store",
    credentials: "same-origin",
    headers: { Accept: "application/json" },
    redirect: "error",
    signal,
  });
  if (!response.ok) throw new Error(`Update check failed with HTTP ${response.status}.`);
  const document: unknown = await response.json();
  if (!isVersionDocument(document)) throw new Error("Update response is invalid.");
  return document.version === CURRENT_APP_VERSION ? null : document.version;
};

export const buildUpdateUrl = (version: string, currentHref = window.location.href): string => {
  if (!VERSION_PATTERN.test(version)) throw new Error("Update version is invalid.");
  const url = new URL(currentHref);
  url.searchParams.set("version", version);
  url.hash = "";
  return url.href;
};

export const reloadToVersion = (version: string): void => {
  window.location.replace(buildUpdateUrl(version));
};
