import { useCallback, useEffect, useRef, useState } from "react";
import { EVIDENCE_MODULES } from "./config/evidenceModules";
import { IMAGE_LIMITS } from "./config/layoutProfiles";
import { AVAILABLE_TEST_TYPES, TEST_TYPE_PROFILES } from "./config/testTypeProfiles";
import {
  insertEvidence,
  isOfficeHostAvailable,
  readCurrentScenario,
  removeLastInsertedEvidence,
  subscribeWorksheetActivated,
  updateScenarioStatus,
} from "./services/excelService";
import { getImageFromClipboard, ImageValidationError } from "./services/imageValidation";
import {
  DEFAULT_WORKBOOK_PREFERENCES,
  readWorkbookPreferences,
  type WorkbookPreferences,
  writeWorkbookPreferences,
} from "./services/workbookPreferences";
import {
  checkForAppUpdate,
  reloadToVersion,
  UPDATE_CHECK_INTERVAL_MS,
} from "./services/updateService";
import type {
  EvidenceStatus,
  GenericSide,
  ModuleId,
  PendingImage,
  ScenarioContext,
  TestTypeId,
} from "./types/evidence";

type Message = { kind: "success" | "error" | "info"; text: string } | null;
const HEADER_PERSIST_DELAY_MS = 400;

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : "Unexpected error occurred.";

export default function App() {
  const [scenario, setScenario] = useState<ScenarioContext | null>(null);
  const [testTypeId, setTestTypeId] = useState<TestTypeId>("GENERAL");
  const [evidenceType, setEvidenceType] = useState<ModuleId>("GENERIC_SCREENSHOT");
  const [caption, setCaption] = useState("");
  const [genericSide, setGenericSide] = useState<GenericSide>("left");
  const [pendingImage, setPendingImage] = useState<PendingImage | null>(null);
  const [status, setStatus] = useState<EvidenceStatus>("UNCHANGED");
  const [refreshing, setRefreshing] = useState(false);
  const [operationBusy, setOperationBusy] = useState(false);
  const [preferencesLoaded, setPreferencesLoaded] = useState(false);
  const [preferenceWarning, setPreferenceWarning] = useState<string | null>(null);
  const [availableVersion, setAvailableVersion] = useState<string | null>(null);
  const [message, setMessage] = useState<Message>({ kind: "info", text: "Loading workbook settings…" });
  const mountedRef = useRef(false);
  const refreshRunningRef = useRef(false);
  const refreshQueuedRef = useRef(false);
  const refreshGenerationRef = useRef(0);
  const preferencesRef = useRef<WorkbookPreferences>(DEFAULT_WORKBOOK_PREFERENCES);
  const requestedPreferencesRef = useRef<WorkbookPreferences>(DEFAULT_WORKBOOK_PREFERENCES);
  const preferenceWriteRunningRef = useRef(false);
  const preferenceWriteTimerRef = useRef<number | undefined>(undefined);
  const updateCheckRunningRef = useRef(false);

  const selectedProfile = TEST_TYPE_PROFILES[testTypeId];
  const availableModules = selectedProfile.moduleIds
    .map((moduleId) => EVIDENCE_MODULES[moduleId])
    .sort((a, b) => a.verticalOrder - b.verticalOrder || a.displayName.localeCompare(b.displayName));

  const refreshScenario = useCallback(async (invalidateImmediately = false) => {
    refreshGenerationRef.current += 1;
    refreshQueuedRef.current = true;
    if (invalidateImmediately && mountedRef.current) {
      setScenario(null);
      setMessage({ kind: "info", text: "Loading test scenario…" });
    }
    if (refreshRunningRef.current) return;

    refreshRunningRef.current = true;
    if (mountedRef.current) setRefreshing(true);
    try {
      while (refreshQueuedRef.current) {
        refreshQueuedRef.current = false;
        const generation = refreshGenerationRef.current;
        if (mountedRef.current) setMessage({ kind: "info", text: "Loading test scenario…" });
        try {
          const result = await readCurrentScenario();
          if (!mountedRef.current || generation !== refreshGenerationRef.current) continue;
          if (result.kind === "loaded") {
            setScenario(result.scenario);
            setStatus(
              result.scenario.status && result.scenario.status !== "UNCHANGED"
                ? result.scenario.status
                : preferencesRef.current.status,
            );
            setMessage({ kind: "success", text: `Loaded ${result.scenario.scenarioId}.` });
          } else {
            setScenario(null);
            setStatus(preferencesRef.current.status);
            setMessage({ kind: "info", text: "No test scenario loaded." });
          }
        } catch (error) {
          if (!mountedRef.current || generation !== refreshGenerationRef.current) continue;
          setScenario(null);
          setStatus(preferencesRef.current.status);
          setMessage({ kind: "error", text: errorMessage(error) });
        }
      }
    } finally {
      refreshRunningRef.current = false;
      if (mountedRef.current) setRefreshing(false);
    }
  }, []);

  const flushPreferences = useCallback(async () => {
    if (preferenceWriteRunningRef.current) return;
    preferenceWriteRunningRef.current = true;
    let lastAttempt: WorkbookPreferences | undefined;
    try {
      while (mountedRef.current) {
        const targetPreferences = requestedPreferencesRef.current;
        lastAttempt = targetPreferences;
        try {
          await writeWorkbookPreferences(targetPreferences);
          if (mountedRef.current) setPreferenceWarning(null);
        } catch {
          if (mountedRef.current) {
            setPreferenceWarning("Form choices changed for this session but could not be saved to the workbook.");
          }
          break;
        }
        if (targetPreferences === requestedPreferencesRef.current) break;
      }
    } finally {
      preferenceWriteRunningRef.current = false;
      if (mountedRef.current && lastAttempt !== requestedPreferencesRef.current) {
        window.setTimeout(() => void flushPreferences(), 0);
      }
    }
  }, []);

  const queuePreferences = useCallback((preferences: WorkbookPreferences, debounce = false) => {
    preferencesRef.current = preferences;
    requestedPreferencesRef.current = preferences;
    if (preferenceWriteTimerRef.current !== undefined) {
      window.clearTimeout(preferenceWriteTimerRef.current);
      preferenceWriteTimerRef.current = undefined;
    }
    if (debounce) {
      preferenceWriteTimerRef.current = window.setTimeout(() => {
        preferenceWriteTimerRef.current = undefined;
        void flushPreferences();
      }, HEADER_PERSIST_DELAY_MS);
      return;
    }
    void flushPreferences();
  }, [flushPreferences]);

  const updateProfilePreferences = useCallback((
    changes: Partial<WorkbookPreferences["profiles"][TestTypeId]>,
    debounce = false,
  ) => {
    const current = preferencesRef.current;
    queuePreferences({
      ...current,
      profiles: {
        ...current.profiles,
        [testTypeId]: { ...current.profiles[testTypeId], ...changes },
      },
    }, debounce);
  }, [queuePreferences, testTypeId]);

  useEffect(() => {
    mountedRef.current = true;
    let disposed = false;
    let unsubscribe: (() => Promise<void>) | undefined;

    if (!isOfficeHostAvailable()) {
      setPreferencesLoaded(true);
      setMessage({ kind: "error", text: "Open this page inside an Excel task pane." });
    } else {
      void (async () => {
        let restoredTestTypeId: TestTypeId = "GENERAL";
        try {
          const result = await readWorkbookPreferences();
          restoredTestTypeId = result.preferences.testTypeId;
          preferencesRef.current = result.preferences;
          requestedPreferencesRef.current = result.preferences;
          if (!disposed) setPreferenceWarning(result.warning ?? null);
        } catch {
          if (!disposed) {
            setPreferenceWarning("Workbook form choices could not be loaded. Defaults were selected.");
          }
        }
        if (disposed) return;
        const restoredForm = preferencesRef.current.profiles[restoredTestTypeId];
        setTestTypeId(restoredTestTypeId);
        setEvidenceType(restoredForm.evidenceType);
        setGenericSide(restoredForm.genericSide);
        setCaption(restoredForm.header);
        setStatus(preferencesRef.current.status);
        setPreferencesLoaded(true);

        try {
          const removeHandler = await subscribeWorksheetActivated(async () => {
            if (!disposed) await refreshScenario(true);
          });
          if (disposed) {
            void removeHandler().catch(() => undefined);
            return;
          }
          unsubscribe = removeHandler;
        } catch {
          if (!disposed && mountedRef.current) {
            setMessage({ kind: "info", text: "Automatic sheet refresh unavailable. Use Refresh test scenario." });
          }
        }
        if (!disposed) await refreshScenario();
      })();
    }

    return () => {
      disposed = true;
      mountedRef.current = false;
      if (preferenceWriteTimerRef.current !== undefined) {
        window.clearTimeout(preferenceWriteTimerRef.current);
      }
      if (unsubscribe) void unsubscribe().catch(() => undefined);
    };
  }, [refreshScenario]);

  useEffect(() => {
    const controller = new AbortController();
    const check = async () => {
      if (updateCheckRunningRef.current) return;
      updateCheckRunningRef.current = true;
      try {
        const nextVersion = await checkForAppUpdate(controller.signal);
        if (!controller.signal.aborted) setAvailableVersion(nextVersion);
      } catch {
        // Update checks are best-effort and must not interrupt workbook work.
      } finally {
        updateCheckRunningRef.current = false;
      }
    };
    const onFocus = () => void check();
    void check();
    window.addEventListener("focus", onFocus);
    const interval = window.setInterval(check, UPDATE_CHECK_INTERVAL_MS);
    return () => {
      controller.abort();
      window.clearInterval(interval);
      window.removeEventListener("focus", onFocus);
    };
  }, []);

  const acceptClipboard = useCallback(async (clipboardData: DataTransfer | null) => {
    setOperationBusy(true);
    setMessage({ kind: "info", text: "Reading pasted screenshot…" });
    try {
      const image = await getImageFromClipboard(clipboardData);
      setPendingImage((current) => {
        if (current) URL.revokeObjectURL(current.previewUrl);
        return image;
      });
      setMessage({
        kind: "success",
        text: `Screenshot ready: ${image.width}×${image.height}, ${(image.size / 1024).toFixed(1)} KB.`,
      });
    } catch (error) {
      const text = error instanceof ImageValidationError ? error.message : errorMessage(error);
      setMessage({ kind: "error", text });
    } finally {
      setOperationBusy(false);
    }
  }, []);

  useEffect(() => {
    const onPaste = (event: ClipboardEvent) => {
      const containsSupportedImage = [...(event.clipboardData?.items ?? [])].some(
        (item) => item.kind === "file" && ["image/png", "image/jpeg"].includes(item.type),
      );
      if (!containsSupportedImage) return;
      event.preventDefault();
      void acceptClipboard(event.clipboardData);
    };
    document.addEventListener("paste", onPaste);
    return () => document.removeEventListener("paste", onPaste);
  }, [acceptClipboard]);

  useEffect(
    () => () => {
      if (pendingImage) URL.revokeObjectURL(pendingImage.previewUrl);
    },
    [pendingImage],
  );

  const clearPending = () => {
    setPendingImage((current) => {
      if (current) URL.revokeObjectURL(current.previewUrl);
      return null;
    });
    setMessage({ kind: "info", text: "Pending screenshot cleared." });
  };

  const handleInsert = async () => {
    if (!scenario || !pendingImage) return;
    setOperationBusy(true);
    setMessage({ kind: "info", text: "Inserting screenshot…" });
    try {
      const name = await insertEvidence({
        scenario,
        testTypeId,
        moduleId: evidenceType,
        image: pendingImage,
        caption: evidenceType === "GENERIC_SCREENSHOT" ? caption : "",
        side: evidenceType === "GENERIC_SCREENSHOT" ? genericSide : undefined,
      });
      URL.revokeObjectURL(pendingImage.previewUrl);
      setPendingImage(null);
      setMessage({ kind: "success", text: `Inserted ${name}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setOperationBusy(false);
    }
  };

  const handleStatusChange = async (nextStatus: EvidenceStatus) => {
    const previousStatus = status;
    setStatus(nextStatus);
    if (nextStatus === "UNCHANGED") {
      queuePreferences({ ...preferencesRef.current, status: nextStatus });
      setMessage({ kind: "info", text: "Worksheet Status left unchanged." });
      return;
    }
    if (!scenario) return;

    setOperationBusy(true);
    setMessage({ kind: "info", text: "Updating worksheet Status…" });
    try {
      await updateScenarioStatus(scenario, testTypeId, nextStatus);
      setScenario((current) => current ? { ...current, status: nextStatus } : current);
      queuePreferences({ ...preferencesRef.current, status: nextStatus });
      setMessage({ kind: "success", text: `Worksheet Status updated to ${nextStatus}.` });
    } catch (error) {
      setStatus(previousStatus);
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setOperationBusy(false);
    }
  };

  const handleRemove = async () => {
    if (!scenario) return;
    setOperationBusy(true);
    setMessage({ kind: "info", text: "Removing screenshot…" });
    try {
      const name = await removeLastInsertedEvidence(scenario, testTypeId);
      setMessage({ kind: "success", text: `Removed ${name}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setOperationBusy(false);
    }
  };

  const busy = !preferencesLoaded || refreshing || operationBusy;

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Excel screenshot organizer</p>
          <h1>Test Evidence Inserter</h1>
          <label className="test-type-field">
            <span>Test Type</span>
            <select
              aria-label="Test Type"
              value={testTypeId}
              disabled={busy}
              onChange={(event) => {
                const nextTestTypeId = event.target.value as TestTypeId;
                const nextPreferences: WorkbookPreferences = {
                  ...preferencesRef.current,
                  testTypeId: nextTestTypeId,
                };
                const nextForm = nextPreferences.profiles[nextTestTypeId];
                setTestTypeId(nextTestTypeId);
                setEvidenceType(nextForm.evidenceType);
                setCaption(nextForm.header);
                setGenericSide(nextForm.genericSide);
                queuePreferences(nextPreferences);
              }}
            >
              {AVAILABLE_TEST_TYPES.map((profile) => (
                <option value={profile.id} key={profile.id}>{profile.displayName}</option>
              ))}
            </select>
          </label>
        </div>
        <button
          className="button button-secondary refresh-button"
          type="button"
          aria-label="Refresh test scenario"
          title="Refresh test scenario"
          onClick={() => void refreshScenario(true)}
          disabled={busy}
        >
          <svg className={`refresh-icon ${refreshing ? "is-spinning" : ""}`} viewBox="0 0 24 24" aria-hidden="true">
            <path d="M20 6v5h-5M4 18v-5h5M6.1 9A7 7 0 0 1 18 6l2 5M17.9 15A7 7 0 0 1 6 18l-2-5" />
          </svg>
        </button>
      </header>

      {availableVersion ? (
        <aside className="update-notice" role="status" aria-live="polite">
          <span>
            {pendingImage
              ? "New version available. Insert or clear pending screenshot before updating."
              : "New version available."}
          </span>
          <button
            className="button button-primary"
            type="button"
            disabled={busy || Boolean(pendingImage)}
            onClick={() => reloadToVersion(availableVersion)}
          >
            Update now
          </button>
        </aside>
      ) : null}

      <section className="panel scenario-panel" aria-labelledby="scenario-heading">
        <h2 id="scenario-heading">Test Scenario</h2>
        <dl className="scenario-id">
          <dt>Scenario ID</dt>
          <dd>{scenario?.scenarioId || "Not loaded"}</dd>
        </dl>
        <details open>
          <summary>Condition</summary>
          <p className="workbook-preview">{scenario?.condition || "—"}</p>
        </details>
        <details>
          <summary>Expected Result</summary>
          <p className="workbook-preview">{scenario?.expectedResult || "—"}</p>
        </details>
      </section>

      <section className="panel" aria-labelledby="status-heading">
        <h2 id="status-heading">Worksheet Status</h2>
        <label className="field status-field">
          <span>Status</span>
          <select
            disabled={busy || !scenario?.statusCellAddress}
            value={status}
            onChange={(event) => void handleStatusChange(event.target.value as EvidenceStatus)}
          >
            <option value="UNCHANGED">unchanged</option>
            <option value="PASS/FAIL">PASS/FAIL</option>
            <option value="PASS">PASS</option>
            <option value="FAIL">FAIL</option>
          </select>
        </label>
      </section>

      <section className="panel" aria-labelledby="capture-heading">
        <h2 id="capture-heading">Screenshot to Insert</h2>
        <label className="field">
          <span>Evidence Type</span>
          <select
            value={evidenceType}
            disabled={busy}
            onChange={(event) => {
              const nextType = event.target.value as ModuleId;
              setEvidenceType(nextType);
              updateProfilePreferences({ evidenceType: nextType });
            }}
          >
            {availableModules.map((module) => <option value={module.id} key={module.id}>{module.displayName}</option>)}
          </select>
        </label>
        {evidenceType === "GENERIC_SCREENSHOT" ? (
          <>
            <label className="field">
              <span>Side</span>
              <select
                value={genericSide}
                disabled={busy}
                onChange={(event) => {
                  const nextSide = event.target.value as GenericSide;
                  setGenericSide(nextSide);
                  updateProfilePreferences({ genericSide: nextSide });
                }}
              >
                <option value="left">Left (A:D)</option>
                <option value="right">Right (F:I)</option>
              </select>
            </label>
            <div className="field">
              <span id="header-label">Header</span>
              <div className="header-input-row">
                <input
                  type="text"
                  aria-labelledby="header-label"
                  value={caption}
                  maxLength={500}
                  disabled={busy}
                  onChange={(event) => {
                    const nextHeader = event.target.value;
                    setCaption(nextHeader);
                    updateProfilePreferences({ header: nextHeader }, true);
                  }}
                  onBlur={() => queuePreferences(preferencesRef.current)}
                  placeholder="Dynamic evidence header (optional)"
                />
                <button
                  className="button button-secondary clear-header-button"
                  type="button"
                  aria-label="Clear Header"
                  disabled={busy || !caption}
                  onClick={() => {
                    setCaption("");
                    updateProfilePreferences({ header: "" });
                  }}
                >
                  Clear
                </button>
              </div>
            </div>
          </>
        ) : null}
        <div className={`paste-zone ${pendingImage ? "has-image" : ""}`} tabIndex={0} aria-label="Screenshot paste area">
          {pendingImage ? (
            <>
              <img src={pendingImage.previewUrl} alt="Pending screenshot preview" />
              <p>{pendingImage.mimeType.replace("image/", "").toUpperCase()} · {(pendingImage.size / 1024).toFixed(1)} KB</p>
            </>
          ) : (
            <>
              <strong>Press Ctrl+V</strong>
              <span>Paste PNG or JPEG screenshot, maximum {IMAGE_LIMITS.maximumBytes / 1024 / 1024} MB.</span>
            </>
          )}
        </div>
      </section>

      <div className="actions">
        <button
          className="button button-primary"
          type="button"
          disabled={busy || !scenario || !pendingImage}
          onClick={() => void handleInsert()}
        >
          Insert
        </button>
        <button className="button button-danger" type="button" disabled={busy || !scenario} onClick={() => void handleRemove()}>
          Remove Last Inserted
        </button>
        <button className="button button-secondary" type="button" disabled={busy || !pendingImage} onClick={clearPending}>
          Clear Pending
        </button>
      </div>

      <div className={`message message-${message?.kind ?? "info"}`} role={message?.kind === "error" ? "alert" : "status"}>
        {message?.text || "Ready."}
      </div>
      {preferenceWarning ? <div className="message message-warning" role="status">{preferenceWarning}</div> : null}
    </main>
  );
}
