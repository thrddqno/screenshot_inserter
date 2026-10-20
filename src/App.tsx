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
} from "./services/excelService";
import { getImageFromClipboard, ImageValidationError } from "./services/imageValidation";
import type {
  EvidenceStatus,
  ModuleId,
  PendingImage,
  ScenarioContext,
  TestTypeId,
} from "./types/evidence";

type Message = { kind: "success" | "error" | "info"; text: string } | null;

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : "Unexpected error occurred.";

export default function App() {
  const [scenario, setScenario] = useState<ScenarioContext | null>(null);
  const [testTypeId, setTestTypeId] = useState<TestTypeId>("SOA");
  const [evidenceType, setEvidenceType] = useState<ModuleId>("GENERIC_SCREENSHOT");
  const [caption, setCaption] = useState("");
  const [pendingImage, setPendingImage] = useState<PendingImage | null>(null);
  const [status, setStatus] = useState<EvidenceStatus>("UNCHANGED");
  const [refreshing, setRefreshing] = useState(false);
  const [operationBusy, setOperationBusy] = useState(false);
  const [message, setMessage] = useState<Message>(null);
  const mountedRef = useRef(false);
  const refreshRunningRef = useRef(false);
  const refreshQueuedRef = useRef(false);
  const refreshGenerationRef = useRef(0);
  const testTypeIdRef = useRef<TestTypeId>(testTypeId);

  const selectedProfile = TEST_TYPE_PROFILES[testTypeId];
  const availableModules = selectedProfile.moduleIds
    .map((moduleId) => EVIDENCE_MODULES[moduleId])
    .sort((a, b) => a.verticalOrder - b.verticalOrder || a.displayName.localeCompare(b.displayName));

  const refreshScenario = useCallback(async (invalidateImmediately = false) => {
    refreshGenerationRef.current += 1;
    refreshQueuedRef.current = true;
    if (invalidateImmediately && mountedRef.current) {
      setScenario(null);
      setEvidenceType(TEST_TYPE_PROFILES[testTypeIdRef.current].defaultModuleId);
      setCaption("");
      setMessage({ kind: "info", text: "Reading workbook…" });
    }
    if (refreshRunningRef.current) return;

    refreshRunningRef.current = true;
    if (mountedRef.current) setRefreshing(true);
    try {
      while (refreshQueuedRef.current) {
        refreshQueuedRef.current = false;
        const generation = refreshGenerationRef.current;
        if (mountedRef.current) setMessage({ kind: "info", text: "Reading workbook…" });
        try {
          const nextScenario = await readCurrentScenario();
          if (!mountedRef.current || generation !== refreshGenerationRef.current) continue;
          setScenario(nextScenario);
          setEvidenceType(TEST_TYPE_PROFILES[testTypeIdRef.current].defaultModuleId);
          setCaption("");
          setMessage({ kind: "success", text: `Loaded ${nextScenario.scenarioId}.` });
        } catch (error) {
          if (!mountedRef.current || generation !== refreshGenerationRef.current) continue;
          setScenario(null);
          setMessage({ kind: "error", text: errorMessage(error) });
        }
      }
    } finally {
      refreshRunningRef.current = false;
      if (mountedRef.current) setRefreshing(false);
    }
  }, []);

  useEffect(() => {
    mountedRef.current = true;
    let disposed = false;
    let unsubscribe: (() => Promise<void>) | undefined;

    if (isOfficeHostAvailable()) {
      void refreshScenario();
      void subscribeWorksheetActivated(async () => {
        if (!disposed) await refreshScenario(true);
      })
        .then((removeHandler) => {
          if (disposed) void removeHandler().catch(() => undefined);
          else unsubscribe = removeHandler;
        })
        .catch(() => {
          if (!disposed && mountedRef.current) {
            setMessage({ kind: "info", text: "Automatic sheet refresh unavailable. Use Refresh." });
          }
        });
    } else {
      setMessage({ kind: "error", text: "Open this page inside an Excel task pane." });
    }

    return () => {
      disposed = true;
      mountedRef.current = false;
      if (unsubscribe) void unsubscribe().catch(() => undefined);
    };
  }, [refreshScenario]);

  const acceptClipboard = useCallback(async (clipboardData: DataTransfer | null) => {
    setOperationBusy(true);
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
    try {
      const name = await insertEvidence({
        scenario,
        testTypeId,
        moduleId: evidenceType,
        image: pendingImage,
        caption: evidenceType === "GENERIC_SCREENSHOT" ? caption : "",
        status,
      });
      URL.revokeObjectURL(pendingImage.previewUrl);
      setPendingImage(null);
      setCaption("");
      setMessage({ kind: "success", text: `Inserted ${name}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setOperationBusy(false);
    }
  };

  const handleRemove = async () => {
    if (!scenario) return;
    setOperationBusy(true);
    try {
      const name = await removeLastInsertedEvidence(scenario, testTypeId);
      setMessage({ kind: "success", text: `Removed ${name}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setOperationBusy(false);
    }
  };

  const busy = refreshing || operationBusy;

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Excel evidence organizer</p>
          <h1>Test Evidence Inserter</h1>
          <label className="test-type-field">
            <span>Test Type</span>
            <select
              aria-label="Test Type"
              value={testTypeId}
              onChange={(event) => {
                const nextTestTypeId = event.target.value as TestTypeId;
                testTypeIdRef.current = nextTestTypeId;
                setTestTypeId(nextTestTypeId);
                setEvidenceType(TEST_TYPE_PROFILES[nextTestTypeId].defaultModuleId);
                setCaption("");
              }}
            >
              {AVAILABLE_TEST_TYPES.map((profile) => (
                <option value={profile.id} key={profile.id}>{profile.displayName}</option>
              ))}
            </select>
          </label>
        </div>
        <button className="button button-secondary" type="button" onClick={() => void refreshScenario(true)} disabled={busy}>
          Refresh
        </button>
      </header>

      <section className="panel scenario-panel" aria-labelledby="scenario-heading">
        <h2 id="scenario-heading">Current scenario</h2>
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

      <section className="panel" aria-labelledby="capture-heading">
        <h2 id="capture-heading">Pending evidence</h2>
        <label className="field">
          <span>Evidence Type</span>
          <select
            value={evidenceType}
            onChange={(event) => {
              const nextType = event.target.value as ModuleId;
              setEvidenceType(nextType);
              if (nextType !== "GENERIC_SCREENSHOT") setCaption("");
            }}
          >
            {availableModules.map((module) => <option value={module.id} key={module.id}>{module.displayName}</option>)}
          </select>
        </label>
        {evidenceType === "GENERIC_SCREENSHOT" ? (
          <label className="field">
            <span>Header</span>
            <input
              type="text"
              value={caption}
              maxLength={500}
              onChange={(event) => setCaption(event.target.value)}
              placeholder="Dynamic evidence header (optional)"
            />
          </label>
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
        <label className="field">
          <span>Status</span>
          <select value={status} onChange={(event) => setStatus(event.target.value as EvidenceStatus)}>
            <option value="UNCHANGED">unchanged</option>
            <option value="PASS">PASS</option>
            <option value="FAIL">FAIL</option>
            <option value="BLOCKED">BLOCKED</option>
          </select>
        </label>
      </section>

      <div className="actions">
        <button
          className="button button-primary"
          type="button"
          disabled={busy || !scenario || !pendingImage}
          onClick={() => void handleInsert()}
        >
          Insert Evidence
        </button>
        <button className="button button-danger" type="button" disabled={busy || !scenario} onClick={() => void handleRemove()}>
          Remove Last Inserted Evidence
        </button>
        <button className="button button-secondary" type="button" disabled={!pendingImage} onClick={clearPending}>
          Clear Pending Screenshot
        </button>
      </div>

      <div className={`message message-${message?.kind ?? "info"}`} role={message?.kind === "error" ? "alert" : "status"}>
        {busy ? "Working…" : message?.text || "Ready."}
      </div>
    </main>
  );
}
