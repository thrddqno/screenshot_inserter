import { useCallback, useEffect, useState } from "react";
import { ALL_MODULES } from "./config/evidenceModules";
import { IMAGE_LIMITS } from "./config/layoutProfiles";
import {
  insertEvidence,
  isOfficeHostAvailable,
  readCurrentScenario,
  removeLastInsertedEvidence,
} from "./services/excelService";
import { getImageFromClipboard, ImageValidationError } from "./services/imageValidation";
import type {
  EvidenceStatus,
  ModuleId,
  PendingImage,
  ScenarioContext,
} from "./types/evidence";

type Message = { kind: "success" | "error" | "info"; text: string } | null;

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : "Unexpected error occurred.";

export default function App() {
  const [scenario, setScenario] = useState<ScenarioContext | null>(null);
  const [evidenceType, setEvidenceType] = useState<ModuleId>("GENERIC_SCREENSHOT");
  const [caption, setCaption] = useState("");
  const [pendingImage, setPendingImage] = useState<PendingImage | null>(null);
  const [status, setStatus] = useState<EvidenceStatus>("UNCHANGED");
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<Message>(null);

  const refreshScenario = useCallback(async () => {
    setBusy(true);
    setMessage({ kind: "info", text: "Reading workbook…" });
    try {
      const nextScenario = await readCurrentScenario();
      setScenario(nextScenario);
      setEvidenceType("GENERIC_SCREENSHOT");
      setCaption("");
      setMessage({ kind: "success", text: `Loaded ${nextScenario.scenarioId}.` });
    } catch (error) {
      setScenario(null);
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setBusy(false);
    }
  }, []);

  useEffect(() => {
    if (isOfficeHostAvailable()) void refreshScenario();
    else setMessage({ kind: "error", text: "Open this page inside an Excel task pane." });
  }, [refreshScenario]);

  const acceptClipboard = useCallback(async (clipboardData: DataTransfer | null) => {
    setBusy(true);
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
      setBusy(false);
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
    setBusy(true);
    try {
      const name = await insertEvidence({
        scenario,
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
      setBusy(false);
    }
  };

  const handleRemove = async () => {
    if (!scenario) return;
    setBusy(true);
    try {
      const name = await removeLastInsertedEvidence(scenario);
      setMessage({ kind: "success", text: `Removed ${name}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      setBusy(false);
    }
  };

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <p className="eyebrow">Excel evidence organizer</p>
          <h1>Test Evidence Inserter</h1>
        </div>
        <button className="button button-secondary" type="button" onClick={() => void refreshScenario()} disabled={busy}>
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
            {ALL_MODULES.map((module) => <option value={module.id} key={module.id}>{module.displayName}</option>)}
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
