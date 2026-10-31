import { useEffect, useMemo, useState } from "react";
import { EVIDENCE_MODULES } from "../config/evidenceModules";
import {
  convertExistingEvidence,
  getWorkbookPicturePreview,
  inferExistingHeaderModule,
  readSelectedExistingHeader,
  scanExistingEvidence,
} from "../services/excelService";
import type {
  ExistingEvidenceScan,
  ExistingEvidenceConversionResult,
  ExistingHeaderReference,
  ModuleId,
  PendingImage,
  ScenarioContext,
  TestTypeId,
} from "../types/evidence";

type ConversionMode = "pictures" | "header";

interface ExistingEvidenceConverterProps {
  scenario: ScenarioContext;
  testTypeId: TestTypeId;
  testTypeName: string;
  moduleIds: readonly ModuleId[];
  onAdoptHeader: (header: ExistingHeaderReference, moduleId: ModuleId) => void;
  onConverted: (result: ExistingEvidenceConversionResult) => void;
  onBack: () => void;
}

const message = (error: unknown): string => error instanceof Error ? error.message : String(error);

export default function ExistingEvidenceConverter({
  scenario,
  testTypeId,
  testTypeName,
  moduleIds,
  onAdoptHeader,
  onConverted,
  onBack,
}: ExistingEvidenceConverterProps) {
  const [mode, setMode] = useState<ConversionMode>("pictures");
  const [moduleId, setModuleId] = useState<ModuleId>(moduleIds[0]);
  const [scan, setScan] = useState<ExistingEvidenceScan | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [previews, setPreviews] = useState<Map<string, PendingImage>>(new Map());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  const [progress, setProgress] = useState<string>();

  useEffect(() => () => {
    previews.forEach((preview) => URL.revokeObjectURL(preview.previewUrl));
  }, [previews]);

  const selectedPictures = useMemo(
    () => scan?.pictures.filter((picture) => selectedIds.has(picture.shapeId)) ?? [],
    [scan, selectedIds],
  );

  const readSelection = async () => {
    setBusy(true);
    setProgress("Scanning worksheet…");
    setError(undefined);
    try {
      const header = await readSelectedExistingHeader(scenario);
      const inferred = inferExistingHeaderModule(header.text, moduleIds);
      setModuleId(inferred);
      if (mode === "header") {
        setScan({ header, pictures: [] });
        setSelectedIds(new Set());
      } else {
        const result = await scanExistingEvidence(header, inferred);
        setScan(result);
        setSelectedIds(new Set(result.pictures.map((picture) => picture.shapeId)));
        const loaded = await Promise.all(result.pictures.map(async (picture) => {
          try {
            return [picture.shapeId, await getWorkbookPicturePreview(picture)] as const;
          } catch {
            return undefined;
          }
        }));
        setPreviews(new Map(loaded.filter((item): item is NonNullable<typeof item> => Boolean(item))));
      }
    } catch (cause) {
      setScan(null);
      setError(message(cause));
    } finally {
      setBusy(false);
      setProgress(undefined);
    }
  };

  const rescan = async (nextModuleId: ModuleId) => {
    setModuleId(nextModuleId);
    if (!scan || mode === "header") return;
    setBusy(true);
    setProgress("Scanning worksheet…");
    setError(undefined);
    try {
      const result = await scanExistingEvidence(scan.header, nextModuleId);
      setScan(result);
      setSelectedIds(new Set(result.pictures.map((picture) => picture.shapeId)));
      const loaded = await Promise.all(result.pictures.map(async (picture) => {
        try {
          return [picture.shapeId, await getWorkbookPicturePreview(picture)] as const;
        } catch {
          return undefined;
        }
      }));
      setPreviews(new Map(loaded.filter((item): item is NonNullable<typeof item> => Boolean(item))));
    } catch (cause) {
      setError(message(cause));
    } finally {
      setBusy(false);
      setProgress(undefined);
    }
  };

  const convert = async () => {
    if (!scan) return;
    setBusy(true);
    setProgress("Validating and converting in place…");
    setError(undefined);
    try {
      const result = await convertExistingEvidence({
        scenario,
        testTypeId,
        moduleId,
        header: scan.header,
        pictures: selectedPictures,
      });
      onConverted(result);
    } catch (cause) {
      setError(message(cause));
    } finally {
      setBusy(false);
      setProgress(undefined);
    }
  };

  return (
    <main className="app-shell arranger-shell">
      <header className="arranger-header">
        <button className="button button-secondary back-button" type="button" disabled={busy} onClick={onBack}>← Back</button>
        <p className="eyebrow">{scenario.scenarioId} · {testTypeName}</p>
        <h1 className="content-title">Convert Existing Evidence</h1>
        <p className="arranger-instructions">Select a worksheet header in A:D or F:I, then review what the add-in found.</p>
      </header>

      <section className="panel converter-controls">
        <label className="field">
          <span>Conversion</span>
          <select value={mode} disabled={busy} onChange={(event) => {
            setMode(event.target.value as ConversionMode);
            setScan(null);
            setError(undefined);
          }}>
            <option value="pictures">Header and pictures</option>
            <option value="header">Header for next insertion</option>
          </select>
        </label>
        <label className="field">
          <span>Evidence Type</span>
          <select value={moduleId} disabled={busy} onChange={(event) => void rescan(event.target.value as ModuleId)}>
            {moduleIds.map((id) => <option key={id} value={id}>{EVIDENCE_MODULES[id].displayName}</option>)}
          </select>
        </label>
        <button className="button button-secondary" type="button" disabled={busy} onClick={() => void readSelection()}>
          {scan ? "Read selected header again" : "Read selected header"}
        </button>
      </section>

      {scan ? (
        <section className="panel converter-review">
          <h2>{scan.header.text}</h2>
          <p className="converter-address">{scan.header.address}</p>
          {mode === "pictures" ? (
            scan.pictures.length > 0 ? (
              <ul className="converter-picture-list">
                {scan.pictures.map((picture) => (
                  <li key={picture.shapeId}>
                    <label>
                      <input
                        type="checkbox"
                        checked={selectedIds.has(picture.shapeId)}
                        disabled={busy}
                        onChange={(event) => setSelectedIds((current) => {
                          const next = new Set(current);
                          if (event.target.checked) next.add(picture.shapeId);
                          else next.delete(picture.shapeId);
                          return next;
                        })}
                      />
                      {previews.get(picture.shapeId) ? <img src={previews.get(picture.shapeId)!.previewUrl} alt="" /> : null}
                      <span><strong>{picture.name}</strong><small>{Math.round(picture.width)} × {Math.round(picture.height)}</small></span>
                    </label>
                  </li>
                ))}
              </ul>
            ) : <p className="empty-arranger">No pictures were found below this header before the next header.</p>
          ) : <p className="converter-note">The next screenshot you insert will use this worksheet header and convert it to managed formatting.</p>}
        </section>
      ) : null}

      {scan && mode === "pictures" && selectedPictures.length > 0 ? (
        <p className="converter-note" role="status">
          {selectedPictures.length} selected picture{selectedPictures.length === 1 ? "" : "s"} and header will stay in their current worksheet positions.
        </p>
      ) : null}

      {error ? <div className="message message-error" role="alert">{error}</div> : null}
      {progress ? <div className="message" role="status">{progress}</div> : null}
      <div className="arranger-actions">
        {mode === "header" ? (
          <button className="button button-primary" type="button" disabled={busy || !scan} onClick={() => scan && onAdoptHeader(scan.header, moduleId)}>
            Use for next insertion
          </button>
        ) : (
          <button className="button button-primary" type="button" disabled={busy || selectedPictures.length === 0} onClick={() => void convert()}>
            Convert {selectedPictures.length || ""} in place
          </button>
        )}
      </div>
    </main>
  );
}
