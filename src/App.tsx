import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import EvidenceArranger from "./components/EvidenceArranger";
import ExistingEvidenceConverter from "./components/ExistingEvidenceConverter";
import { EVIDENCE_MODULES } from "./config/evidenceModules";
import { IMAGE_LIMITS } from "./config/layoutProfiles";
import { AVAILABLE_TEST_TYPES, TEST_TYPE_PROFILES } from "./config/testTypeProfiles";
import {
  applyEvidenceArrangement,
  getWorkbookPicturePreview,
  insertEvidence,
  isOfficeHostAvailable,
  readCurrentScenario,
  readEvidenceArrangement,
  removeLastInsertedEvidence,
  resetEvidenceArrangement,
  subscribeWorkbookPictureActivated,
  type WorkbookPictureSubscription,
  subscribeWorksheetActivated,
  updateScenarioStatus,
  type EvidenceArrangement,
} from "./services/excelService";
import {
  clipboardContainsFiles,
  getImageFromClipboard,
  ImageValidationError,
} from "./services/imageValidation";
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
  ExistingHeaderReference,
  ExistingEvidenceConversionResult,
  GenericSide,
  ModuleId,
  PendingImage,
  ScenarioContext,
  TestTypeId,
  WorkbookPictureReference,
} from "./types/evidence";

type Message = { kind: "success" | "error" | "info"; text: string } | null;
type AppView = "insert" | "arrange" | "convert";
const HEADER_PERSIST_DELAY_MS = 400;

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : "Unexpected error occurred.";

type AppIconName = "scenario" | "status" | "image" | "paste" | "plus" | "trash" | "close" | "success" | "error" | "info" | "help";

const AppIcon = ({ name, className = "" }: { name: AppIconName; className?: string }) => {
  const paths: Record<AppIconName, ReactNode> = {
    scenario: <><path d="M7 3h7l4 4v14H7z" /><path d="M14 3v5h5M10 12h5M10 16h5" /></>,
    status: <><path d="M5 20v-7h4v7zM10 20V8h4v12zM15 20V4h4v16z" /></>,
    image: <><rect x="3" y="4" width="18" height="16" rx="2" /><circle cx="9" cy="10" r="2" /><path d="m5 18 5-5 3 3 3-3 3 3" /></>,
    paste: <><path d="M9 5h6M10 3h4a2 2 0 0 1 2 2v2H8V5a2 2 0 0 1 2-2Z" /><path d="M7 6H5v15h14V6h-2" /></>,
    plus: <path d="M12 5v14M5 12h14" />,
    trash: <><path d="M4 7h16M9 7V4h6v3M7 7l1 14h8l1-14M10 11v6M14 11v6" /></>,
    close: <path d="m7 7 10 10M17 7 7 17" />,
    success: <><circle cx="12" cy="12" r="9" /><path d="m8 12 3 3 5-6" /></>,
    error: <><circle cx="12" cy="12" r="9" /><path d="m9 9 6 6M15 9l-6 6" /></>,
    info: <><circle cx="12" cy="12" r="9" /><path d="M12 11v6M12 7h.01" /></>,
    help: <><circle cx="12" cy="12" r="9" /><path d="M9.7 9a2.5 2.5 0 1 1 4.1 1.9c-1 .8-1.8 1.2-1.8 2.6M12 17h.01" /></>,
  };
  return <svg className={`app-icon ${className}`} viewBox="0 0 24 24" aria-hidden="true">{paths[name]}</svg>;
};

export default function App() {
  const [view, setView] = useState<AppView>("insert");
  const [arrangement, setArrangement] = useState<EvidenceArrangement | null>(null);
  const [arrangerError, setArrangerError] = useState<string | undefined>();
  const [scenario, setScenario] = useState<ScenarioContext | null>(null);
  const [testTypeId, setTestTypeId] = useState<TestTypeId>("GENERAL");
  const [evidenceType, setEvidenceType] = useState<ModuleId>("GENERIC_SCREENSHOT");
  const [caption, setCaption] = useState("");
  const [genericSide, setGenericSide] = useState<GenericSide>("left");
  const [pendingImage, setPendingImage] = useState<PendingImage | null>(null);
  const [pendingWorkbookPicture, setPendingWorkbookPicture] = useState<WorkbookPictureReference | null>(null);
  const [pendingAdoptedHeader, setPendingAdoptedHeader] = useState<{
    header: ExistingHeaderReference;
    moduleId: ModuleId;
  } | null>(null);
  const [pictureSelectionArmed, setPictureSelectionArmed] = useState(false);
  const [status, setStatus] = useState<EvidenceStatus>("PASS/FAIL");
  const [refreshing, setRefreshing] = useState(false);
  const [operationBusy, setOperationBusy] = useState(false);
  const [pasteBusy, setPasteBusy] = useState(false);
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
  const pictureSelectionArmedRef = useRef(false);
  const pictureSubscriptionRef = useRef<WorkbookPictureSubscription | null>(null);
  const workbookOperationRef = useRef(false);
  const deferredRefreshRef = useRef(false);
  const pasteGenerationRef = useRef(0);
  const pendingImageRef = useRef<PendingImage | null>(null);
  const pendingWorkbookPictureRef = useRef<WorkbookPictureReference | null>(null);
  const viewRef = useRef<AppView>("insert");

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
            setPendingAdoptedHeader((current) =>
              current?.header.worksheetName === result.scenario.worksheetName ? current : null,
            );
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

  const finishWorkbookOperation = useCallback(() => {
    workbookOperationRef.current = false;
    if (mountedRef.current) setOperationBusy(false);
    if (deferredRefreshRef.current) {
      deferredRefreshRef.current = false;
      void refreshScenario(true);
    }
  }, [refreshScenario]);

  const beginWorkbookOperation = useCallback((): boolean => {
    if (workbookOperationRef.current) return false;
    workbookOperationRef.current = true;
    setOperationBusy(true);
    return true;
  }, []);

  const replacePendingImage = useCallback((
    next: PendingImage | null,
    workbookPicture: WorkbookPictureReference | null = null,
  ) => {
    const previous = pendingImageRef.current;
    pendingImageRef.current = next;
    pendingWorkbookPictureRef.current = workbookPicture;
    setPendingImage(next);
    setPendingWorkbookPicture(workbookPicture);
    if (previous && previous !== next) URL.revokeObjectURL(previous.previewUrl);
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

  const acceptImage = useCallback(async (
    readImage: () => Promise<PendingImage>,
    progressMessage: string,
    workbookPicture: WorkbookPictureReference | null = null,
    successMessage?: string,
  ) => {
    const generation = ++pasteGenerationRef.current;
    pictureSelectionArmedRef.current = false;
    setPictureSelectionArmed(false);
    setPasteBusy(true);
    setMessage({ kind: "info", text: progressMessage });
    try {
      const image = await readImage();
      if (!mountedRef.current || generation !== pasteGenerationRef.current) {
        URL.revokeObjectURL(image.previewUrl);
        return;
      }
      replacePendingImage(image, workbookPicture);
      setMessage({
        kind: "success",
        text: successMessage ??
          `Screenshot ready: ${image.width}×${image.height}, ${(image.size / 1024).toFixed(1)} KB.`,
      });
    } catch (error) {
      if (!mountedRef.current || generation !== pasteGenerationRef.current) return;
      const text = error instanceof ImageValidationError ? error.message : errorMessage(error);
      setMessage({ kind: "error", text });
    } finally {
      if (mountedRef.current && generation === pasteGenerationRef.current) setPasteBusy(false);
    }
  }, [replacePendingImage]);

  useEffect(() => {
    mountedRef.current = true;
    let disposed = false;
    let unsubscribe: (() => Promise<void>) | undefined;
    let pictureSubscription: WorkbookPictureSubscription | undefined;

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
          pictureSubscription = await subscribeWorkbookPictureActivated(async (picture) => {
            if (disposed || viewRef.current !== "insert" || !pictureSelectionArmedRef.current) return;
            if (workbookOperationRef.current) {
              setMessage({ kind: "error", text: "Wait for workbook operation to finish before selecting a picture." });
              return;
            }
            await acceptImage(
              () => getWorkbookPicturePreview(picture),
              `Reading pasted Excel picture ${picture.name}…`,
              picture,
              `Pasted Excel picture ready: ${picture.name}.`,
            );
          });
          if (disposed) {
            void pictureSubscription.unsubscribe().catch(() => undefined);
            return;
          }
          pictureSubscriptionRef.current = pictureSubscription;
        } catch {
          // Clipboard screenshots remain available if shape activation events are unsupported.
        }

        try {
          const removeHandler = await subscribeWorksheetActivated(async () => {
            if (disposed) return;
            if (viewRef.current !== "insert") {
              viewRef.current = "insert";
              setView("insert");
              setArrangement(null);
              setArrangerError(undefined);
            }
            if (workbookOperationRef.current) {
              deferredRefreshRef.current = true;
              return;
            }
            await refreshScenario(true);
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
      pictureSelectionArmedRef.current = false;
      pictureSubscriptionRef.current = null;
      pasteGenerationRef.current += 1;
      const pending = pendingImageRef.current;
      pendingImageRef.current = null;
      pendingWorkbookPictureRef.current = null;
      if (pending) URL.revokeObjectURL(pending.previewUrl);
      if (preferenceWriteTimerRef.current !== undefined) {
        window.clearTimeout(preferenceWriteTimerRef.current);
      }
      if (unsubscribe) void unsubscribe().catch(() => undefined);
      if (pictureSubscription) void pictureSubscription.unsubscribe().catch(() => undefined);
    };
  }, [acceptImage, refreshScenario]);

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
    await acceptImage(() => getImageFromClipboard(clipboardData), "Reading pasted picture…");
  }, [acceptImage]);

  useEffect(() => {
    const onPaste = (event: ClipboardEvent) => {
      if (viewRef.current !== "insert") return;
      if (!clipboardContainsFiles(event.clipboardData)) return;
      event.preventDefault();
      if (workbookOperationRef.current) {
        setMessage({ kind: "error", text: "Wait for workbook operation to finish before pasting." });
        return;
      }
      void acceptClipboard(event.clipboardData);
    };
    document.addEventListener("paste", onPaste);
    return () => document.removeEventListener("paste", onPaste);
  }, [acceptClipboard]);

  const clearPending = () => {
    pictureSelectionArmedRef.current = false;
    setPictureSelectionArmed(false);
    pasteGenerationRef.current += 1;
    setPasteBusy(false);
    replacePendingImage(null);
    setMessage({ kind: "info", text: "Pending screenshot cleared." });
  };

  const handleUsePastedPicture = async () => {
    if (pictureSelectionArmedRef.current) {
      pictureSelectionArmedRef.current = false;
      setPictureSelectionArmed(false);
      setMessage({ kind: "info", text: "Pasted picture selection cancelled." });
      return;
    }
    if (!beginWorkbookOperation()) return;
    setMessage({ kind: "info", text: "Finding pasted Excel pictures…" });
    try {
      const subscription = pictureSubscriptionRef.current;
      if (!subscription) throw new Error("Worksheet picture selection is unavailable in this Excel version.");
      await subscription.refresh();
      pictureSelectionArmedRef.current = true;
      setPictureSelectionArmed(true);
      setMessage({ kind: "info", text: "Click the picture you pasted onto the destination worksheet." });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleInsert = async () => {
    const imageIntent = pendingImageRef.current;
    const scenarioIntent = scenario;
    if (!scenarioIntent || !imageIntent || !beginWorkbookOperation()) return;
    const insertIntent = {
      scenario: scenarioIntent,
      testTypeId,
      moduleId: evidenceType,
      image: imageIntent,
      caption: evidenceType === "GENERIC_SCREENSHOT" ? caption : "",
      side: evidenceType === "GENERIC_SCREENSHOT" ? genericSide : undefined,
      sourceShape: pendingWorkbookPictureRef.current ?? undefined,
      adoptedHeader: pendingAdoptedHeader?.moduleId === evidenceType ? pendingAdoptedHeader.header : undefined,
    } as const;
    setMessage({ kind: "info", text: "Inserting screenshot…" });
    try {
      const name = await insertEvidence(insertIntent);
      if (pendingImageRef.current === imageIntent) replacePendingImage(null);
      setPendingAdoptedHeader(null);
      setMessage({ kind: "success", text: `Inserted ${name} into ${scenarioIntent.worksheetName}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleStatusChange = async (nextStatus: Exclude<EvidenceStatus, "UNCHANGED">) => {
    const previousStatus = status;
    if (!beginWorkbookOperation()) return;
    setStatus(nextStatus);
    if (!scenario) {
      finishWorkbookOperation();
      return;
    }

    setMessage({ kind: "info", text: "Updating worksheet Status…" });
    try {
      await updateScenarioStatus(scenario, nextStatus);
      setScenario((current) => current ? { ...current, status: nextStatus } : current);
      queuePreferences({ ...preferencesRef.current, status: nextStatus });
      setMessage({ kind: "success", text: `Worksheet Status updated to ${nextStatus}.` });
    } catch (error) {
      setStatus(previousStatus);
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleRemove = async () => {
    const scenarioIntent = scenario;
    if (!scenarioIntent || !beginWorkbookOperation()) return;
    setMessage({ kind: "info", text: "Removing screenshot…" });
    try {
      const name = await removeLastInsertedEvidence(scenarioIntent, testTypeId);
      setMessage({ kind: "success", text: `Removed ${name}.` });
    } catch (error) {
      setMessage({ kind: "error", text: errorMessage(error) });
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleOpenArranger = async () => {
    const scenarioIntent = scenario;
    if (!scenarioIntent || !beginWorkbookOperation()) return;
    pictureSelectionArmedRef.current = false;
    setPictureSelectionArmed(false);
    pasteGenerationRef.current += 1;
    setPasteBusy(false);
    viewRef.current = "arrange";
    setView("arrange");
    setArrangement(null);
    setArrangerError(undefined);
    try {
      setArrangement(await readEvidenceArrangement(scenarioIntent, testTypeId));
    } catch (error) {
      setArrangerError(errorMessage(error));
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleApplyArrangement = async (order: readonly string[]) => {
    const scenarioIntent = scenario;
    if (!scenarioIntent || !beginWorkbookOperation()) return;
    setArrangerError(undefined);
    try {
      const result = await applyEvidenceArrangement(scenarioIntent, testTypeId, order);
      setArrangement(result);
      setMessage({ kind: "success", text: `Evidence arrangement applied to ${scenarioIntent.worksheetName}.` });
      viewRef.current = "insert";
      setView("insert");
    } catch (error) {
      setArrangerError(errorMessage(error));
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleResetArrangement = async () => {
    const scenarioIntent = scenario;
    if (!scenarioIntent || !beginWorkbookOperation()) return;
    setArrangerError(undefined);
    try {
      const result = await resetEvidenceArrangement(scenarioIntent, testTypeId);
      setArrangement(result);
      setMessage({ kind: "success", text: "Evidence order reset to chronological." });
    } catch (error) {
      setArrangerError(errorMessage(error));
    } finally {
      finishWorkbookOperation();
    }
  };

  const handleCloseArranger = () => {
    viewRef.current = "insert";
    setView("insert");
    setArrangement(null);
    setArrangerError(undefined);
  };

  const handleOpenConverter = () => {
    if (!scenario || busy) return;
    pictureSelectionArmedRef.current = false;
    setPictureSelectionArmed(false);
    viewRef.current = "convert";
    setView("convert");
  };

  const handleAdoptExistingHeader = (header: ExistingHeaderReference, moduleId: ModuleId) => {
    setPendingAdoptedHeader({ header, moduleId });
    setEvidenceType(moduleId);
    if (moduleId === "GENERIC_SCREENSHOT") setCaption(header.text);
    updateProfilePreferences({
      evidenceType: moduleId,
      ...(moduleId === "GENERIC_SCREENSHOT" ? { header: header.text } : {}),
    });
    viewRef.current = "insert";
    setView("insert");
    setMessage({ kind: "info", text: `Header ${header.address} will be used for the next insertion.` });
  };

  const handleConvertedExistingEvidence = (result: ExistingEvidenceConversionResult) => {
    viewRef.current = "insert";
    setView("insert");
    const count = result.converted.length;
    setMessage({ kind: "success", text: `Converted ${count} existing picture${count === 1 ? "" : "s"} in place.` });
  };

  const busy = !preferencesLoaded || refreshing || operationBusy || pasteBusy;
  const statusClass = status === "PASS" ? "status-pass" : status === "FAIL" ? "status-fail" : "status-pending";
  const messageIcon: AppIconName = message?.kind === "success" ? "success" : message?.kind === "error" ? "error" : "info";

  if (view === "arrange" && scenario) {
    return (
      <EvidenceArranger
        key={arrangement?.groups.map((group) => group.key).join("|") ?? "loading"}
        scenarioId={scenario.scenarioId}
        testTypeName={selectedProfile.displayName}
        groups={arrangement?.groups ?? []}
        hasManualOrder={arrangement?.hasManualOrder ?? false}
        busy={operationBusy}
        loading={operationBusy && arrangement === null && !arrangerError}
        error={arrangerError}
        onApply={handleApplyArrangement}
        onReset={handleResetArrangement}
        onBack={handleCloseArranger}
      />
    );
  }

  if (view === "convert" && scenario) {
    return (
      <ExistingEvidenceConverter
        scenario={scenario}
        testTypeId={testTypeId}
        testTypeName={selectedProfile.displayName}
        moduleIds={selectedProfile.moduleIds}
        onAdoptHeader={handleAdoptExistingHeader}
        onConverted={handleConvertedExistingEvidence}
        onBack={() => {
          viewRef.current = "insert";
          setView("insert");
        }}
      />
    );
  }

  return (
    <main className="app-shell">
      <header className="app-header">
        <p className="eyebrow">Excel screenshot organizer</p>
        <h1 className="content-title">Insert Evidence</h1>
        <label className="test-type-label" htmlFor="test-type">Test Type</label>
        <div className="test-type-actions">
          <select
            id="test-type"
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
              setPendingAdoptedHeader(null);
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
          <a
            className="button button-secondary header-action-button help-link"
            href={import.meta.env.BASE_URL}
            target="_blank"
            rel="noopener noreferrer"
            aria-label="Help"
            title="Help"
          >
            <AppIcon name="help" />
          </a>
          <button
            className="button button-secondary header-action-button refresh-button"
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
        </div>
      </header>

      {availableVersion ? (
        <aside className="update-notice" role="status" aria-live="polite">
          <span>
            {pendingImage
              ? "New version available. Insert or remove pending screenshot before updating."
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

      <div className="evidence-tools">
        <button
          className="button button-secondary arrange-evidence-button"
          type="button"
          disabled={busy || !scenario}
          onClick={() => void handleOpenArranger()}
        >
          Arrange Evidence
        </button>
        <button
          className="button button-secondary"
          type="button"
          disabled={busy || !scenario}
          onClick={handleOpenConverter}
        >
          Convert Existing
        </button>
      </div>

      <section className="panel scenario-panel" aria-labelledby="scenario-heading">
        <h2 id="scenario-heading"><AppIcon name="scenario" />Test Scenario</h2>
        <dl className="scenario-id">
          <dt>Scenario ID</dt>
          <dd>{scenario?.scenarioId || "Not loaded"}</dd>
        </dl>
        <details>
          <summary>Condition</summary>
          <p className="workbook-preview">{scenario?.condition || "—"}</p>
        </details>
        <details>
          <summary>Expected Result</summary>
          <p className="workbook-preview">{scenario?.expectedResult || "—"}</p>
        </details>
      </section>

      <section className="panel status-panel" aria-labelledby="status-heading">
        <h2 id="status-heading"><AppIcon name="status" />Worksheet Status</h2>
        <label className="field status-field">
          <span>Status</span>
          <select className={statusClass}
            disabled={busy || !scenario?.statusCellAddress}
            value={status}
            onChange={(event) => void handleStatusChange(event.target.value as Exclude<EvidenceStatus, "UNCHANGED">)}
          >
            <option value="PASS/FAIL">PASS/FAIL</option>
            <option value="PASS">PASS</option>
            <option value="FAIL">FAIL</option>
          </select>
        </label>
      </section>

      <section className="panel capture-panel" aria-labelledby="capture-heading">
        <h2 id="capture-heading"><AppIcon name="image" />Screenshot to Insert</h2>
        <label className="field">
          <span>Evidence Type</span>
          <select
            value={evidenceType}
            disabled={busy}
            onChange={(event) => {
              const nextType = event.target.value as ModuleId;
              setEvidenceType(nextType);
              if (pendingAdoptedHeader?.moduleId !== nextType) setPendingAdoptedHeader(null);
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
        {pendingAdoptedHeader ? (
          <div className="adopted-header-notice" role="status">
            <span>Using existing header {pendingAdoptedHeader.header.address} for the next insertion.</span>
            <button className="button button-secondary" type="button" disabled={busy} onClick={() => setPendingAdoptedHeader(null)}>Cancel</button>
          </div>
        ) : null}
        <div className={`paste-zone ${pendingImage ? "has-image" : ""}`} tabIndex={0} aria-label="Screenshot paste area">
          {pendingImage ? (
            <>
              <div className="preview-media">
                <img src={pendingImage.previewUrl} alt="Pending screenshot preview" />
                <button
                  className="button button-secondary preview-clear-button"
                  type="button"
                  aria-label="Remove pending screenshot"
                  title="Remove pending screenshot"
                  disabled={busy}
                  onClick={clearPending}
                >
                  <AppIcon name="close" />
                </button>
              </div>
              <div className="preview-metadata">
                <p>{pendingImage.mimeType.replace("image/", "").toUpperCase()} · {(pendingImage.size / 1024).toFixed(1)} KB</p>
                {pendingWorkbookPicture ? <p>Worksheet picture: {pendingWorkbookPicture.name}</p> : null}
              </div>
            </>
          ) : (
            <>
              <AppIcon name="image" className="paste-placeholder-icon" />
              <strong>Press Ctrl+V</strong>
              <span>Paste a PNG or JPEG screenshot, maximum {IMAGE_LIMITS.maximumBytes / 1024 / 1024} MB.</span>
            </>
          )}
        </div>
        <div className="capture-source-actions">
          <button
            className="button button-secondary"
            type="button"
            disabled={busy}
            onClick={() => void handleUsePastedPicture()}
          >
            <AppIcon name="paste" />
            {pictureSelectionArmed ? "Cancel picture selection" : "Use pasted Excel picture"}
          </button>
        </div>
        <p className="selection-help">
          Copy in workbook A, paste onto this scenario sheet in workbook B, then select the pasted picture here.
        </p>
      </section>

      <div className="actions">
        <button
          className="button button-primary"
          type="button"
          disabled={busy || !scenario || !pendingImage}
          onClick={() => void handleInsert()}
        >
          <AppIcon name="plus" />
          Insert
        </button>
        <button className="button button-danger" type="button" disabled={busy || !scenario} onClick={() => void handleRemove()}>
          <AppIcon name="trash" />
          Remove Last Inserted
        </button>
      </div>

      <div className={`message message-${message?.kind ?? "info"}`} role={message?.kind === "error" ? "alert" : "status"}>
        <AppIcon name={messageIcon} />
        {message?.text || "Ready."}
      </div>
      {preferenceWarning ? <div className="message message-warning" role="status">{preferenceWarning}</div> : null}
    </main>
  );
}
