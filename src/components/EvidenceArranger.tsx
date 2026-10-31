import { useEffect, useMemo, useState } from "react";
import {
  closestCenter,
  DndContext,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from "@dnd-kit/core";
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import type { EvidenceGroupDescriptor } from "../services/layoutEngine";

interface EvidenceArrangerProps {
  scenarioId: string;
  testTypeName: string;
  groups: readonly EvidenceGroupDescriptor[];
  hasManualOrder: boolean;
  busy: boolean;
  loading?: boolean;
  error?: string;
  onApply: (order: readonly string[]) => Promise<void>;
  onReset: () => Promise<void>;
  onBack: () => void;
}

const sameOrder = (first: readonly string[], second: readonly string[]): boolean =>
  first.length === second.length && first.every((key, index) => key === second[index]);

const SortableEvidenceGroup = ({
  group,
  position,
  total,
  busy,
  onMove,
}: {
  group: EvidenceGroupDescriptor;
  position: number;
  total: number;
  busy: boolean;
  onMove: (offset: -1 | 1) => void;
}) => {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
    id: group.key,
    disabled: busy || group.anchored,
  });
  return (
    <li
      ref={setNodeRef}
      className={`arranger-item ${isDragging ? "is-dragging" : ""}`}
      style={{ transform: CSS.Transform.toString(transform), transition }}
    >
      <span className="arranger-position" aria-label={`Position ${position}`}>{position}</span>
      <button
        className="drag-handle"
        type="button"
        disabled={busy || group.anchored}
        aria-label={`Drag ${group.label}`}
        {...attributes}
        {...listeners}
      >
        ⠿
      </button>
      <div className="arranger-item-copy">
        <strong>{group.label}</strong>
        <span>{group.imageCount} screenshot{group.imageCount === 1 ? "" : "s"}</span>
        <div className="arranger-badges">
          {group.kind === "paired" ? <span className="arranger-badge">Paired</span> : null}
          {!group.showHeading ? <span className="arranger-badge muted">No heading</span> : null}
          {group.anchored ? <span className="arranger-badge">Fixed in place</span> : null}
        </div>
      </div>
      <div className="arranger-move-buttons">
        <button
          className="button button-secondary"
          type="button"
          aria-label={`Move ${group.label} up`}
          disabled={busy || group.anchored || position === 1}
          onClick={() => onMove(-1)}
        >↑</button>
        <button
          className="button button-secondary"
          type="button"
          aria-label={`Move ${group.label} down`}
          disabled={busy || group.anchored || position === total}
          onClick={() => onMove(1)}
        >↓</button>
      </div>
    </li>
  );
};

export default function EvidenceArranger({
  scenarioId,
  testTypeName,
  groups,
  hasManualOrder,
  busy,
  loading = false,
  error,
  onApply,
  onReset,
  onBack,
}: EvidenceArrangerProps) {
  const initialOrder = useMemo(() => groups.map((group) => group.key), [groups]);
  const groupByKey = useMemo(() => new Map(groups.map((group) => [group.key, group])), [groups]);
  const [order, setOrder] = useState(initialOrder);
  const [confirmBack, setConfirmBack] = useState(false);
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates }),
  );

  useEffect(() => setOrder(initialOrder), [initialOrder]);

  const orderedGroups = order.flatMap((key) => {
    const group = groupByKey.get(key);
    return group ? [group] : [];
  });
  const dirty = !sameOrder(order, initialOrder);
  const hasAnchoredGroups = groups.some((group) => group.anchored);

  const handleDragEnd = ({ active, over }: DragEndEvent) => {
    if (!over || active.id === over.id) return;
    setOrder((current) => {
      const oldIndex = current.indexOf(String(active.id));
      const newIndex = current.indexOf(String(over.id));
      return oldIndex < 0 || newIndex < 0 ? current : arrayMove(current, oldIndex, newIndex);
    });
  };

  const requestBack = () => {
    if (dirty) setConfirmBack(true);
    else onBack();
  };

  return (
    <main className="app-shell arranger-shell">
      <header className="arranger-header">
        <button className="button button-secondary back-button" type="button" disabled={busy} onClick={requestBack}>
          ← Back
        </button>
        <p className="eyebrow">{scenarioId} · {testTypeName}</p>
        <h1 className="content-title">Arrange Evidence</h1>
        <p className="arranger-instructions">
          {hasAnchoredGroups
            ? "Converted sections stay fixed to their existing worksheet headers."
            : "Drag evidence sections into worksheet order."}
        </p>
      </header>

      {loading ? (
        <section className="panel empty-arranger" role="status">Loading evidence arrangement…</section>
      ) : groups.length > 0 ? (
        <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
          <SortableContext items={order} strategy={verticalListSortingStrategy}>
            <ol className="arranger-list">
              {orderedGroups.map((group, index) => (
                <SortableEvidenceGroup
                  key={group.key}
                  group={group}
                  position={index + 1}
                  total={orderedGroups.length}
                  busy={busy}
                  onMove={(offset) => setOrder((current) => arrayMove(current, index, index + offset))}
                />
              ))}
            </ol>
          </SortableContext>
        </DndContext>
      ) : (
        <section className="panel empty-arranger" role="status">No evidence exists for this scenario and test type.</section>
      )}

      {error ? <div className="message message-error" role="alert">{error}</div> : null}
      <div className="arranger-actions">
        <button
          className="button button-secondary"
          type="button"
          disabled={busy || hasAnchoredGroups || groups.length === 0 || (!hasManualOrder && !dirty)}
          onClick={() => {
            if (hasManualOrder) void onReset();
            else setOrder(initialOrder);
          }}
        >
          Reset to chronological
        </button>
        <button
          className="button button-primary"
          type="button"
          disabled={busy || hasAnchoredGroups || groups.length === 0 || !dirty}
          onClick={() => void onApply(order)}
        >
          Apply Arrangement
        </button>
      </div>

      {confirmBack ? (
        <div className="modal-backdrop">
          <section className="confirm-dialog" role="dialog" aria-modal="true" aria-labelledby="discard-title">
            <h2 id="discard-title">Discard arrangement changes?</h2>
            <p>Changes have not been applied.</p>
            <div className="confirm-actions">
              <button className="button button-danger" type="button" onClick={onBack}>Discard and go back</button>
              <button className="button button-secondary" type="button" autoFocus onClick={() => setConfirmBack(false)}>Cancel</button>
            </div>
          </section>
        </div>
      ) : null}
    </main>
  );
}
