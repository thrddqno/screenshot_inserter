import { render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import App from "./App";

describe("task pane", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("renders capture controls and disables insertion without a screenshot", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    render(<App />);

    expect(screen.getByRole("heading", { name: "Test Evidence Inserter" })).toBeVisible();
    expect(screen.getByText("Press Ctrl+V")).toBeVisible();
    expect(screen.getByRole("button", { name: "Insert Evidence" })).toBeDisabled();
    const evidenceType = screen.getByLabelText("Evidence Type");
    expect(evidenceType).toContainHTML("Retailer Statement Duplicate");
    expect(screen.getByRole("option", { name: "PDF" })).toBeVisible();
    expect(await screen.findByRole("alert")).toHaveTextContent("Open this page inside an Excel task pane.");
  });
});
