import { test, describe } from "node:test";
import assert from "node:assert/strict";
import { summarizeGenerationStatus } from "./generation.js";

describe("summarizeGenerationStatus", () => {
  test("thumbnail outputs are excluded and never block completion", () => {
    // Regression test: poll_generation (App panel) and handleGetStatus (headless chat flow)
    // used to compute allDone independently, and only handleGetStatus filtered out thumbnail
    // outputs. A stalled/failed thumbnail made poll_generation report "Generating" forever
    // while handleGetStatus reported "Completed" for the same job.
    const raw = {
      id: "gid-1",
      status: 1, // Generating
      outputs: [
        { id: "o1", status: 2, format: "pptx", fileName: "Deck.pptx" },
        { id: "o2", status: 2, format: "pdf", fileName: "Deck.pdf" },
        { id: "o3", status: 1, format: "thumbnail" }, // still generating — must not block allDone
      ],
    };
    const summary = summarizeGenerationStatus(raw);
    assert.equal(summary.outputs.length, 2);
    assert.equal(summary.allDone, true);
    assert.equal(summary.overallStatus, "Completed");
  });

  test("not done while a real (non-thumbnail) output is still generating", () => {
    const raw = {
      id: "gid-2",
      status: 1,
      outputs: [
        { id: "o1", status: 2, format: "pptx", fileName: "Deck.pptx" },
        { id: "o2", status: 1, format: "pdf", fileName: "Deck.pdf" },
      ],
    };
    const summary = summarizeGenerationStatus(raw);
    assert.equal(summary.allDone, false);
    assert.equal(summary.overallStatus, "Generating");
  });

  test("overallStatus is Failed once all outputs are terminal and at least one failed", () => {
    const raw = {
      id: "gid-3",
      status: 1,
      outputs: [
        { id: "o1", status: 2, format: "pptx", fileName: "Deck.pptx" },
        { id: "o2", status: 3, format: "pdf", errorMessage: "render error" },
      ],
    };
    const summary = summarizeGenerationStatus(raw);
    assert.equal(summary.allDone, true);
    assert.equal(summary.overallStatus, "Failed");
    assert.equal(summary.outputs.find(o => o.format === "pdf")?.errorString, "render error");
  });

  test("job-level Failed short-circuits to Failed even before outputs finish", () => {
    const raw = {
      id: "gid-4",
      status: 3, // Failed
      outputs: [{ id: "o1", status: 0, format: "pptx" }],
    };
    const summary = summarizeGenerationStatus(raw);
    assert.equal(summary.allDone, false);
    assert.equal(summary.overallStatus, "Failed");
  });

  test("no outputs at all is not considered done", () => {
    const summary = summarizeGenerationStatus({ id: "gid-5", status: 0, outputs: [] });
    assert.equal(summary.allDone, false);
    assert.equal(summary.overallStatus, "Generating");
  });
});
