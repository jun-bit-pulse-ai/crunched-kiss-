import assert from "assert";
import {
  hasSeenTour,
  isLastTourStep,
  markTourSeen,
  nextTourIndex,
  prevTourIndex,
  TOUR_SEEN_KEY,
  TOUR_STEPS,
  TOUR_TARGETS,
  type TourStore,
} from "../src/app/tour";

function memoryStore(initial: Record<string, string> = {}): TourStore {
  const data = { ...initial };
  return {
    getItem(key: string) {
      return data[key] ?? null;
    },
    setItem(key: string, value: string) {
      data[key] = value;
    },
  };
}

describe("TOUR_STEPS", () => {
  it("is a short first-run tour with five kid-simple stops", () => {
    assert.strictEqual(TOUR_STEPS.length, 5);
    const targets = TOUR_STEPS.map((step) => step.target);
    assert.deepStrictEqual(targets, [
      TOUR_TARGETS.thread,
      TOUR_TARGETS.chips,
      TOUR_TARGETS.composer,
      TOUR_TARGETS.tools,
      TOUR_TARGETS.newChat,
    ]);
    for (const step of TOUR_STEPS) {
      assert.ok(step.title.length > 0);
      assert.ok(step.body.length > 0);
      assert.doesNotMatch(step.body, /leverage|utilize|synerg|stakeholder/i);
    }
  });

  it("explains tool cards even before any card exists", () => {
    const tools = TOUR_STEPS.find((step) => step.id === "tools");
    assert.ok(tools);
    assert.match(tools.body, /No cards yet/i);
  });
});

describe("tour index helpers", () => {
  it("walks forward and back without leaving the list", () => {
    assert.strictEqual(nextTourIndex(0), 1);
    assert.strictEqual(nextTourIndex(4), 4);
    assert.strictEqual(prevTourIndex(0), 0);
    assert.strictEqual(prevTourIndex(2), 1);
    assert.strictEqual(isLastTourStep(4), true);
    assert.strictEqual(isLastTourStep(3), false);
  });
});

describe("tour seen store", () => {
  it("auto-starts when the store is empty", () => {
    assert.strictEqual(hasSeenTour(memoryStore()), false);
  });

  it("remembers a seen tour on the injected store", () => {
    const store = memoryStore();
    markTourSeen(store);
    assert.strictEqual(store.getItem(TOUR_SEEN_KEY), "1");
    assert.strictEqual(hasSeenTour(store), true);
  });

  it("treats a missing store as unseen", () => {
    assert.strictEqual(hasSeenTour(null), false);
  });
});
