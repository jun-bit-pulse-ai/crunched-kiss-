/** First-run coach marks for the Excel pane. Copy is kid-simple on purpose. */

export const TOUR_SEEN_KEY = "crunched-tour-seen";

export const TOUR_TARGETS = {
  thread: "thread",
  chips: "chips",
  composer: "composer",
  tools: "tools",
  newChat: "new-chat",
} as const;

export type TourTarget = (typeof TOUR_TARGETS)[keyof typeof TOUR_TARGETS];

export type TourStep = {
  id: string;
  target: TourTarget;
  title: string;
  body: string;
};

export const TOUR_STEPS: TourStep[] = [
  {
    id: "thread",
    target: TOUR_TARGETS.thread,
    title: "The story",
    body: "This is the chat. You talk. Crunched talks back. Newest words land at the bottom, like a comic strip.",
  },
  {
    id: "chips",
    target: TOUR_TARGETS.chips,
    title: "Ready-made questions",
    body: "These three buttons are starter questions. Tap one if you do not want to type. They match the 15-minute demo.",
  },
  {
    id: "composer",
    target: TOUR_TARGETS.composer,
    title: "Your box",
    body: "Type your own question here, then press Send. This tour never locks the box — you can type as soon as you skip.",
  },
  {
    id: "tools",
    target: TOUR_TARGETS.tools,
    title: "Little peek cards",
    body: "When Crunched peeks at the sheet, a green-edged card shows up. That is Crunched looking — not reading every cell. No cards yet? That is okay. They appear after you ask something.",
  },
  {
    id: "new-chat",
    target: TOUR_TARGETS.newChat,
    title: "Start over",
    body: "New chat wipes the story so you can ask something else. Same spreadsheet. Empty chat.",
  },
];

export type TourStore = {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
};

function defaultStore(): TourStore | null {
  try {
    if (typeof sessionStorage !== "undefined") {
      return sessionStorage;
    }
  } catch {
    // Private mode or a non-browser test runner.
  }
  return null;
}

export function hasSeenTour(store: TourStore | null = defaultStore()): boolean {
  try {
    return store?.getItem(TOUR_SEEN_KEY) === "1";
  } catch {
    return false;
  }
}

export function markTourSeen(store: TourStore | null = defaultStore()): void {
  try {
    store?.setItem(TOUR_SEEN_KEY, "1");
  } catch {
    // Memory-only: React state still closes the tour.
  }
}

export function nextTourIndex(index: number, length = TOUR_STEPS.length): number {
  return Math.min(index + 1, length - 1);
}

export function prevTourIndex(index: number): number {
  return Math.max(index - 1, 0);
}

export function isLastTourStep(index: number, length = TOUR_STEPS.length): boolean {
  return index >= length - 1;
}
