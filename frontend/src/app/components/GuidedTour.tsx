import { useEffect, useId, useRef, useState } from "react";
import {
  isLastTourStep,
  nextTourIndex,
  prevTourIndex,
  TOUR_STEPS,
  type TourStep,
} from "../tour";

type SpotRect = {
  top: number;
  left: number;
  width: number;
  height: number;
};

type GuidedTourProps = {
  open: boolean;
  onClose: (reason: "skip" | "done") => void;
};

function measureSpot(target: string): SpotRect | null {
  const shell = document.querySelector(".shell");
  let el = document.querySelector(`[data-tour="${target}"]`);
  if (target === "tools") {
    el = document.querySelector(".tool-card") ?? document.querySelector('[data-tour="thread"]') ?? el;
  }
  if (!shell || !(el instanceof HTMLElement)) {
    return null;
  }
  const shellBox = shell.getBoundingClientRect();
  const box = el.getBoundingClientRect();
  const width = box.width;
  const height = box.height;
  if (width < 8 || height < 8) {
    return null;
  }
  return {
    top: box.top - shellBox.top,
    left: box.left - shellBox.left,
    width,
    height,
  };
}

function dockCardAtTop(step: TourStep): boolean {
  return step.target === "composer" || step.target === "chips";
}

export function GuidedTour({ open, onClose }: GuidedTourProps) {
  const [index, setIndex] = useState(0);
  const [spot, setSpot] = useState<SpotRect | null>(null);
  const cardRef = useRef<HTMLDivElement | null>(null);
  const titleId = useId();
  const bodyId = useId();

  const step = TOUR_STEPS[index];
  const last = isLastTourStep(index);

  useEffect(() => {
    if (open) {
      setIndex(0);
    }
  }, [open]);

  useEffect(() => {
    if (!open || !step) {
      setSpot(null);
      return;
    }

    const update = () => setSpot(measureSpot(step.target));
    update();

    const thread = document.querySelector(".thread");
    window.addEventListener("resize", update);
    thread?.addEventListener("scroll", update);
    const observer = typeof ResizeObserver === "function" ? new ResizeObserver(update) : null;
    observer?.observe(document.querySelector(".shell") ?? document.body);

    return () => {
      window.removeEventListener("resize", update);
      thread?.removeEventListener("scroll", update);
      observer?.disconnect();
    };
  }, [open, step]);

  useEffect(() => {
    if (!open) {
      return;
    }
    const primary = cardRef.current?.querySelector<HTMLButtonElement>("[data-tour-primary]");
    primary?.focus();
  }, [open, index]);

  useEffect(() => {
    if (!open) {
      return;
    }

    function onKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose("skip");
        return;
      }
      if (event.key === "ArrowRight") {
        event.preventDefault();
        if (isLastTourStep(index)) {
          onClose("done");
        } else {
          setIndex((current) => nextTourIndex(current));
        }
        return;
      }
      if (event.key === "ArrowLeft") {
        event.preventDefault();
        setIndex((current) => prevTourIndex(current));
        return;
      }
      if (event.key !== "Tab") {
        return;
      }
      const buttons = cardRef.current?.querySelectorAll<HTMLButtonElement>("button");
      if (!buttons || buttons.length === 0) {
        return;
      }
      const list = Array.from(buttons);
      const first = list[0];
      const lastButton = list[list.length - 1];
      const active = document.activeElement;
      if (event.shiftKey && active === first) {
        event.preventDefault();
        lastButton.focus();
      } else if (!event.shiftKey && active === lastButton) {
        event.preventDefault();
        first.focus();
      }
    }

    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [open, index, onClose]);

  if (!open || !step) {
    return null;
  }

  return (
    <div className="tour-overlay" aria-hidden="false">
      {spot ? (
        <div
          className="tour-spot"
          style={{
            top: spot.top,
            left: spot.left,
            width: spot.width,
            height: spot.height,
          }}
        />
      ) : null}
      <div
        ref={cardRef}
        className={`tour-card${dockCardAtTop(step) ? " tour-card-top" : ""}`}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        aria-describedby={bodyId}
      >
        <p className="tour-count">
          {index + 1} of {TOUR_STEPS.length}
        </p>
        <h2 id={titleId} className="tour-title">
          {step.title}
        </h2>
        <p id={bodyId} className="tour-body">
          {step.body}
        </p>
        <div className="tour-actions">
          <button type="button" className="tour-text-btn" onClick={() => onClose("skip")}>
            Skip
          </button>
          <button
            type="button"
            className="tour-text-btn"
            onClick={() => setIndex((current) => prevTourIndex(current))}
            disabled={index === 0}
          >
            Back
          </button>
          <button
            type="button"
            className="tour-next"
            data-tour-primary=""
            onClick={() => {
              if (last) {
                onClose("done");
              } else {
                setIndex((current) => nextTourIndex(current));
              }
            }}
          >
            {last ? "Done" : "Next"}
          </button>
        </div>
      </div>
    </div>
  );
}
