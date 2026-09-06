import { createRoot } from "react-dom/client";
import App from "../app/App";
import "../app/styles.css";

/* global document, Office */

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

function renderApp() {
  root?.render(<App />);
}

Office.onReady(() => {
  renderApp();
});

window.addEventListener("load", () => {
  window.setTimeout(() => {
    if (rootElement && rootElement.childElementCount === 0) {
      renderApp();
    }
  }, 1500);
});
