import { initTheme } from "./core/theme.js";
import { initDocxBuilder } from "./components/docx-builder.js";

document.addEventListener("DOMContentLoaded", () => {
  initTheme();
  initDocxBuilder();
});