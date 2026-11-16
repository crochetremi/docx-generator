import { PALETTES } from "../../data/palettes.js";

export function applyPalette(name) {
  const palette = PALETTES[name];
  if (!palette) return;
  for (const [id, value] of Object.entries(palette)) {
    const el = document.getElementById(id);
    if (el) el.value = value;
  }
}

export function initTheme() {
  const select = document.getElementById("palette-select");
  const btn = document.getElementById("btn-apply-palette");
  
  if (select && btn) {
    applyPalette(select.value);
    btn.addEventListener("click", () => applyPalette(select.value));
  }
}