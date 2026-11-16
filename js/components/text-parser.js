const RE_PAGE_FOOTER = /^page\s+\d+\s+sur\s+\d+/i;
const RE_INTRODUCTION = /^introduction\./i;
const RE_PARTIE = /^partie\s+\d+/i;
const RE_THEME = /^(?:th[eèé]me|titre)\s+\d+/i;
const RE_CHAPITRE = /^chapitre\s+\d+/i;
const RE_SUBNUM = /^(\d+(?:\.\d+)+)\.?\s+(.+)$/;
const RE_PLAIN_NUM = /^(\d+)\.\s+(.+)$/;
const RE_ROMAN_OR_LETTER = /^([IVXLCDM]+|[A-Z])\.\s+(.+)$/;
const RE_LOWER_LETTER = /^([a-z])\.\s+(.+)$/;
const RE_BULLET = /^[-–—*]\s+(.+)$/;
const RE_ARROW = /^→\s*(.+)$/;
const RE_QUOTE = /^[«"](.+)$/;

export function isMostlyUppercase(str) {
  const letters = str.replace(/[^A-Za-zÀ-ÿ]/g, "");
  return letters.length > 0 && letters === letters.toUpperCase();
}

export function classifyLine(line) {
  if (RE_PAGE_FOOTER.test(line)) return { type: "skip", content: "" };
  if (RE_INTRODUCTION.test(line)) return { type: "grandTitre-intro", content: line.toUpperCase() };
  if (RE_PARTIE.test(line)) return { type: "grandTitre-partie", content: line };
  if (RE_THEME.test(line)) return { type: "grandTitre-partie", content: line };
  if (RE_CHAPITRE.test(line)) return { type: "chapitre", content: line };

  let m;
  m = line.match(RE_SUBNUM);
  if (m) return { type: "sousSousSection", content: line };

  m = line.match(RE_ROMAN_OR_LETTER);
  if (m) {
    const marker = m[1];
    const rest = m[2];
    const isRoman = /^[IVXLCDM]+$/.test(marker);
    if (marker.length > 1 && isRoman) return { type: "section", content: line };
    if (marker.length === 1 && isRoman) {
      return isMostlyUppercase(rest) ? { type: "section", content: line } : { type: "sousSection", content: line };
    }
    return { type: "sousSection", content: line };
  }

  m = line.match(RE_PLAIN_NUM);
  if (m) return { type: "section", content: line };

  m = line.match(RE_LOWER_LETTER);
  if (m) return { type: "detail", content: line };

  if (RE_BULLET.test(line)) return { type: "bullet", content: line.replace(RE_BULLET, "$1") };
  if (RE_ARROW.test(line)) return { type: "arrow", content: "→ " + line.replace(RE_ARROW, "$1") };
  if (RE_QUOTE.test(line)) return { type: "citation", content: line };

  return { type: "text", content: line };
}