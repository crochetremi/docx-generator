// Extraction des classes nécessaires exposées par le bundle docx
const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, UnderlineType, ShadingType,
  AlignmentType, BorderStyle
} = window.docx;

/* ============================================================
   0. PALETTES DE COULEURS PRÉDÉFINIES
   ============================================================
   Chaque palette ne modifie QUE les couleurs (pas les tailles,
   pas les gras/italique) afin de ne pas écraser les préférences
   de mise en forme déjà choisies par l'utilisateur.
*/
const PALETTES = {
  bleu: {
    "theme-color": "#1f4e79", "theme-bg-color": "#d9e6f2",
    "chapitre-color": "#2e75b6", "chapitre-bg-color": "#eaf1fa",
    "section-color": "#2e75b6", "section-bg-color": "#f2f6fb",
    "sousSection-color": "#5b9bd5", "sousSection-bg-color": "#f5f8fc",
    "sousSousSection-color": "#7fb2e5",
    "detail-color": "#404040",
    "citation-color": "#555555", "citation-bg-color": "#f3f3f3",
    "text-color": "#262626", "bullet-color": "#262626",
    "grandTitre-color": "#1f4e79", "grandTitre-border-color": "#2e75b6", "grandTitre-bg-color": "#eaf1fa",
    "legende-bg-color": "#eaf1fa", "legende-border-color": "#2e75b6",
    "articlesLois-color": "#b5651d", "jurisprudence-color": "#1f4e79",
    "notions-color": "#2e75b6", "exceptions-color": "#b5651d"
  },
  vert: {
    "theme-color": "#1f3d2b", "theme-bg-color": "#dfe6e0",
    "chapitre-color": "#2f5233", "chapitre-bg-color": "#e6ece4",
    "section-color": "#2f5233", "section-bg-color": "#eef2ec",
    "sousSection-color": "#4f7a5c", "sousSection-bg-color": "#f2f6f1",
    "sousSousSection-color": "#6f9a7c",
    "detail-color": "#3f4a41",
    "citation-color": "#55605a", "citation-bg-color": "#eef1ee",
    "text-color": "#2b2b2b", "bullet-color": "#2b2b2b",
    "grandTitre-color": "#2f5233", "grandTitre-border-color": "#3f5b48", "grandTitre-bg-color": "#eef2ec",
    "legende-bg-color": "#e3e9e2", "legende-border-color": "#3f5b48",
    "articlesLois-color": "#b5651d", "jurisprudence-color": "#2f6b3a",
    "notions-color": "#2f6b3a", "exceptions-color": "#b5651d"
  },
  bordeaux: {
    "theme-color": "#5e1a24", "theme-bg-color": "#f0dfe1",
    "chapitre-color": "#7a2431", "chapitre-bg-color": "#f5e7e9",
    "section-color": "#7a2431", "section-bg-color": "#f8eef0",
    "sousSection-color": "#a4505c", "sousSection-bg-color": "#faf1f2",
    "sousSousSection-color": "#c07d85",
    "detail-color": "#4a4a4a",
    "citation-color": "#5a5a5a", "citation-bg-color": "#f4f0f0",
    "text-color": "#262626", "bullet-color": "#262626",
    "grandTitre-color": "#5e1a24", "grandTitre-border-color": "#7a2431", "grandTitre-bg-color": "#f5e7e9",
    "legende-bg-color": "#f0dfe1", "legende-border-color": "#7a2431",
    "articlesLois-color": "#8a5a00", "jurisprudence-color": "#5e1a24",
    "notions-color": "#7a2431", "exceptions-color": "#8a5a00"
  },
  violet: {
    "theme-color": "#4b0082", "theme-bg-color": "#e6e6fa",
    "chapitre-color": "#6a0dad", "chapitre-bg-color": "#f3e5f5",
    "section-color": "#6a0dad", "section-bg-color": "#f8f0fc",
    "sousSection-color": "#9370db", "sousSection-bg-color": "#fbf7fd",
    "sousSousSection-color": "#ba55d3",
    "detail-color": "#404040",
    "citation-color": "#555555", "citation-bg-color": "#f3f3f3",
    "text-color": "#262626", "bullet-color": "#262626",
    "grandTitre-color": "#4b0082", "grandTitre-border-color": "#6a0dad", "grandTitre-bg-color": "#e6e6fa",
    "legende-bg-color": "#f3e5f5", "legende-border-color": "#6a0dad",
    "articlesLois-color": "#b5651d", "jurisprudence-color": "#4b0082",
    "notions-color": "#6a0dad", "exceptions-color": "#b5651d"
  },
  automnale: {
    "theme-color": "#5D4037", "theme-bg-color": "#EFEBE9", 
    "chapitre-color": "#795548", "chapitre-bg-color": "#F5F5F6",  
    "section-color": "#8D6E63", "section-bg-color": "#FAF9F9", 
    "sousSection-color": "#A1887F", "sousSection-bg-color": "#FCFCFC",
    "sousSousSection-color": "#BCAAA4", 
    "detail-color": "#4E342E", 
    "citation-color": "#5C5C5C", "citation-bg-color": "#F3F0EF", 
    "text-color": "#2C2C2C", "bullet-color": "#2C2C2C", 
    "grandTitre-color": "#5D4037", "grandTitre-border-color": "#8D6E63", "grandTitre-bg-color": "#EFEBE9", 
    "legende-bg-color": "#F5F5F6", "legende-border-color": "#795548",
    "articlesLois-color": "#D84315", "jurisprudence-color": "#5D4037", 
    "notions-color": "#795548", "exceptions-color": "#BF360C" 
  }
};

function applyPalette(name) {
  const palette = PALETTES[name];
  if (!palette) return;
  for (const [id, value] of Object.entries(palette)) {
    const el = document.getElementById(id);
    if (el) el.value = value;
  }
}

document.getElementById("btn-apply-palette").addEventListener("click", () => {
  const select = document.getElementById("palette-select");
  applyPalette(select.value);
});

// Applique la palette sélectionnée par défaut dès le chargement de la page
applyPalette(document.getElementById("palette-select").value);

/* ============================================================
   1. LECTURE DES RÉGLAGES DE STYLE DANS LE PANNEAU DE GAUCHE
   ============================================================ */

function getStyleConfig(prefix) {
  const colorInput = document.getElementById(`${prefix}-color`);
  const sizeInput = document.getElementById(`${prefix}-size`);
  const boldInput = document.getElementById(`${prefix}-bold`);
  const italicInput = document.getElementById(`${prefix}-italic`);
  const underlineInput = document.getElementById(`${prefix}-underline`);
  const bgEnabledInput = document.getElementById(`${prefix}-bg-enabled`);
  const bgColorInput = document.getElementById(`${prefix}-bg-color`);

  const runProps = {
    color: colorInput ? colorInput.value.replace("#", "") : "000000",
    size: sizeInput ? Math.round(parseFloat(sizeInput.value) * 2) : 22,
    bold: boldInput ? boldInput.checked : false,
    italics: italicInput ? italicInput.checked : false,
    underline: underlineInput && underlineInput.checked ? { type: UnderlineType.SINGLE } : undefined
  };

  let shading;
  if (bgEnabledInput && bgEnabledInput.checked && bgColorInput) {
    shading = { type: ShadingType.CLEAR, color: "auto", fill: bgColorInput.value.replace("#", "") };
  }

  return { runProps, shading };
}

/** Réglages du bloc "Titre encadré" (Introduction / Parties). */
function getGrandTitreConfig() {
  const base = getStyleConfig("grandTitre");
  const borderEnabledInput = document.getElementById("grandTitre-border-enabled");
  const borderColorInput = document.getElementById("grandTitre-border-color");
  return {
    runProps: base.runProps,
    shading: base.shading,
    borderEnabled: borderEnabledInput ? borderEnabledInput.checked : false,
    borderColor: borderColorInput ? borderColorInput.value.replace("#", "") : "000000"
  };
}

/** Réglages de la mise en valeur automatique (articles de loi / jurisprudence). */
function getHighlightConfig() {
  const enabledInput = document.getElementById("highlight-enabled");
  const readCategory = (prefix) => {
    const colorInput = document.getElementById(`${prefix}-color`);
    const boldInput = document.getElementById(`${prefix}-bold`);
    const italicInput = document.getElementById(`${prefix}-italic`);
    return {
      color: colorInput ? colorInput.value.replace("#", "") : "000000",
      bold: boldInput ? boldInput.checked : false,
      italics: italicInput ? italicInput.checked : false
    };
  };
  return {
    enabled: enabledInput ? enabledInput.checked : false,
    articlesLois: readCategory("articlesLois"),
    jurisprudence: readCategory("jurisprudence")
  };
}

/** Réglages de l'en-tête du document (Titre, sous-titre, modalités) */
function getCoverConfig() {
  const titleInput = document.getElementById("doc-title");
  const subtitleInput = document.getElementById("doc-subtitle");
  const modalitiesInput = document.getElementById("doc-modalities");
  const themeColorInput = document.getElementById("theme-color");

  return {
    title: titleInput ? titleInput.value.trim() : "",
    subtitle: subtitleInput ? subtitleInput.value.trim() : "",
    modalities: modalitiesInput ? modalitiesInput.value.trim() : "",
    themeColor: themeColorInput ? themeColorInput.value.replace("#", "") : "1f4e79"
  };
}

/** Réglages de l'encadré "Légende" inséré en début de document. */
function getLegendConfig() {
  const enabledInput = document.getElementById("legende-enabled");
  const bgEnabledInput = document.getElementById("legende-bg-enabled");
  const bgColorInput = document.getElementById("legende-bg-color");
  const borderEnabledInput = document.getElementById("legende-border-enabled");
  const borderColorInput = document.getElementById("legende-border-color");
  const readCategory = (prefix) => {
    const colorInput = document.getElementById(`${prefix}-color`);
    const boldInput = document.getElementById(`${prefix}-bold`);
    const italicInput = document.getElementById(`${prefix}-italic`);
    return {
      color: colorInput ? colorInput.value.replace("#", "") : "000000",
      bold: boldInput ? boldInput.checked : false,
      italics: italicInput ? italicInput.checked : false
    };
  };
  return {
    enabled: enabledInput ? enabledInput.checked : false,
    bgEnabled: bgEnabledInput ? bgEnabledInput.checked : false,
    bgColor: bgColorInput ? bgColorInput.value.replace("#", "") : "ffffff",
    borderEnabled: borderEnabledInput ? borderEnabledInput.checked : false,
    borderColor: borderColorInput ? borderColorInput.value.replace("#", "") : "000000",
    notions: readCategory("notions"),
    exceptions: readCategory("exceptions")
  };
}

function getFontFamily() {
  const select = document.getElementById("font-family");
  return select ? select.value : "Calibri";
}

/* ============================================================
   2. DÉTECTION DES RÉFÉRENCES LÉGALES DANS UNE LIGNE DE TEXTE
   ============================================================ */

   // Articles
const ARTICLES_REGEX_SRC =
  "articles?\\s+[A-Z]\\.?\\s?\\d[\\d\\-.]*(?:,\\s*alinéa\\s*\\d+)?(?:\\s+et\\s+suivants)?(?:\\s+du\\s+Code\\s+du\\s+travail)?";
  // Jurisprudences
const JURIS_REGEX_SRC =
  "Cass\\.\\s*(?:soc|com|civ|crim)\\.?,?\\s*\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}(?:,?\\s*n°\\s*[\\d\\-.,\\s]+\\d)?" +
  "|CJUE,?\\s*\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}(?:,\\s*(?:aff\\.\\s*)?C-[\\d/]+)?" +
  "|C\\.?\\s?E\\.?,?\\s+\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}";

const HIGHLIGHT_REGEX = new RegExp(`(?<articles>${ARTICLES_REGEX_SRC})|(?<juris>${JURIS_REGEX_SRC})`, "gi");

/**
 * Construit une liste de TextRun pour un texte donné : le texte "normal"
 * suit baseRunProps, les portions reconnues comme article de loi ou comme
 * jurisprudence suivent respectivement highlightConfig.articlesLois /
 * highlightConfig.jurisprudence (si la détection est activée).
 */
function buildTextRuns(rawText, baseRunProps, highlightConfig, fontFamily) {
  if (!highlightConfig || !highlightConfig.enabled) {
    return [new TextRun({ text: rawText, font: fontFamily, ...baseRunProps })];
  }

  const runs = [];
  let lastIndex = 0;
  let match;
  HIGHLIGHT_REGEX.lastIndex = 0;

  while ((match = HIGHLIGHT_REGEX.exec(rawText)) !== null) {
    if (match.index > lastIndex) {
      runs.push(new TextRun({ text: rawText.slice(lastIndex, match.index), font: fontFamily, ...baseRunProps }));
    }
    const category = match.groups && match.groups.articles ? highlightConfig.articlesLois : highlightConfig.jurisprudence;
    runs.push(new TextRun({ text: match[0], font: fontFamily, ...baseRunProps, ...category }));
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < rawText.length) {
    runs.push(new TextRun({ text: rawText.slice(lastIndex), font: fontFamily, ...baseRunProps }));
  }
  if (runs.length === 0) {
    runs.push(new TextRun({ text: rawText, font: fontFamily, ...baseRunProps }));
  }
  return runs;
}

/* ============================================================
   3. CLASSIFICATION D'UNE LIGNE DE TEXTE BRUT
   ============================================================ */

const RE_PAGE_FOOTER = /^page\s+\d+\s+sur\s+\d+/i;
const RE_INTRODUCTION = /^introduction\./i;
const RE_PARTIE = /^partie\s+\d+/i;
const RE_THEME = /^(?:th[eèé]me|titre)\s+\d+/i;
const RE_CHAPITRE = /^chapitre\s+\d+/i;
const RE_SUBNUM = /^(\d+(?:\.\d+)+)\.?\s+(.+)$/;            // 1.1.  /  2.1.
const RE_PLAIN_NUM = /^(\d+)\.\s+(.+)$/;                     // 1.  /  2.
const RE_ROMAN_OR_LETTER = /^([IVXLCDM]+|[A-Z])\.\s+(.+)$/;  // I. / A. / C. ...
const RE_LOWER_LETTER = /^([a-z])\.\s+(.+)$/;                // a. / b. / c.
const RE_BULLET = /^[-–—*]\s+(.+)$/;
const RE_ARROW = /^→\s*(.+)$/;
const RE_QUOTE = /^[«"](.+)$/;

function isMostlyUppercase(str) {
  const letters = str.replace(/[^A-Za-zÀ-ÿ]/g, "");
  return letters.length > 0 && letters === letters.toUpperCase();
}

/**
 * Analyse une ligne et retourne { type, content }.
 * type ∈ 'skip' | 'grandTitre-intro' | 'grandTitre-partie' | 'theme' | 'chapitre'
 *       | 'section' | 'sousSection' | 'sousSousSection' | 'detail'
 *       | 'citation' | 'bullet' | 'text'
 */
function classifyLine(line) {
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
      return isMostlyUppercase(rest)
        ? { type: "section", content: line }
        : { type: "sousSection", content: line };
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

/* ============================================================
   4. CONSTRUCTION DES PARAGRAPHES DOCX
   ============================================================ */

function buildBorder(color, size = 6, space = 8) {
  const style = BorderStyle.SINGLE;
  return { color, space, style, size };
}

/** Construit le paragraphe encadré "INTRODUCTION" (+ sous-titre optionnel), "PARTIE N" ou "THÈME/TITRE N". */
function buildGrandTitreParagraph(firstLine, secondLine, config, fontFamily) {
  const runs = [new TextRun({ text: firstLine.toUpperCase(), font: fontFamily, ...config.runProps })];
  
  if (secondLine) {
    runs.push(new TextRun({ text: secondLine.toUpperCase(), font: fontFamily, ...config.runProps, break: 1 }));
  }
  
  const border = config.borderEnabled
    ? { 
        top: buildBorder(config.borderColor), 
        bottom: buildBorder(config.borderColor),
        left: buildBorder(config.borderColor),
        right: buildBorder(config.borderColor)
      }
    : undefined;
    
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 320 },
    border,
    shading: config.shading,
    children: runs
  });
}

/** Construit l'en-tête complet et les paragraphes de l'encadré "Légende" */
function buildLegendParagraphs(coverConfig, legendConfig, highlightConfig, fontFamily) {
  const paragraphs = [];

  // --- 1. TITRE ET SOUS-TITRE (Au-dessus de l'encadré) ---
  if (coverConfig.title) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: coverConfig.title, font: fontFamily, bold: true, size: 64, color: coverConfig.themeColor })]
    }));
  }
  
  if (coverConfig.subtitle) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 360 },
      children: [new TextRun({ text: coverConfig.subtitle, font: fontFamily, italics: true, size: 28, color: "666666" })]
    }));
  }

  // Si pas de légende et pas de modalités, on s'arrête là
  if (!legendConfig.enabled && !coverConfig.modalities) {
    return paragraphs;
  }

  // --- 2. CONFIGURATION DE L'ENCADRÉ (Fond et bordures) ---
  const shading = legendConfig.bgEnabled
    ? { type: ShadingType.CLEAR, color: "auto", fill: legendConfig.bgColor }
    : undefined;

  // --- 3. MODALITÉS DU COURS (Dans l'encadré) ---
  if (coverConfig.modalities) {
    paragraphs.push(new Paragraph({
      spacing: { before: 100, after: 100 },
      shading,
      // La bordure supérieure s'applique ici si les modalités existent
      border: legendConfig.borderEnabled ? { top: buildBorder(legendConfig.borderColor) } : undefined,
      children: [
        new TextRun({ text: "À PRÉVOIR AVANT LE COURS", font: fontFamily, bold: true, size: 24, color: coverConfig.themeColor })
      ]
    }));
    
    paragraphs.push(new Paragraph({
      spacing: { before: 0, after: 300 },
      shading,
      children: [
        new TextRun({ text: coverConfig.modalities, font: fontFamily, italics: true, size: 22, color: "444444" })
      ]
    }));
  }

  // --- 4. LA LÉGENDE CLASSIQUE (Dans le même encadré) ---
  if (legendConfig.enabled) {
    const categories = [
      { symbol: "■", label: "Notions et définitions", desc: "vocabulaire juridique et notions clés à maîtriser", style: legendConfig.notions },
      { symbol: "■", label: "Articles, lois, textes normatifs", desc: "références légales et réglementaires", style: highlightConfig.articlesLois },
      { symbol: "■", label: "Jurisprudence, arrêts et décisions", desc: "solutions dégagées par les juridictions", style: highlightConfig.jurisprudence },
      { symbol: "■", label: "Exceptions et nuances", desc: "tempéraments et nuances apportés au principe général", style: legendConfig.exceptions }
    ];

    paragraphs.push(new Paragraph({
      // Si les modalités sont absentes, on met plus d'espace avant. Sinon, on colle à la suite.
      spacing: { before: coverConfig.modalities ? 0 : 40, after: 100 },
      shading,
      // Si les modalités étaient absentes, la bordure supérieure doit s'appliquer sur le titre "Légende"
      border: (!coverConfig.modalities && legendConfig.borderEnabled) ? { top: buildBorder(legendConfig.borderColor) } : undefined,
      children: [new TextRun({ text: "Légende", font: fontFamily, bold: true, size: 24, color: "1A202C" })]
    }));

    categories.forEach((cat, idx) => {
      const isLast = idx === categories.length - 1;
      paragraphs.push(new Paragraph({
        spacing: { before: 20, after: 20 },
        shading,
        border: (isLast && legendConfig.borderEnabled) ? { bottom: buildBorder(legendConfig.borderColor) } : undefined,
        children: [
          new TextRun({ text: `${cat.symbol} `, font: fontFamily, color: cat.style.color }),
          new TextRun({ text: cat.label, font: fontFamily, bold: cat.style.bold, italics: cat.style.italics, color: cat.style.color }),
          new TextRun({ text: ` — ${cat.desc}`, font: fontFamily, color: "555555" })
        ]
      }));
    });
  }

  // Espace de respiration à la fin de l'encadré
  paragraphs.push(new Paragraph({ spacing: { before: 0, after: 300 }, children: [] }));

  return paragraphs;
}
function buildParagraph(type, content, styles, highlightConfig, fontFamily) {
  // 1. Détection des titres et conversion en majuscules (DRY: Don't Repeat Yourself)
  const isTitle = ["theme", "chapitre", "section", "sousSection", "sousSousSection"].includes(type);
  const finalContent = isTitle ? content.toUpperCase() : content;

  switch (type) {
    case "theme": {
      const { runProps, shading } = styles.theme;
      return new Paragraph({
        heading: HeadingLevel.TITLE,
        spacing: { before: 360, after: 200 },
        shading,
        children: [new TextRun({ text: finalContent, font: fontFamily, ...runProps })]
      });
    }
    case "chapitre": {
      const { runProps, shading } = styles.chapitre;
      // 2. Création d'un encadré complet pour le chapitre avec la couleur du texte
      const chapBorder = {
        top: buildBorder(runProps.color),
        bottom: buildBorder(runProps.color),
        left: buildBorder(runProps.color),
        right: buildBorder(runProps.color)
      };
      return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 280, after: 140 },
        shading,
        border: chapBorder,
        children: [new TextRun({ text: finalContent, font: fontFamily, ...runProps })]
      });
    }
    case "section": {
      const { runProps, shading } = styles.section;
      return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 220, after: 110 },
        shading,
        children: [new TextRun({ text: finalContent, font: fontFamily, ...runProps })]
      });
    }
    case "sousSection": {
      const { runProps, shading } = styles.sousSection;
      return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 },
        shading,
        children: [new TextRun({ text: finalContent, font: fontFamily, ...runProps })]
      });
    }
    case "sousSousSection": {
      const { runProps } = styles.sousSousSection;
      return new Paragraph({
        heading: HeadingLevel.HEADING_4,
        spacing: { before: 160, after: 80 },
        children: [new TextRun({ text: finalContent, font: fontFamily, ...runProps })]
      });
    }
    case "detail": {
      const { runProps } = styles.detail;
      return new Paragraph({
        heading: HeadingLevel.HEADING_5,
        spacing: { before: 120, after: 60 },
        // Les détails (a, b, c) restent en minuscules, on utilise 'content' brut
        children: [new TextRun({ text: content, font: fontFamily, ...runProps })]
      });
    }
    case "citation": {
      const { runProps, shading } = styles.citation;
      return new Paragraph({
        indent: { left: 480, right: 480 },
        spacing: { before: 100, after: 100 },
        shading,
        children: buildTextRuns(content, runProps, highlightConfig, fontFamily)
      });
    }
    case "bullet": {
      const { runProps } = styles.bullet;
      return new Paragraph({
        bullet: { level: 0 },
        spacing: { before: 40, after: 40 },
        children: buildTextRuns(content, runProps, highlightConfig, fontFamily)
      });
    }
    case "arrow": {
      const { runProps } = styles.arrow;
      return new Paragraph({
        indent: { left: 720 },
        spacing: { before: 40, after: 40 },
        children: buildTextRuns(content, runProps, highlightConfig, fontFamily)
      });
    }
    case "text":
    default: {
      const { runProps } = styles.text;
      return new Paragraph({
        spacing: { before: 60, after: 120, line: 276 },
        children: buildTextRuns(content, runProps, highlightConfig, fontFamily)
      });
    }
  }
}

/* ============================================================
   5. TÉLÉCHARGEMENT DU FICHIER GÉNÉRÉ
   ============================================================ */

function triggerFileDownload(blob, fileName) {
  const downloadUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = downloadUrl;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(downloadUrl);
}
/* ============================================================
   6. ÉCOUTEUR PRINCIPAL (GÉNÉRATION DU DOCUMENT)
   ============================================================ */

document.getElementById("btn-generate").addEventListener("click", async () => {
  const rawText = document.getElementById("raw-text").value;

  if (!rawText.trim()) {
    alert("Veuillez coller du texte d'abord.");
    return;
  }

  const fontFamily = getFontFamily();

  // Récupération de tous les styles depuis l'interface
  const styles = {
    theme: getStyleConfig("theme"),
    chapitre: getStyleConfig("chapitre"),
    section: getStyleConfig("section"),
    sousSection: getStyleConfig("sousSection"),
    sousSousSection: getStyleConfig("sousSousSection"),
    detail: getStyleConfig("detail"),
    citation: getStyleConfig("citation"),
    text: getStyleConfig("text"),
    bullet: getStyleConfig("bullet"),
    arrow: getStyleConfig("arrow") // N'oublie pas la flèche !
  };
  
  const grandTitreConfig = getGrandTitreConfig();
  const highlightConfig = getHighlightConfig();
  const legendConfig = getLegendConfig();
  const coverConfig = getCoverConfig(); 

  const documentChildren = [];

  // 1. En-tête et Encadré "Légende" en tout début de document
  documentChildren.push(...buildLegendParagraphs(coverConfig, legendConfig, highlightConfig, fontFamily));

  // 2. Découpage ligne par ligne, avec fusion et séparation des encadrés
  const rawLines = rawText.split("\n");
  let i = 0;
  
  // Variable "mémoire" pour savoir si on doit insérer un espace entre deux encadrés
  let lastWasBox = false; 

  while (i < rawLines.length) {
    const line = rawLines[i].trim();
    i++;
    if (!line) continue;

    const cls = classifyLine(line);
    if (cls.type === "skip") continue;

    // Gestion de l'Introduction
    if (cls.type === "grandTitre-intro") {
      let j = i;
      while (j < rawLines.length && rawLines[j].trim() === "") j++;
      let secondLine = null;
      if (j < rawLines.length) {
        const candidate = rawLines[j].trim();
        const candidateCls = classifyLine(candidate);
        if (candidateCls.type === "text" && isMostlyUppercase(candidate)) {
          secondLine = candidate;
          i = j + 1;
        }
      }
      
      if (lastWasBox) {
        // Ajout d'un paragraphe vide pour séparer deux encadrés consécutifs
        documentChildren.push(new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }));
      }
      
      documentChildren.push(buildGrandTitreParagraph(cls.content, secondLine, grandTitreConfig, fontFamily));
      lastWasBox = true;
      continue;
    }

    // Gestion des Parties et Thèmes/Titres (qui partagent le même style)
    if (cls.type === "grandTitre-partie") {
      if (lastWasBox) {
        // Ajout d'un paragraphe vide pour séparer deux encadrés consécutifs
        documentChildren.push(new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }));
      }
      
      documentChildren.push(buildGrandTitreParagraph(cls.content, null, grandTitreConfig, fontFamily));
      lastWasBox = true;
      continue;
    }

    // Pour tout le reste du texte (Chapitres, Sections, puces, texte normal)
    documentChildren.push(buildParagraph(cls.type, cls.content, styles, highlightConfig, fontFamily));
    
    // On réinitialise la mémoire car ce n'était pas un grand encadré
    lastWasBox = false; 
  }

  // 3. Construction finale de l'objet Document
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } // Marges de 2,5 cm
          }
        },
        children: documentChildren
      }
    ]
  });

  // 4. Conversion en fichier téléchargeable et déclenchement du téléchargement
  const blob = await Packer.toBlob(doc);
  triggerFileDownload(blob, "Cours_Formate.docx");
});