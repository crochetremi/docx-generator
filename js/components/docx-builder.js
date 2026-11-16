const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, UnderlineType, ShadingType,
  AlignmentType, BorderStyle
} = window.docx;

import { classifyLine, isMostlyUppercase } from "./text-parser.js";

function getStyleConfig(prefix) {
  const colorInput = document.getElementById(`${prefix}-color`);
  const sizeInput = document.getElementById(`${prefix}-size`);
  const boldInput = document.getElementById(`${prefix}-bold`);
  const italicInput = document.getElementById(`${prefix}-italic`);

  const runProps = {
    color: colorInput ? colorInput.value.replace("#", "") : "000000",
    size: sizeInput ? Math.round(parseFloat(sizeInput.value) * 2) : 22,
    bold: boldInput ? boldInput.checked : false,
    italics: italicInput ? italicInput.checked : false
  };
  return { runProps };
}

function getGrandTitreConfig() {
  const base = getStyleConfig("theme");
  return {
    runProps: base.runProps,
    borderEnabled: true,
    borderColor: "2e75b6",
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "eaf1fa" }
  };
}

function getHighlightConfig() {
  const enabledInput = document.getElementById("highlight-enabled");
  return {
    enabled: enabledInput ? enabledInput.checked : false,
    articlesLois: { color: "b5651d", bold: true },
    jurisprudence: { color: "2f6b3a", bold: true, italics: true }
  };
}

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

function getLegendConfig() {
  const enabledInput = document.getElementById("legende-enabled");
  const bgEnabledInput = document.getElementById("legende-bg-enabled");
  const bgColorInput = document.getElementById("legende-bg-color");
  const borderEnabledInput = document.getElementById("legende-border-enabled");
  const borderColorInput = document.getElementById("legende-border-color");

  return {
    enabled: enabledInput ? enabledInput.checked : false,
    bgEnabled: bgEnabledInput ? bgEnabledInput.checked : false,
    bgColor: bgColorInput ? bgColorInput.value.replace("#", "") : "e3e9e2",
    borderEnabled: borderEnabledInput ? borderEnabledInput.checked : false,
    borderColor: borderColorInput ? borderColorInput.value.replace("#", "") : "3f5b48",
    notions: { color: "2f6b3a", bold: true },
    exceptions: { color: "b5651d", bold: true, italics: true }
  };
}

function getFontFamily() {
  const select = document.getElementById("font-family");
  return select ? select.value : "Calibri";
}

const ARTICLES_REGEX_SRC = "articles?\\s+[A-Z]\\.?\\s?\\d[\\d\\-.]*";
const JURIS_REGEX_SRC = "Cass\\.\\s*(?:soc|com|civ|crim)\\.?,?\\s*\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}";
const HIGHLIGHT_REGEX = new RegExp(`(?<articles>${ARTICLES_REGEX_SRC})|(?<juris>${JURIS_REGEX_SRC})`, "gi");

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

function buildBorder(color, size = 6, space = 8) {
  return { color, space, style: BorderStyle.SINGLE, size };
}

function buildGrandTitreParagraph(firstLine, secondLine, config, fontFamily) {
  const runs = [new TextRun({ text: firstLine.toUpperCase(), font: fontFamily, ...config.runProps })];
  if (secondLine) {
    runs.push(new TextRun({ text: secondLine.toUpperCase(), font: fontFamily, ...config.runProps, break: 1 }));
  }
  const border = config.borderEnabled ? {
    top: buildBorder(config.borderColor), bottom: buildBorder(config.borderColor),
    left: buildBorder(config.borderColor), right: buildBorder(config.borderColor)
  } : undefined;

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 320 },
    border,
    shading: config.shading,
    children: runs
  });
}

function buildLegendParagraphs(coverConfig, legendConfig, highlightConfig, fontFamily) {
  const paragraphs = [];
  if (coverConfig.title) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: coverConfig.title, font: fontFamily, bold: true, size: 64, color: coverConfig.themeColor })]
    }));
  }
  if (coverConfig.subtitle) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 360 },
      children: [new TextRun({ text: coverConfig.subtitle, font: fontFamily, italics: true, size: 28, color: "666666" })]
    }));
  }
  return paragraphs;
}

function buildParagraph(type, content, styles, highlightConfig, fontFamily) {
  const isTitle = ["theme", "chapitre", "section", "sousSection", "sousSousSection"].includes(type);
  const finalContent = isTitle ? content.toUpperCase() : content;
  const { runProps } = styles[type] || styles.text;

  return new Paragraph({
    spacing: { before: 100, after: 100 },
    children: buildTextRuns(finalContent, runProps, highlightConfig, fontFamily)
  });
}

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

export function initDocxBuilder() {
  const handleGeneration = async () => {
    const rawText = document.getElementById("raw-text").value;
    if (!rawText.trim()) {
      alert("Veuillez coller du texte d'abord.");
      return;
    }

    const fontFamily = getFontFamily();
    const styles = {
      theme: getStyleConfig("theme"),
      chapitre: getStyleConfig("chapitre"),
      section: getStyleConfig("section"),
      sousSection: getStyleConfig("sousSection"),
      sousSousSection: getStyleConfig("sousSection"),
      detail: getStyleConfig("detail"),
      citation: getStyleConfig("citation"),
      text: getStyleConfig("text"),
      bullet: getStyleConfig("text"),
      arrow: getStyleConfig("text")
    };

    const grandTitreConfig = getGrandTitreConfig();
    const highlightConfig = getHighlightConfig();
    const legendConfig = getLegendConfig();
    const coverConfig = getCoverConfig();

    const documentChildren = [];
    documentChildren.push(...buildLegendParagraphs(coverConfig, legendConfig, highlightConfig, fontFamily));

    const rawLines = rawText.split("\n");
    for (const line of rawLines) {
      if (!line.trim()) continue;
      const cls = classifyLine(line.trim());
      if (cls.type === "skip") continue;
      documentChildren.push(buildParagraph(cls.type, cls.content, styles, highlightConfig, fontFamily));
    }

    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
        children: documentChildren
      }]
    });

    const blob = await Packer.toBlob(doc);
    triggerFileDownload(blob, "Cours_Formate.docx");
  };

  const btnGenerate = document.getElementById("btn-generate");
  const btnGenerateTop = document.getElementById("btn-generate-top");

  if (btnGenerate) btnGenerate.addEventListener("click", handleGeneration);
  if (btnGenerateTop) btnGenerateTop.addEventListener("click", handleGeneration);

  // Mise à jour dynamique de l'aperçu et du compteur de mots
  const rawTextEl = document.getElementById("raw-text");
  const titleEl = document.getElementById("doc-title");
  const subtitleEl = document.getElementById("doc-subtitle");
  const wordCountEl = document.getElementById("word-count");

  const updatePreview = () => {
    if (titleEl && document.getElementById("prev-title")) {
      document.getElementById("prev-title").textContent = titleEl.value || "DROIT DU TRAVAIL";
    }
    if (subtitleEl && document.getElementById("prev-subtitle")) {
      document.getElementById("prev-subtitle").textContent = subtitleEl.value || "Notes de cours";
    }
    if (rawTextEl && wordCountEl) {
      const words = rawTextEl.value.trim() ? rawTextEl.value.trim().split(/\s+/).length : 0;
      wordCountEl.textContent = `${words} mot${words > 1 ? 's' : ''}`;
    }
  };

  if (rawTextEl) rawTextEl.addEventListener("input", updatePreview);
  if (titleEl) titleEl.addEventListener("input", updatePreview);
  if (subtitleEl) subtitleEl.addEventListener("input", updatePreview);
}