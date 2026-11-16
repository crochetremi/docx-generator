// Extraction des classes nécessaires exposées par le bundle docx
const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, UnderlineType, ShadingType,
  AlignmentType, BorderStyle
} = window.docx;

import { classifyLine, isMostlyUppercase } from "./text-parser.js";

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

const ARTICLES_REGEX_SRC =
  "articles?\\s+[A-Z]\\.?\\s?\\d[\\d\\-.]*(?:,\\s*alinéa\\s*\\d+)?(?:\\s+et\\s+suivants)?(?:\\s+du\\s+Code\\s+du\\s+travail)?";
const JURIS_REGEX_SRC =
  "Cass\\.\\s*(?:soc|com|civ|crim)\\.?,?\\s*\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}(?:,?\\s*n°\\s*[\\d\\-.,\\s]+\\d)?" +
  "|CJUE,?\\s*\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}(?:,\\s*(?:aff\\.\\s*)?C-[\\d/]+)?" +
  "|C\\.?\\s?E\\.?,?\\s+\\d{1,2}\\s+[a-zA-Zéûôàè]+\\s+\\d{4}";

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

/* ============================================================
   4. CONSTRUCTION DES PARAGRAPHES DOCX
   ============================================================ */

function buildBorder(color, size = 6, space = 8) {
  const style = BorderStyle.SINGLE;
  return { color, space, style, size };
}

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

function buildLegendParagraphs(coverConfig, legendConfig, highlightConfig, fontFamily) {
  const paragraphs = [];

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

  if (!legendConfig.enabled && !coverConfig.modalities) {
    return paragraphs;
  }

  const shading = legendConfig.bgEnabled
    ? { type: ShadingType.CLEAR, color: "auto", fill: legendConfig.bgColor }
    : undefined;

  if (coverConfig.modalities) {
    paragraphs.push(new Paragraph({
      spacing: { before: 100, after: 100 },
      shading,
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

  if (legendConfig.enabled) {
    const categories = [
      { symbol: "■", label: "Notions et définitions", desc: "vocabulaire juridique et notions clés à maîtriser", style: legendConfig.notions },
      { symbol: "■", label: "Articles, lois, textes normatifs", desc: "références légales et réglementaires", style: highlightConfig.articlesLois },
      { symbol: "■", label: "Jurisprudence, arrêts et décisions", desc: "solutions dégagées par les juridictions", style: highlightConfig.jurisprudence },
      { symbol: "■", label: "Exceptions et nuances", desc: "tempéraments et nuances apportés au principe général", style: legendConfig.exceptions }
    ];

    paragraphs.push(new Paragraph({
      spacing: { before: coverConfig.modalities ? 0 : 40, after: 100 },
      shading,
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

  paragraphs.push(new Paragraph({ spacing: { before: 0, after: 300 }, children: [] }));
  return paragraphs;
}

function buildParagraph(type, content, styles, highlightConfig, fontFamily) {
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
   6. INITIALISATION DE L'ÉCOUTEUR DE GÉNÉRATION
   ============================================================ */

export function initDocxBuilder() {
  const btnGenerate = document.getElementById("btn-generate");
  if (!btnGenerate) return;

  btnGenerate.addEventListener("click", async () => {
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
      sousSousSection: getStyleConfig("sousSousSection"),
      detail: getStyleConfig("detail"),
      citation: getStyleConfig("citation"),
      text: getStyleConfig("text"),
      bullet: getStyleConfig("bullet"),
      arrow: getStyleConfig("arrow")
    };
    
    const grandTitreConfig = getGrandTitreConfig();
    const highlightConfig = getHighlightConfig();
    const legendConfig = getLegendConfig();
    const coverConfig = getCoverConfig(); 

    const documentChildren = [];

    documentChildren.push(...buildLegendParagraphs(coverConfig, legendConfig, highlightConfig, fontFamily));

    const rawLines = rawText.split("\n");
    let i = 0;
    let lastWasBox = false; 

    while (i < rawLines.length) {
      const line = rawLines[i].trim();
      i++;
      if (!line) continue;

      const cls = classifyLine(line);
      if (cls.type === "skip") continue;

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
          documentChildren.push(new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }));
        }
        
        documentChildren.push(buildGrandTitreParagraph(cls.content, secondLine, grandTitreConfig, fontFamily));
        lastWasBox = true;
        continue;
      }

      if (cls.type === "grandTitre-partie") {
        if (lastWasBox) {
          documentChildren.push(new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }));
        }
        
        documentChildren.push(buildGrandTitreParagraph(cls.content, null, grandTitreConfig, fontFamily));
        lastWasBox = true;
        continue;
      }

      documentChildren.push(buildParagraph(cls.type, cls.content, styles, highlightConfig, fontFamily));
      lastWasBox = false; 
    }

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
          },
          children: documentChildren
        }
      ]
    });

    const blob = await Packer.toBlob(doc);
    triggerFileDownload(blob, "Cours_Formate.docx");
  });
}