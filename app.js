const state = {
  zip: null,
  fileName: "",
  documentXml: null,
  footnotesXml: null,
  stylesXml: null,
  textNodes: new Map(),
  footnoteTextNodes: new Map(),
  footnoteMap: new Map(),
  paragraphNodes: [],
  paragraphStyleMap: new Map(),
  headingStyleIds: {
    h1: null,
    h2: null,
    h3: null,
  },
  nextTextId: 1,
  isDocx: false,
  autoSaveEnabled: false,
  openedFileHandle: null,
  autoSaveTimer: null,
  isAutoSaving: false,
  autoSavePending: false,
  footnotePointerMove: null,
  activeFootnoteId: null,
  isResizingFootnotes: false,
  isDraggingDocumentPosition: false,
  deferredInstallPrompt: null,
  savedEditorRange: null,
};

const els = {
  openFileButton: document.querySelector("#openFileButton"),
  fileInput: document.querySelector("#fileInput"),
  newDocButton: document.querySelector("#newDocButton"),
  saveButton: document.querySelector("#saveButton"),
  autoSaveButton: document.querySelector("#autoSaveButton"),
  installAppButton: document.querySelector("#installAppButton"),
  editor: document.querySelector("#editor"),
  editorShell: document.querySelector("#editorShell"),
  status: document.querySelector("#status"),
  documentName: document.querySelector("#documentName"),
  documentMeta: document.querySelector("#documentMeta"),
  documentStats: document.querySelector("#documentStats"),
  saveState: document.querySelector("#saveState"),
  zoomInput: document.querySelector("#zoomInput"),
  zoomOutput: document.querySelector("#zoomOutput"),
  footnoteZoomInput: document.querySelector("#footnoteZoomInput"),
  footnoteZoomOutput: document.querySelector("#footnoteZoomOutput"),
  footnoteZoomControl: document.querySelector("#footnoteZoomControl"),
  formatSelect: document.querySelector("#formatSelect"),
  insertFootnoteButton: document.querySelector("#insertFootnoteButton"),
  toggleFootnotesButton: document.querySelector("#toggleFootnotesButton"),
  footnotesPane: document.querySelector("#footnotesPane"),
  footnotesList: document.querySelector("#footnotesList"),
  activeFootnoteLabel: document.querySelector("#activeFootnoteLabel"),
  footnoteResizeHandle: document.querySelector("#footnoteResizeHandle"),
  documentPositionInput: document.querySelector("#documentPositionInput"),
  documentPositionOutput: document.querySelector("#documentPositionOutput"),
  headingNav: document.querySelector("#headingNav"),
  editorContextMenu: document.querySelector("#editorContextMenu"),
  contextInsertFootnoteButton: document.querySelector("#contextInsertFootnoteButton"),
};

const parser = new DOMParser();
const serializer = new XMLSerializer();
const wordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const localDraftKey = "hebrew-text-editor-draft-v1";
const footnoteTextStyleId = "FootnoteText";
const footnoteReferenceStyleId = "FootnoteReference";
const wordFontSize10pt = "20";
let zoomMode = "shared";

if (!window.showOpenFilePicker) {
  document.body.classList.add("no-file-picker");
}

function setStatus(message, tone = "ready") {
  els.status.textContent = message;
  els.saveState.textContent = message;
  els.saveState.dataset.tone = tone;
}

function qName(node, localName) {
  return node.getElementsByTagNameNS(wordNs, localName);
}

function attr(node, localName) {
  return node.getAttributeNS(wordNs, localName) || node.getAttribute(`w:${localName}`);
}

function escapeHtml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function resetDocxState() {
  clearTimeout(state.autoSaveTimer);
  state.zip = null;
  state.fileName = "";
  state.documentXml = null;
  state.footnotesXml = null;
  state.stylesXml = null;
  state.textNodes.clear();
  state.footnoteTextNodes.clear();
  state.footnoteMap.clear();
  state.paragraphNodes = [];
  state.paragraphStyleMap.clear();
  state.headingStyleIds = {
    h1: null,
    h2: null,
    h3: null,
  };
  state.nextTextId = 1;
  state.isDocx = false;
  state.autoSaveEnabled = false;
  state.openedFileHandle = null;
  state.autoSaveTimer = null;
  state.isAutoSaving = false;
  state.autoSavePending = false;
  state.footnotePointerMove = null;
  state.activeFootnoteId = null;
  state.isResizingFootnotes = false;
  state.isDraggingDocumentPosition = false;
  state.savedEditorRange = null;
  els.documentPositionInput.value = "0";
  els.documentPositionOutput.textContent = "0%";
  els.headingNav.innerHTML = `<p class="nav-empty">אין כותרות במסמך</p>`;
  els.footnotesList.innerHTML = "";
  els.activeFootnoteLabel.textContent = "אין הערות במסמך";
  els.footnotesPane.hidden = false;
  els.toggleFootnotesButton.classList.add("active");
  els.toggleFootnotesButton.setAttribute("aria-pressed", "true");
  els.toggleFootnotesButton.textContent = "הסתר חלונית";
  if (els.autoSaveButton) {
    els.autoSaveButton.disabled = true;
    els.autoSaveButton.classList.remove("active");
    els.autoSaveButton.setAttribute("aria-pressed", "false");
    els.autoSaveButton.textContent = "שמירה אוטומטית";
  }
}

function parseXml(xmlText) {
  const xml = parser.parseFromString(xmlText, "application/xml");
  const error = xml.querySelector("parsererror");
  if (error) {
    throw new Error("לא הצלחתי לקרוא את מבנה ה־XML של המסמך.");
  }
  return xml;
}

function getParagraphText(paragraph) {
  return Array.from(qName(paragraph, "t")).map((node) => node.textContent).join("");
}

function isWordToggleOn(node) {
  if (!node) return false;
  const value = attr(node, "val");
  return value !== "0" && value !== "false" && value !== "off";
}

function isBoldRun(run) {
  return Array.from(qName(run, "b")).some(isWordToggleOn) || Array.from(qName(run, "bCs")).some(isWordToggleOn);
}

function runStyleDeclaration(run) {
  const properties = qName(run, "rPr")[0];
  if (!properties) return "";

  const styles = [];
  if (Array.from(qName(properties, "i")).some(isWordToggleOn) || Array.from(qName(properties, "iCs")).some(isWordToggleOn)) {
    styles.push("font-style:italic");
  }

  const underline = qName(properties, "u")[0];
  const underlineValue = underline ? attr(underline, "val") : "";
  if (underline && underlineValue !== "none") styles.push("text-decoration:underline");

  const color = qName(properties, "color")[0];
  const colorValue = color ? attr(color, "val") : "";
  if (colorValue && colorValue !== "auto") styles.push(`color:#${colorValue}`);

  const highlight = qName(properties, "highlight")[0];
  const highlightValue = highlight ? attr(highlight, "val") : "";
  const highlightColors = {
    yellow: "#fff59d",
    green: "#c8e6c9",
    cyan: "#b2ebf2",
    magenta: "#f8bbd0",
    blue: "#bbdefb",
    red: "#ffcdd2",
    darkBlue: "#90caf9",
    darkCyan: "#80deea",
    darkGreen: "#a5d6a7",
    darkMagenta: "#ce93d8",
    darkRed: "#ef9a9a",
    darkYellow: "#ffe082",
    darkGray: "#bdbdbd",
    lightGray: "#eeeeee",
    black: "#212121",
  };
  if (highlightColors[highlightValue]) styles.push(`background-color:${highlightColors[highlightValue]}`);

  const size = qName(properties, "sz")[0] || qName(properties, "szCs")[0];
  const sizeValue = size ? Number(attr(size, "val")) : 0;
  if (sizeValue > 0) styles.push(`font-size:${sizeValue / 2}pt`);

  const fonts = qName(properties, "rFonts")[0];
  const fontValue = fonts ? (attr(fonts, "cs") || attr(fonts, "ascii") || attr(fonts, "hAnsi")) : "";
  if (fontValue) styles.push(`font-family:${JSON.stringify(fontValue)}`);

  return styles.join(";");
}

function normalizeStyleText(value) {
  return (value || "").replace(/\s+/g, "").toLowerCase();
}

function headingTagFromStyle(styleId, styleName, outlineLevel) {
  const combined = `${styleId || ""} ${styleName || ""}`;
  const compact = normalizeStyleText(combined);

  if (outlineLevel === "0" || /heading1|headline1/i.test(combined) || compact.includes("כותרת1") || compact.includes("כותרתא") || compact.includes("title")) return "h1";
  if (outlineLevel === "1" || /heading2/i.test(combined) || compact.includes("כותרת2") || compact.includes("כותרתב")) return "h2";
  if (outlineLevel === "2" || /heading3/i.test(combined) || compact.includes("כותרת3") || compact.includes("כותרתג")) return "h3";
  return "p";
}

function buildParagraphStyleMap() {
  state.paragraphStyleMap.clear();
  if (!state.stylesXml) return;

  const styles = Array.from(qName(state.stylesXml, "style"));
  styles.forEach((style) => {
    if (attr(style, "type") !== "paragraph") return;

    const styleId = attr(style, "styleId");
    if (!styleId) return;

    const nameNode = qName(style, "name")[0];
    const name = nameNode ? attr(nameNode, "val") : styleId;
    const outlineNode = qName(style, "outlineLvl")[0];
    const outlineLevel = outlineNode ? attr(outlineNode, "val") : "";
    const tag = headingTagFromStyle(styleId, name, outlineLevel);

    state.paragraphStyleMap.set(styleId, {
      id: styleId,
      name,
      outlineLevel,
      tag,
    });

    if (tag === "h1" && !state.headingStyleIds.h1) state.headingStyleIds.h1 = styleId;
    if (tag === "h2" && !state.headingStyleIds.h2) state.headingStyleIds.h2 = styleId;
    if (tag === "h3" && !state.headingStyleIds.h3) state.headingStyleIds.h3 = styleId;

    if (tag === "h1" && /heading1|כותרת\s*1/i.test(`${styleId} ${name}`)) state.headingStyleIds.h1 = styleId;
    if (tag === "h2" && /heading2|כותרת\s*2/i.test(`${styleId} ${name}`)) state.headingStyleIds.h2 = styleId;
    if (tag === "h3" && /heading3|כותרת\s*3/i.test(`${styleId} ${name}`)) state.headingStyleIds.h3 = styleId;
  });
}

function paragraphStyleInfo(paragraph) {
  const pStyle = qName(paragraph, "pStyle")[0];
  const value = pStyle ? attr(pStyle, "val") : "";
  const mappedStyle = value ? state.paragraphStyleMap.get(value) : null;
  if (mappedStyle) return mappedStyle;

  const tag = headingTagFromStyle(value, value, "");
  return {
    id: value,
    name: value || "Normal",
    outlineLevel: "",
    tag,
  };
}

function paragraphAlignment(paragraph) {
  const jc = qName(paragraph, "jc")[0];
  return jc ? attr(jc, "val") : "";
}

function buildFootnoteMap() {
  state.footnoteMap.clear();
  if (!state.footnotesXml) return;

  const notes = Array.from(qName(state.footnotesXml, "footnote"));
  notes.forEach((note) => {
    const id = attr(note, "id");
    if (!id || Number(id) < 1) return;

    const pieces = [];
    Array.from(qName(note, "t")).forEach((textNode) => {
      if (!pieces.length && /^\s+$/.test(textNode.textContent)) return;
      const textId = `fn-${id}-${state.nextTextId++}`;
      state.footnoteTextNodes.set(textId, textNode);
      pieces.push(`<span class="docx-footnote-text" data-footnote-text-id="${textId}">${escapeHtml(textNode.textContent)}</span>`);
    });

    state.footnoteMap.set(id, pieces.join(""));
  });
}

function renderFootnoteCard(id) {
  const body = state.footnoteMap.get(id);
  if (!body) return "";
  return `
    <article class="footnote-card" data-footnote-id="${escapeHtml(id)}" contenteditable="false">
      <div class="footnote-number">${escapeHtml(id)}</div>
      <div class="footnote-body" contenteditable="true">${body}</div>
      <button class="footnote-delete" type="button" data-delete-footnote="${escapeHtml(id)}" aria-label="מחק הערת שוליים ${escapeHtml(id)}">מחיקה</button>
    </article>
  `;
}

function renderFootnotesPane() {
  if (!state.footnoteMap.size) {
    els.footnotesList.innerHTML = `<p class="footnotes-empty">אין הערות שוליים במסמך הזה.</p>`;
    els.activeFootnoteLabel.textContent = "אין הערות במסמך";
    return;
  }

  const html = Array.from(state.footnoteMap.keys())
    .sort((a, b) => Number(a) - Number(b))
    .map(renderFootnoteCard)
    .join("");

  els.footnotesList.innerHTML = html;
  els.activeFootnoteLabel.textContent = `${state.footnoteMap.size} הערות`;
}

function createWordElement(name, doc = state.documentXml) {
  return doc.createElementNS(wordNs, `w:${name}`);
}

function setWordAttr(node, name, value) {
  node.setAttributeNS(wordNs, `w:${name}`, value);
}

function ensureRunProperties(run) {
  let properties = qName(run, "rPr")[0];
  if (properties) return properties;

  const doc = run.ownerDocument || state.documentXml;
  properties = doc.createElementNS(wordNs, "w:rPr");
  run.insertBefore(properties, run.firstChild);
  return properties;
}

function removeWordChildren(parent, localName) {
  Array.from(qName(parent, localName)).forEach((node) => node.remove());
}

function setRunBold(run, shouldBold) {
  if (!run) return;

  const properties = ensureRunProperties(run);
  removeWordChildren(properties, "b");
  removeWordChildren(properties, "bCs");

  if (!shouldBold) return;

  const doc = run.ownerDocument || state.documentXml;
  properties.append(doc.createElementNS(wordNs, "w:b"));
  properties.append(doc.createElementNS(wordNs, "w:bCs"));
}

function ensureRunStyle(run, styleId) {
  const properties = ensureRunProperties(run);
  let style = qName(properties, "rStyle")[0];
  if (!style) {
    const doc = run.ownerDocument || state.documentXml;
    style = doc.createElementNS(wordNs, "w:rStyle");
    properties.insertBefore(style, properties.firstChild);
  }
  setWordAttr(style, "val", styleId);
}

function setRunFontSize(run, size = wordFontSize10pt) {
  const properties = ensureRunProperties(run);
  removeWordChildren(properties, "sz");
  removeWordChildren(properties, "szCs");

  const doc = run.ownerDocument || state.documentXml;
  const sz = doc.createElementNS(wordNs, "w:sz");
  const szCs = doc.createElementNS(wordNs, "w:szCs");
  setWordAttr(sz, "val", size);
  setWordAttr(szCs, "val", size);
  properties.append(sz, szCs);
}

function setRunRtl(run, isRtl = true) {
  const properties = ensureRunProperties(run);
  removeWordChildren(properties, "rtl");
  if (!isRtl) return;

  const doc = run.ownerDocument || state.documentXml;
  properties.append(doc.createElementNS(wordNs, "w:rtl"));
}

function setRunSuperscript(run) {
  const properties = ensureRunProperties(run);
  removeWordChildren(properties, "vertAlign");

  const doc = run.ownerDocument || state.documentXml;
  const vertAlign = doc.createElementNS(wordNs, "w:vertAlign");
  setWordAttr(vertAlign, "val", "superscript");
  properties.append(vertAlign);
}

function setParagraphStyle(paragraph, styleId) {
  const properties = ensureParagraphProperties(paragraph);
  let pStyle = qName(properties, "pStyle")[0];

  if (!pStyle) {
    const doc = paragraph.ownerDocument || state.documentXml;
    pStyle = doc.createElementNS(wordNs, "w:pStyle");
    properties.insertBefore(pStyle, properties.firstChild);
  }

  setWordAttr(pStyle, "val", styleId);
}

function setParagraphRtl(paragraph, isRtl = true) {
  const properties = ensureParagraphProperties(paragraph);
  removeWordChildren(properties, "bidi");
  if (!isRtl) return;

  const doc = paragraph.ownerDocument || state.documentXml;
  properties.append(doc.createElementNS(wordNs, "w:bidi"));
}

function setFootnoteParagraphDefaults(paragraph) {
  setParagraphStyle(paragraph, footnoteTextStyleId);
  setParagraphRtl(paragraph, true);
  setParagraphAlignment(paragraph, "right");
}

function setFootnoteTextRunDefaults(run) {
  setRunFontSize(run, wordFontSize10pt);
  setRunRtl(run, true);
}

function setFootnoteReferenceRunDefaults(run) {
  ensureRunStyle(run, footnoteReferenceStyleId);
  setRunSuperscript(run);
  setRunRtl(run, true);
}

function createFootnoteReferenceRun(id) {
  const run = createWordElement("r");
  setFootnoteReferenceRunDefaults(run);

  const reference = createWordElement("footnoteReference");
  setWordAttr(reference, "id", id);
  run.append(reference);
  return run;
}

function createFootnoteXmlNode(id, textNode) {
  const doc = state.footnotesXml;
  const footnote = createWordElement("footnote", doc);
  setWordAttr(footnote, "id", id);

  const paragraph = createWordElement("p", doc);
  setFootnoteParagraphDefaults(paragraph);

  const referenceRun = createWordElement("r", doc);
  setFootnoteReferenceRunDefaults(referenceRun);
  referenceRun.append(createWordElement("footnoteRef", doc));

  const spaceRun = createWordElement("r", doc);
  setFootnoteTextRunDefaults(spaceRun);
  const space = createWordElement("t", doc);
  space.setAttribute("xml:space", "preserve");
  space.textContent = " ";
  spaceRun.append(space);

  const textRun = createWordElement("r", doc);
  setFootnoteTextRunDefaults(textRun);
  const text = createWordElement("t", doc);
  text.textContent = textNode;
  textRun.append(text);
  paragraph.append(referenceRun, spaceRun, textRun);
  footnote.append(paragraph);
  return { footnote, text };
}

function ensureFootnotesXml() {
  if (state.footnotesXml) return;

  state.footnotesXml = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:footnotes xmlns:w="${wordNs}">
      <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
      <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
    </w:footnotes>`);
}

function nextFootnoteId() {
  const ids = Array.from(state.footnoteMap.keys()).map(Number).filter((id) => Number.isFinite(id));
  return String(ids.length ? Math.max(...ids) + 1 : 1);
}

function saveEditorSelection() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return;

  const range = selection.getRangeAt(0);
  const container = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
    ? range.commonAncestorContainer
    : range.commonAncestorContainer.parentElement;
  if (container && els.editor.contains(container)) {
    state.savedEditorRange = range.cloneRange();
  }
}

function restoreEditorSelection() {
  if (!state.savedEditorRange) return false;

  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(state.savedEditorRange);
  return true;
}

function textOffsetInRun(range, runSpan) {
  const walker = document.createTreeWalker(runSpan, NodeFilter.SHOW_TEXT);
  let offset = 0;
  let node = walker.nextNode();

  while (node) {
    if (node === range.startContainer) return offset + range.startOffset;
    offset += node.textContent.length;
    node = walker.nextNode();
  }

  return runSpan.textContent.length;
}

function createFootnoteRefElement(id) {
  const ref = document.createElement("sup");
  ref.className = "footnote-ref";
  ref.contentEditable = "false";
  ref.draggable = false;
  ref.dataset.footnoteRef = id;
  ref.textContent = id;
  return ref;
}

function insertReferenceIntoDocxRun(range, runSpan, id) {
  const textId = runSpan.dataset.textId;
  const xmlTextNode = state.textNodes.get(textId);
  if (!xmlTextNode) return false;

  const run = xmlTextNode.parentNode;
  const paragraph = run.parentNode;
  const originalText = runSpan.textContent;
  const offset = Math.min(Math.max(textOffsetInRun(range, runSpan), 0), originalText.length);
  const before = originalText.slice(0, offset);
  const after = originalText.slice(offset);
  const referenceRun = createFootnoteReferenceRun(id);

  xmlTextNode.textContent = before;
  runSpan.textContent = before;

  if (after) {
    const afterRun = run.cloneNode(true);
    const afterText = qName(afterRun, "t")[0];
    afterText.textContent = after;
    const afterTextId = `t-${state.nextTextId++}`;
    state.textNodes.set(afterTextId, afterText);

    paragraph.insertBefore(referenceRun, run.nextSibling);
    paragraph.insertBefore(afterRun, referenceRun.nextSibling);

    const afterSpan = runSpan.cloneNode(false);
    afterSpan.dataset.textId = afterTextId;
    afterSpan.textContent = after;
    runSpan.after(createFootnoteRefElement(id), afterSpan);
  } else {
    paragraph.insertBefore(referenceRun, run.nextSibling);
    runSpan.after(createFootnoteRefElement(id));
  }

  return true;
}

function appendReferenceToDocxParagraph(block, id) {
  const paragraphIndex = Number(block?.dataset.paragraphIndex);
  const paragraph = Number.isInteger(paragraphIndex) ? state.paragraphNodes[paragraphIndex] : null;
  if (!paragraph) return false;

  paragraph.append(createFootnoteReferenceRun(id));
  block.append(createFootnoteRefElement(id));
  return true;
}

function addFootnoteXml(id, defaultText) {
  ensureFootnotesXml();
  const root = qName(state.footnotesXml, "footnotes")[0] || state.footnotesXml.documentElement;
  const { footnote, text } = createFootnoteXmlNode(id, defaultText);
  root.append(footnote);

  const textId = `fn-${id}-${state.nextTextId++}`;
  state.footnoteTextNodes.set(textId, text);
  state.footnoteMap.set(id, `<span class="docx-footnote-text" data-footnote-text-id="${textId}">${escapeHtml(defaultText)}</span>`);
}

async function ensureDocxFootnotePackageParts() {
  if (!state.zip || !state.footnotesXml) return;

  const contentTypesEntry = state.zip.file("[Content_Types].xml");
  if (contentTypesEntry) {
    const text = await contentTypesEntry.async("text");
    if (!text.includes('PartName="/word/footnotes.xml"')) {
      const xml = parseXml(text);
      const override = xml.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types", "Override");
      override.setAttribute("PartName", "/word/footnotes.xml");
      override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml");
      xml.documentElement.append(override);
      state.zip.file("[Content_Types].xml", serializer.serializeToString(xml));
    }
  }

  const relsEntry = state.zip.file("word/_rels/document.xml.rels");
  if (relsEntry) {
    const text = await relsEntry.async("text");
    if (!text.includes("/footnotes") && !text.includes("footnotes.xml")) {
      const xml = parseXml(text);
      const relsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
      const existingIds = Array.from(xml.getElementsByTagNameNS(relsNs, "Relationship"))
        .map((node) => node.getAttribute("Id"))
        .filter(Boolean);
      let next = 1;
      while (existingIds.includes(`rId${next}`)) next += 1;
      const rel = xml.createElementNS(relsNs, "Relationship");
      rel.setAttribute("Id", `rId${next}`);
      rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes");
      rel.setAttribute("Target", "footnotes.xml");
      xml.documentElement.append(rel);
      state.zip.file("word/_rels/document.xml.rels", serializer.serializeToString(xml));
    }
  }
}

function focusFootnoteBody(id) {
  const body = els.footnotesList.querySelector(`[data-footnote-id="${CSS.escape(id)}"] .footnote-body`);
  if (!body) return;

  body.focus();
  const range = document.createRange();
  range.selectNodeContents(body);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

function insertFootnote() {
  restoreEditorSelection();
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) {
    setStatus("בחר מקום במסמך להוספת הערת שוליים", "error");
    return;
  }

  const range = selection.getRangeAt(0);
  const container = range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
    ? range.commonAncestorContainer
    : range.commonAncestorContainer.parentElement;
  if (!container || !els.editor.contains(container)) {
    setStatus("אפשר להוסיף הערה רק מתוך אזור העריכה", "error");
    return;
  }

  const id = nextFootnoteId();
  const defaultText = "הערה חדשה";
  let insertedReference = false;

  if (state.isDocx && state.documentXml) {
    addFootnoteXml(id, defaultText);
    const runSpan = container.closest?.(".docx-run");
    const block = container.closest?.(".docx-paragraph, p, h1, h2, h3");
    insertedReference = runSpan
      ? insertReferenceIntoDocxRun(range, runSpan, id)
      : appendReferenceToDocxParagraph(block, id);
  } else {
    state.footnoteMap.set(id, escapeHtml(defaultText));
    const ref = createFootnoteRefElement(id);
    range.collapse(false);
    range.insertNode(ref);
    insertedReference = true;
  }

  if (!insertedReference) {
    state.footnoteMap.delete(id);
    setStatus("לא הצלחתי למצוא פסקה להוספת ההערה", "error");
    return;
  }

  renderFootnotesPane();
  els.footnotesPane.hidden = false;
  els.toggleFootnotesButton.classList.add("active");
  els.toggleFootnotesButton.setAttribute("aria-pressed", "true");
  els.toggleFootnotesButton.textContent = "הסתר חלונית";
  setActiveFootnote(id);
  focusFootnoteBody(id);
  markDocumentDirty();
  setStatus("הערת שוליים נוספה", "dirty");
}

function removeFootnote(id) {
  if (!id || !state.footnoteMap.has(id)) return;

  els.editor.querySelectorAll(`.footnote-ref[data-footnote-ref="${CSS.escape(id)}"]`).forEach((ref) => ref.remove());

  if (state.documentXml) {
    Array.from(qName(state.documentXml, "footnoteReference"))
      .filter((reference) => attr(reference, "id") === id)
      .forEach((reference) => {
        const run = reference.parentNode;
        reference.remove();
        if (run?.localName === "r" && !Array.from(run.childNodes).some((node) =>
          ["t", "tab", "br", "footnoteReference"].includes(node.localName)
        )) run.remove();
      });
  }

  if (state.footnotesXml) {
    Array.from(qName(state.footnotesXml, "footnote"))
      .filter((note) => attr(note, "id") === id)
      .forEach((note) => {
        state.footnoteTextNodes.forEach((textNode, textId) => {
          if (note.contains(textNode)) state.footnoteTextNodes.delete(textId);
        });
        note.remove();
      });
  }

  state.footnoteMap.delete(id);
  if (state.activeFootnoteId === id) state.activeFootnoteId = null;
  renderFootnotesPane();
  syncFootnoteToVisibleParagraph();
  markDocumentDirty();
  setStatus(`הערת שוליים ${id} נמחקה`, "dirty");
}

function renderRun(run) {
  const chunks = [];
  const children = Array.from(run.childNodes);
  const isBold = isBoldRun(run);
  const style = runStyleDeclaration(run);
  const styleAttr = style ? ` style="${escapeHtml(style)}"` : "";
  children.forEach((child) => {
    if (child.localName === "t") {
      const textId = `t-${state.nextTextId++}`;
      state.textNodes.set(textId, child);
      chunks.push(
        `<span class="docx-run" data-text-id="${textId}" data-bold="${isBold}" data-original-bold="${isBold}"${styleAttr}>${escapeHtml(child.textContent)}</span>`
      );
    }

    if (child.localName === "tab") chunks.push("	");
    if (child.localName === "br") chunks.push("<br>");
    if (child.localName === "footnoteReference") {
      const id = attr(child, "id");
      chunks.push(`<sup class="footnote-ref" contenteditable="false" draggable="false" data-footnote-ref="${escapeHtml(id)}">${escapeHtml(id)}</sup>`);
    }
  });
  return chunks.join("");
}

function renderParagraph(paragraph, index) {
  const styleInfo = paragraphStyleInfo(paragraph);
  const tag = styleInfo.tag;
  const align = paragraphAlignment(paragraph);
  const runs = Array.from(qName(paragraph, "r"));
  const content = runs.map(renderRun).join("") || escapeHtml(getParagraphText(paragraph));
  const style = align === "center" ? "text-align:center" : align === "right" ? "text-align:right" : align === "both" ? "text-align:justify" : align === "left" ? "text-align:left" : "";
  const styleId = styleInfo.id ? ` data-word-style-id="${escapeHtml(styleInfo.id)}"` : "";
  const styleName = styleInfo.name ? ` data-word-style-name="${escapeHtml(styleInfo.name)}"` : "";
  const styleLevel = tag !== "p" ? ` data-style-level="${tag}"` : "";
  const wordAlign = align ? ` data-word-align="${escapeHtml(align)}"` : "";
  return `<${tag} class="docx-paragraph word-style-${tag}" data-paragraph-index="${index}"${styleId}${styleName}${styleLevel}${wordAlign} style="${style}">${content || "<br>"}</${tag}>`;
}

async function openDocx(file, fileHandle = null) {
  if (!window.JSZip) {
    throw new Error("ספריית JSZip עדיין לא נטענה. בדוק חיבור אינטרנט ורענן את הדף.");
  }

  resetDocxState();
  setStatus("קורא את קובץ Word...", "busy");

  const buffer = await file.arrayBuffer();
  state.zip = await JSZip.loadAsync(buffer);
  state.fileName = file.name;
  state.openedFileHandle = fileHandle;

  const documentEntry = state.zip.file("word/document.xml");
  if (!documentEntry) {
    throw new Error("זה לא נראה כמו קובץ DOCX תקין.");
  }

  state.documentXml = parseXml(await documentEntry.async("text"));
  const footnotesEntry = state.zip.file("word/footnotes.xml");
  state.footnotesXml = footnotesEntry ? parseXml(await footnotesEntry.async("text")) : null;
  const stylesEntry = state.zip.file("word/styles.xml");
  state.stylesXml = stylesEntry ? parseXml(await stylesEntry.async("text")) : null;
  buildParagraphStyleMap();
  buildFootnoteMap();

  const body = qName(state.documentXml, "body")[0];
  const paragraphs = Array.from(qName(body, "p"));
  state.paragraphNodes = paragraphs;
  const html = paragraphs.map(renderParagraph).join("");

  els.editor.innerHTML = html || `<p class="placeholder">המסמך נפתח, אבל לא נמצאו פסקאות טקסט לעריכה.</p>`;
  renderFootnotesPane();
  buildHeadingNav();
  updateDocumentPositionSlider();
  updateDocumentStats();
  setTimeout(syncFootnoteToVisibleParagraph, 0);
  els.documentName.textContent = file.name;
  els.documentMeta.textContent = `${paragraphs.length} פסקאות, ${state.footnoteMap.size} הערות שוליים`;
  els.saveButton.disabled = false;
  els.autoSaveButton.disabled = !state.openedFileHandle;
  state.isDocx = true;
  setStatus(state.openedFileHandle ? "המסמך פתוח לעריכה עם הרשאת שמירה" : "המסמך פתוח לעריכה ללא הרשאת שמירה לקובץ המקורי", "ready");
}

function setParagraphAlignment(paragraph, alignment) {
  if (!paragraph) return;
  if (!alignment) return;

  const properties = ensureParagraphProperties(paragraph);
  removeWordChildren(properties, "jc");

  const doc = paragraph.ownerDocument || state.documentXml;
  const jc = doc.createElementNS(wordNs, "w:jc");
  setWordAttr(jc, "val", alignment);
  properties.append(jc);
}

function paragraphAlignmentFromBlock(block) {
  const value = (block.style.textAlign || "").toLowerCase();
  if (value === "center") return "center";
  if (value === "left") return "left";
  if (value === "right") return "right";
  if (value === "justify" || value === "start") return "both";
  if (!value) return block.dataset.wordAlign || "right";
  return "";
}

function formatFromBlock(block) {
  if (!block) return "p";
  if (block.dataset.styleLevel) return block.dataset.styleLevel;
  if (block.matches("h1")) return "h1";
  if (block.matches("h2")) return "h2";
  if (block.matches("h3")) return "h3";
  return "p";
}

function syncParagraphFormattingFromEditor() {
  if (!state.documentXml) return;

  els.editor.querySelectorAll("[data-paragraph-index]").forEach((block) => {
    const paragraphIndex = Number(block.dataset.paragraphIndex);
    const paragraph = Number.isInteger(paragraphIndex) ? state.paragraphNodes[paragraphIndex] : null;
    if (!paragraph) return;

    const format = formatFromBlock(block);
    if (format !== "p" || !block.dataset.wordStyleId) {
      setParagraphWordStyle(paragraph, format);
    }
    setParagraphRtl(paragraph, true);
    setParagraphAlignment(paragraph, paragraphAlignmentFromBlock(block));
  });
}

function footnoteXmlById(id) {
  if (!state.footnotesXml) return null;
  return Array.from(qName(state.footnotesXml, "footnote")).find((note) => attr(note, "id") === id) || null;
}

function visibleFootnoteText(card) {
  const body = card.querySelector(".footnote-body");
  return body ? body.innerText.replace(/\n{2,}/g, "\n").trim() : "";
}

function syncFootnotesFromEditor() {
  if (!state.footnotesXml) return;

  state.footnoteTextNodes.forEach((xmlNode, id) => {
    const htmlNode = els.footnotesList.querySelector(`[data-footnote-text-id="${CSS.escape(id)}"]`);
    if (htmlNode) xmlNode.textContent = htmlNode.textContent;
  });

  els.footnotesList.querySelectorAll(".footnote-card[data-footnote-id]").forEach((card) => {
    const note = footnoteXmlById(card.dataset.footnoteId);
    if (!note) return;

    const body = card.querySelector(".footnote-body");
    const textNodes = Array.from(qName(note, "t"));
    if (!textNodes.length) return;

    const visibleText = visibleFootnoteText(card);
    const trackedText = Array.from(body?.querySelectorAll("[data-footnote-text-id]") || [])
      .map((node) => node.textContent)
      .join("")
      .trim();
    if (trackedText === visibleText) return;

    textNodes[0].textContent = visibleText;
    setTextSpacePreserve(textNodes[0], visibleText);
    textNodes.slice(1).forEach((node) => {
      node.textContent = "";
    });
  });
}

function isRealFootnote(note) {
  const id = Number(attr(note, "id"));
  return Number.isFinite(id) && id > 0;
}

function runHasChild(run, localName) {
  return Array.from(run.childNodes).some((child) => child.localName === localName);
}

function createFootnoteSpaceRun(doc) {
  const run = createWordElement("r", doc);
  setFootnoteTextRunDefaults(run);
  const text = createWordElement("t", doc);
  text.setAttribute("xml:space", "preserve");
  text.textContent = " ";
  run.append(text);
  return run;
}

function runText(run) {
  return Array.from(qName(run, "t")).map((node) => node.textContent).join("");
}

function ensureSpaceAfterFootnoteNumber(paragraph, referenceRun) {
  let next = referenceRun.nextSibling;
  while (next && next.nodeType !== Node.ELEMENT_NODE) next = next.nextSibling;

  if (next?.localName === "r" && /^\s/.test(runText(next))) {
    setFootnoteTextRunDefaults(next);
    return;
  }

  paragraph.insertBefore(createFootnoteSpaceRun(paragraph.ownerDocument), referenceRun.nextSibling);
}

function normalizeFootnotesXmlForWord() {
  if (!state.footnotesXml) return;

  Array.from(qName(state.footnotesXml, "footnote")).forEach((note) => {
    if (!isRealFootnote(note)) return;

    Array.from(qName(note, "p")).forEach((paragraph) => {
      setFootnoteParagraphDefaults(paragraph);

      const runs = Array.from(qName(paragraph, "r"));
      const referenceRun = runs.find((run) => runHasChild(run, "footnoteRef"));

      runs.forEach((run) => {
        if (runHasChild(run, "footnoteRef")) {
          setFootnoteReferenceRunDefaults(run);
          return;
        }
        setFootnoteTextRunDefaults(run);
      });

      if (referenceRun) ensureSpaceAfterFootnoteNumber(paragraph, referenceRun);
    });
  });
}

function normalizeDocumentFootnoteReferences() {
  if (!state.documentXml) return;

  Array.from(qName(state.documentXml, "r")).forEach((run) => {
    if (runHasChild(run, "footnoteReference")) {
      setFootnoteReferenceRunDefaults(run);
    }
  });
}

function styleById(styleId) {
  if (!state.stylesXml) return null;
  return Array.from(qName(state.stylesXml, "style")).find((style) => attr(style, "styleId") === styleId) || null;
}

function ensureStyleChild(style, localName) {
  let child = qName(style, localName)[0];
  if (child) return child;

  child = state.stylesXml.createElementNS(wordNs, `w:${localName}`);
  style.append(child);
  return child;
}

function appendStyleMeta(style, localName, value = null) {
  if (qName(style, localName)[0]) return;

  const node = state.stylesXml.createElementNS(wordNs, `w:${localName}`);
  if (value !== null) setWordAttr(node, "val", value);
  style.append(node);
}

function createWordStyle(styleId, type, name) {
  const style = state.stylesXml.createElementNS(wordNs, "w:style");
  setWordAttr(style, "type", type);
  setWordAttr(style, "styleId", styleId);
  appendStyleMeta(style, "name", name);
  return style;
}

function ensureFootnoteTextStyle() {
  if (!state.stylesXml) return;

  let style = styleById(footnoteTextStyleId);
  if (!style) {
    style = createWordStyle(footnoteTextStyleId, "paragraph", "footnote text");
    appendStyleMeta(style, "basedOn", "Normal");
    appendStyleMeta(style, "uiPriority", "99");
    appendStyleMeta(style, "semiHidden");
    appendStyleMeta(style, "unhideWhenUsed");
    state.stylesXml.documentElement.append(style);
  }

  const pPr = ensureStyleChild(style, "pPr");
  removeWordChildren(pPr, "bidi");
  removeWordChildren(pPr, "jc");
  pPr.append(state.stylesXml.createElementNS(wordNs, "w:bidi"));
  const jc = state.stylesXml.createElementNS(wordNs, "w:jc");
  setWordAttr(jc, "val", "right");
  pPr.append(jc);

  const rPr = ensureStyleChild(style, "rPr");
  removeWordChildren(rPr, "sz");
  removeWordChildren(rPr, "szCs");
  removeWordChildren(rPr, "rtl");
  const sz = state.stylesXml.createElementNS(wordNs, "w:sz");
  const szCs = state.stylesXml.createElementNS(wordNs, "w:szCs");
  setWordAttr(sz, "val", wordFontSize10pt);
  setWordAttr(szCs, "val", wordFontSize10pt);
  rPr.append(sz, szCs, state.stylesXml.createElementNS(wordNs, "w:rtl"));
}

function ensureFootnoteReferenceStyle() {
  if (!state.stylesXml) return;

  let style = styleById(footnoteReferenceStyleId);
  if (!style) {
    style = createWordStyle(footnoteReferenceStyleId, "character", "footnote reference");
    appendStyleMeta(style, "basedOn", "DefaultParagraphFont");
    appendStyleMeta(style, "uiPriority", "99");
    appendStyleMeta(style, "semiHidden");
    appendStyleMeta(style, "unhideWhenUsed");
    state.stylesXml.documentElement.append(style);
  }

  const rPr = ensureStyleChild(style, "rPr");
  removeWordChildren(rPr, "vertAlign");
  const vertAlign = state.stylesXml.createElementNS(wordNs, "w:vertAlign");
  setWordAttr(vertAlign, "val", "superscript");
  rPr.append(vertAlign);
}

function ensureFootnoteStylesForWord() {
  ensureFootnoteTextStyle();
  ensureFootnoteReferenceStyle();
}

function updateXmlFromEditor() {
  if (state.isDocx) syncNativeBoldMarkupFromEditor();

  state.textNodes.forEach((xmlNode, id) => {
    const htmlNode = els.editor.querySelector(`[data-text-id="${CSS.escape(id)}"]`);
    if (htmlNode) {
      xmlNode.textContent = htmlNode.textContent;
      setTextSpacePreserve(xmlNode, htmlNode.textContent);
      if (htmlNode.dataset.boldDirty === "true" || htmlNode.dataset.bold !== htmlNode.dataset.originalBold) {
        setRunBold(xmlNode.parentNode, htmlNode.dataset.bold === "true");
        htmlNode.dataset.originalBold = htmlNode.dataset.bold;
        delete htmlNode.dataset.boldDirty;
      }
    } else if (xmlNode.isConnected) {
      const run = xmlNode.parentNode;
      xmlNode.remove();
      if (run?.localName === "r" && !qName(run, "t").length && !qName(run, "footnoteReference").length) run.remove();
    }
  });

  syncParagraphFormattingFromEditor();
  syncFootnotesFromEditor();
  normalizeDocumentFootnoteReferences();
  normalizeFootnotesXmlForWord();
  ensureFootnoteStylesForWord();
}

function editorParagraphBlocks() {
  return Array.from(els.editor.children).filter((node) =>
    node.matches?.(".docx-paragraph, p, h1, h2, h3, div")
  );
}

function paragraphAncestor(node) {
  let current = node;
  while (current && current !== state.documentXml) {
    if (current.localName === "p") return current;
    current = current.parentNode;
  }
  return null;
}

function replaceWordParagraphText(paragraph, text) {
  Array.from(paragraph.childNodes).forEach((child) => {
    if (child.localName !== "pPr") child.remove();
  });
  const run = createWordElement("r", paragraph.ownerDocument);
  setRunRtl(run, true);
  const textNode = createWordElement("t", paragraph.ownerDocument);
  textNode.textContent = text;
  setTextSpacePreserve(textNode, text);
  run.append(textNode);
  paragraph.append(run);
}

function createWordParagraphFromBlock(block, parent) {
  const doc = parent.ownerDocument || state.documentXml;
  const paragraph = createWordElement("p", doc);
  setParagraphRtl(paragraph, true);
  setParagraphAlignment(paragraph, paragraphAlignmentFromBlock(block));
  const format = formatFromBlock(block);
  if (format !== "p") setParagraphWordStyle(paragraph, format);
  replaceWordParagraphText(paragraph, block.innerText.replace(/\n+/g, " "));
  return paragraph;
}

function syncParagraphStructureFromEditor() {
  if (!state.isDocx || !state.documentXml) return;

  const blocks = editorParagraphBlocks();
  const originalParagraphs = [...state.paragraphNodes];
  const seenParagraphIndexes = new Set();
  blocks.forEach((block) => {
    const index = block.dataset.paragraphIndex;
    if (index === undefined) return;
    if (seenParagraphIndexes.has(index)) {
      delete block.dataset.paragraphIndex;
      block.dataset.newParagraph = "true";
      return;
    }
    seenParagraphIndexes.add(index);
  });
  const mappedBlocks = blocks.filter((block) => block.dataset.paragraphIndex !== undefined);
  const mappedParagraphs = mappedBlocks.map((block) => originalParagraphs[Number(block.dataset.paragraphIndex)]);

  if (mappedParagraphs.some((paragraph) => !paragraph)) {
    throw new Error("מיפוי הפסקאות השתבש. יש לפתוח מחדש את המסמך לפני השמירה.");
  }
  const mappedOriginalOrder = mappedParagraphs.map((paragraph) => originalParagraphs.indexOf(paragraph));
  if (mappedOriginalOrder.some((value, index) => index && value < mappedOriginalOrder[index - 1])) {
    throw new Error("שינוי סדר פסקאות עדיין אינו נתמך, ולכן השמירה נעצרה.");
  }

  originalParagraphs.forEach((paragraph) => {
    if (!mappedParagraphs.includes(paragraph)) paragraph.remove();
  });

  mappedBlocks.forEach((block) => {
    const paragraph = originalParagraphs[Number(block.dataset.paragraphIndex)];
    const trackedRuns = Array.from(block.querySelectorAll("[data-text-id]"));
    const foreignRun = trackedRuns.some((span) => {
      const xmlText = state.textNodes.get(span.dataset.textId);
      return xmlText && paragraphAncestor(xmlText) !== paragraph;
    });
    if (!trackedRuns.length || foreignRun) {
      replaceWordParagraphText(paragraph, block.innerText.replace(/\n+/g, " "));
    }
  });

  blocks.forEach((block, blockIndex) => {
    if (block.dataset.paragraphIndex !== undefined) return;

    const previousBlock = blocks.slice(0, blockIndex).reverse().find((item) => item.dataset.paragraphIndex !== undefined);
    const nextBlock = blocks.slice(blockIndex + 1).find((item) => item.dataset.paragraphIndex !== undefined);
    const previousParagraph = previousBlock ? originalParagraphs[Number(previousBlock.dataset.paragraphIndex)] : null;
    const nextParagraph = nextBlock ? originalParagraphs[Number(nextBlock.dataset.paragraphIndex)] : null;
    const parent = previousParagraph?.parentNode || nextParagraph?.parentNode;
    if (!parent) throw new Error("לא נמצא מקום בטוח להוספת הפסקה החדשה.");
    if (previousParagraph && nextParagraph && previousParagraph.parentNode !== nextParagraph.parentNode) {
      throw new Error("לא ניתן להוסיף פסקה בגבול שבין שני מבנים שונים במסמך Word.");
    }

    const paragraph = createWordParagraphFromBlock(block, parent);
    if (nextParagraph?.parentNode === parent) parent.insertBefore(paragraph, nextParagraph);
    else parent.insertBefore(paragraph, previousParagraph?.nextSibling || null);
    originalParagraphs.push(paragraph);
    block.dataset.newParagraph = "true";
    block._wordParagraph = paragraph;
  });

  const body = qName(state.documentXml, "body")[0];
  state.paragraphNodes = Array.from(qName(body, "p"));
  blocks.forEach((block) => {
    const paragraph = block._wordParagraph || originalParagraphs[Number(block.dataset.paragraphIndex)];
    const newIndex = state.paragraphNodes.indexOf(paragraph);
    if (newIndex < 0) throw new Error("לא ניתן היה להשלים את מיפוי הפסקה החדשה.");
    block.dataset.paragraphIndex = String(newIndex);
    block.classList.add("docx-paragraph");
    delete block._wordParagraph;
  });
}

function purgeOrphanedFootnotes() {
  if (!state.documentXml) return;
  const referencedIds = new Set(Array.from(qName(state.documentXml, "footnoteReference")).map((node) => attr(node, "id")));
  Array.from(state.footnoteMap.keys()).forEach((id) => {
    if (referencedIds.has(id)) return;
    if (state.footnotesXml) {
      Array.from(qName(state.footnotesXml, "footnote"))
        .filter((note) => attr(note, "id") === id)
        .forEach((note) => note.remove());
    }
    state.footnoteMap.delete(id);
    state.footnoteTextNodes.forEach((node, key) => {
      if (!node.isConnected) state.footnoteTextNodes.delete(key);
    });
  });
}

function syncFootnoteReferencePositions() {
  if (!state.documentXml) return;

  Array.from(qName(state.documentXml, "footnoteReference")).forEach((reference) => {
    const run = reference.parentNode;
    reference.remove();
    if (run?.localName === "r" && !qName(run, "t").length) run.remove();
  });

  editorParagraphBlocks().forEach((block) => {
    const paragraph = state.paragraphNodes[Number(block.dataset.paragraphIndex)];
    if (!paragraph) return;
    Array.from(block.querySelectorAll(".footnote-ref[data-footnote-ref]")).forEach((ref) => {
      const id = ref.dataset.footnoteRef;
      const referenceRun = createFootnoteReferenceRun(id);
      const tracked = Array.from(block.querySelectorAll("[data-text-id]"));
      const previous = tracked.filter((span) => span.compareDocumentPosition(ref) & Node.DOCUMENT_POSITION_FOLLOWING).at(-1);
      const next = tracked.find((span) => span.compareDocumentPosition(ref) & Node.DOCUMENT_POSITION_PRECEDING);
      const previousRun = previous ? state.textNodes.get(previous.dataset.textId)?.parentNode : null;
      const nextRun = next ? state.textNodes.get(next.dataset.textId)?.parentNode : null;
      if (previousRun?.parentNode === paragraph) paragraph.insertBefore(referenceRun, previousRun.nextSibling);
      else if (nextRun?.parentNode === paragraph) paragraph.insertBefore(referenceRun, nextRun);
      else paragraph.append(referenceRun);
    });
  });
  purgeOrphanedFootnotes();
}

function ensureReferencedFootnotesExist() {
  if (!state.documentXml) return;
  const referencedIds = new Set(Array.from(qName(state.documentXml, "footnoteReference")).map((node) => attr(node, "id")));
  referencedIds.forEach((id) => {
    if (footnoteXmlById(id)) return;
    if (!state.footnoteMap.has(id)) {
      throw new Error(`הפניה להערת שוליים ${id} אינה מצביעה על הערה קיימת.`);
    }
    ensureFootnotesXml();
    const card = els.footnotesList.querySelector(`[data-footnote-id="${CSS.escape(id)}"]`);
    const textValue = card ? visibleFootnoteText(card) : "הערה חדשה";
    const root = qName(state.footnotesXml, "footnotes")[0] || state.footnotesXml.documentElement;
    const { footnote, text } = createFootnoteXmlNode(id, textValue);
    root.append(footnote);
    state.footnoteTextNodes.set(`fn-${id}-${state.nextTextId++}`, text);
  });
}

function renumberFootnotesByReferenceOrder() {
  if (!state.documentXml || !state.footnotesXml) return;
  const orderedIds = [];
  Array.from(qName(state.documentXml, "footnoteReference")).forEach((reference) => {
    const id = attr(reference, "id");
    if (!orderedIds.includes(id)) orderedIds.push(id);
  });
  const mapping = new Map(orderedIds.map((id, index) => [id, String(index + 1)]));
  if ([...mapping].every(([oldId, newId]) => oldId === newId)) return;

  Array.from(qName(state.documentXml, "footnoteReference")).forEach((reference) => {
    const newId = mapping.get(attr(reference, "id"));
    if (newId) setWordAttr(reference, "id", `tmp-${newId}`);
  });
  Array.from(qName(state.footnotesXml, "footnote")).forEach((note) => {
    const newId = mapping.get(attr(note, "id"));
    if (newId) setWordAttr(note, "id", `tmp-${newId}`);
  });
  Array.from(qName(state.documentXml, "footnoteReference")).forEach((reference) => {
    const temporary = attr(reference, "id");
    if (temporary?.startsWith("tmp-")) setWordAttr(reference, "id", temporary.slice(4));
  });
  Array.from(qName(state.footnotesXml, "footnote")).forEach((note) => {
    const temporary = attr(note, "id");
    if (temporary?.startsWith("tmp-")) setWordAttr(note, "id", temporary.slice(4));
  });

  const remappedNotes = new Map();
  orderedIds.forEach((oldId) => remappedNotes.set(mapping.get(oldId), state.footnoteMap.get(oldId)));
  state.footnoteMap = remappedNotes;
  state.activeFootnoteId = mapping.get(state.activeFootnoteId) || null;
  els.editor.querySelectorAll(".footnote-ref[data-footnote-ref]").forEach((ref) => {
    const newId = mapping.get(ref.dataset.footnoteRef);
    if (newId) {
      ref.dataset.footnoteRef = newId;
      ref.textContent = newId;
    }
  });
  renderFootnotesPane();
}

function assertXmlIsReadable(xmlText, partName) {
  const xml = parser.parseFromString(xmlText, "application/xml");
  if (xml.querySelector("parsererror")) {
    throw new Error(`קובץ ה־XML ${partName} אינו תקין לאחר השמירה.`);
  }
  return xml;
}

async function validateGeneratedDocx(blob) {
  let verificationZip;
  try {
    verificationZip = await JSZip.loadAsync(blob);
  } catch (error) {
    throw new Error("קובץ ה־DOCX שנוצר אינו אריזת ZIP תקינה.");
  }

  const requiredParts = ["[Content_Types].xml", "_rels/.rels", "word/document.xml"];
  const missingPart = requiredParts.find((path) => !verificationZip.file(path));
  if (missingPart) {
    throw new Error(`קובץ ה־DOCX שנוצר חסר את החלק החיוני ${missingPart}.`);
  }

  const documentText = await verificationZip.file("word/document.xml").async("text");
  const documentXml = assertXmlIsReadable(documentText, "word/document.xml");
  const documentBody = qName(documentXml, "body")[0];
  if (!documentBody) throw new Error("קובץ ה־DOCX שנוצר אינו מכיל גוף מסמך.");

  const references = Array.from(qName(documentXml, "footnoteReference"))
    .map((node) => attr(node, "id"))
    .filter((id) => Number(id) > 0);
  if (references.length) {
    const footnotesEntry = verificationZip.file("word/footnotes.xml");
    if (!footnotesEntry) {
      throw new Error("המסמך מכיל הפניות להערות שוליים, אך קובץ ההערות חסר.");
    }
    const footnotesText = await footnotesEntry.async("text");
    const footnotesXml = assertXmlIsReadable(footnotesText, "word/footnotes.xml");
    const noteIds = new Set(
      Array.from(qName(footnotesXml, "footnote")).map((node) => attr(node, "id"))
    );
    const missingNote = references.find((id) => !noteIds.has(id));
    if (missingNote) {
      throw new Error(`הפניה להערת שוליים ${missingNote} אינה מצביעה על הערה קיימת.`);
    }
  }

  return true;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function plainEditorText() {
  return els.editor.innerText.replace(/\n{3,}/g, "\n\n").trim();
}

function updateDocumentStats() {
  const text = plainEditorText();
  const words = text ? text.split(/\s+/).filter(Boolean).length : 0;
  const chars = text.length;
  els.documentStats.textContent = `${words} מילים · ${chars} תווים`;
}

function saveLocalDraft() {
  if (state.isDocx) return;

  const draft = {
    html: els.editor.innerHTML,
    footnotesHtml: els.footnotesList.innerHTML,
    savedAt: new Date().toISOString(),
  };
  localStorage.setItem(localDraftKey, JSON.stringify(draft));
}

function restoreLocalDraft() {
  const raw = localStorage.getItem(localDraftKey);
  if (!raw) {
    updateDocumentStats();
    buildHeadingNav();
    updateDocumentPositionSlider();
    return;
  }

  try {
    const draft = JSON.parse(raw);
    if (!draft?.html) return;

    els.editor.innerHTML = draft.html;
    if (draft.footnotesHtml) {
      els.footnotesList.innerHTML = draft.footnotesHtml;
      state.footnoteMap.clear();
      els.footnotesList.querySelectorAll(".footnote-card").forEach((card) => {
        state.footnoteMap.set(card.dataset.footnoteId, card.querySelector(".footnote-body")?.innerHTML || "");
      });
      els.activeFootnoteLabel.textContent = state.footnoteMap.size ? `${state.footnoteMap.size} הערות` : "אין הערות במסמך";
    }
    els.documentName.textContent = "טיוטה מקומית";
    els.documentMeta.textContent = "נשמרת אוטומטית בדפדפן";
    setStatus("טיוטה מקומית נטענה", "ready");
    updateDocumentStats();
    buildHeadingNav();
    updateDocumentPositionSlider();
  } catch (error) {
    console.warn("Could not restore local draft", error);
    localStorage.removeItem(localDraftKey);
    updateDocumentStats();
  }
}

async function createDocxBlob() {
  if (!state.zip || !state.documentXml) return;

  updateXmlFromEditor();
  syncParagraphStructureFromEditor();
  syncFootnoteReferencePositions();
  ensureReferencedFootnotesExist();
  renumberFootnotesByReferenceOrder();
  normalizeDocumentFootnoteReferences();
  normalizeFootnotesXmlForWord();
  await ensureDocxFootnotePackageParts();
  state.zip.file("word/document.xml", serializer.serializeToString(state.documentXml));
  if (state.footnotesXml) {
    state.zip.file("word/footnotes.xml", serializer.serializeToString(state.footnotesXml));
  }
  if (state.stylesXml) {
    state.zip.file("word/styles.xml", serializer.serializeToString(state.stylesXml));
  }

  const blob = await state.zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  await validateGeneratedDocx(blob);
  return blob;
}

async function saveDocx() {
  if (!state.zip || !state.documentXml) return;

  setStatus("שומר DOCX...", "busy");
  const blob = await createDocxBlob();

  const cleanName = state.fileName.replace(/\.docx$/i, "");
  downloadBlob(blob, `${cleanName}-edited.docx`);
  setStatus("המסמך נשמר כקובץ חדש", "ready");
}

async function writeAutoSave() {
  if (!state.autoSaveEnabled || !state.openedFileHandle) return;
  if (state.isAutoSaving) {
    state.autoSavePending = true;
    return;
  }

  state.isAutoSaving = true;
  state.autoSavePending = false;
  try {
    setStatus("שומר אוטומטית...", "busy");
    const blob = await createDocxBlob();
    let saved = false;
    for (let attempt = 0; attempt < 3 && !saved; attempt += 1) {
      try {
        await state.openedFileHandle.getFile();
        const writable = await state.openedFileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        saved = true;
      } catch (error) {
        if (error.name !== "InvalidStateError" || attempt === 2) throw error;
        await new Promise((resolve) => setTimeout(resolve, 150 * (attempt + 1)));
      }
    }
    setStatus("נשמר אוטומטית בקובץ המקורי", "ready");
  } catch (error) {
    console.error(error);
    state.autoSaveEnabled = false;
    els.autoSaveButton.classList.remove("active");
    els.autoSaveButton.setAttribute("aria-pressed", "false");
    els.autoSaveButton.textContent = "שמירה אוטומטית";
    setStatus("השמירה האוטומטית נכשלה וכובתה", "error");
  } finally {
    state.isAutoSaving = false;
    if (state.autoSaveEnabled && state.autoSavePending) {
      state.autoSavePending = false;
      clearTimeout(state.autoSaveTimer);
      state.autoSaveTimer = setTimeout(writeAutoSave, 250);
    }
  }
}

function scheduleAutoSave() {
  if (!state.autoSaveEnabled) return;
  if (state.isAutoSaving) state.autoSavePending = true;
  clearTimeout(state.autoSaveTimer);
  state.autoSaveTimer = setTimeout(() => {
    writeAutoSave();
  }, 1600);
}

async function toggleAutoSave() {
  if (!state.isDocx) {
    setStatus("שמירה אוטומטית זמינה אחרי פתיחת קובץ DOCX", "error");
    return;
  }

  if (state.autoSaveEnabled) {
    clearTimeout(state.autoSaveTimer);
    state.autoSaveEnabled = false;
    els.autoSaveButton.classList.remove("active");
    els.autoSaveButton.setAttribute("aria-pressed", "false");
    els.autoSaveButton.textContent = "שמירה אוטומטית";
    setStatus("שמירה אוטומטית כבויה", "ready");
    return;
  }

  if (!state.openedFileHandle) {
    setStatus("כדי לשמור אוטומטית לקובץ המקורי, פתח את הקובץ דרך כפתור פתיחת Word בדפדפן Chrome או Edge.", "error");
    return;
  }

  try {
    state.autoSaveEnabled = true;
    els.autoSaveButton.classList.add("active");
    els.autoSaveButton.setAttribute("aria-pressed", "true");
    els.autoSaveButton.textContent = "שמירה אוטומטית פעילה";
    setStatus("שמירה אוטומטית פעילה לקובץ המקורי", "ready");
    await writeAutoSave();
  } catch (error) {
    if (error.name !== "AbortError") {
      console.error(error);
      setStatus("לא הצלחתי להפעיל שמירה אוטומטית", "error");
    }
  }
}

function newDocument() {
  resetDocxState();
  localStorage.removeItem(localDraftKey);
  els.editor.innerHTML = `
    <h2>מסמך חדש</h2>
    <p></p>
  `;
  els.footnotesList.innerHTML = `<p class="footnotes-empty">אין הערות שוליים במסמך הזה.</p>`;
  els.activeFootnoteLabel.textContent = "אין הערות במסמך";
  buildHeadingNav();
  updateDocumentPositionSlider();
  els.documentName.textContent = "מסמך ללא שם";
  els.documentMeta.textContent = "תצוגה זורמת, מימין לשמאל";
  els.saveButton.disabled = true;
  updateDocumentStats();
  setStatus("מסמך חדש מוכן לעריכה", "ready");
}

function selectedDocxRuns() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return [];

  const range = selection.getRangeAt(0);
  if (range.collapsed) {
    const node = selection.anchorNode?.nodeType === Node.TEXT_NODE ? selection.anchorNode.parentElement : selection.anchorNode;
    const run = node?.closest?.(".docx-run");
    return run ? [run] : [];
  }

  return Array.from(els.editor.querySelectorAll(".docx-run")).filter((run) => range.intersectsNode(run));
}

function textLengthOfRunSpan(runSpan) {
  return runSpan.textContent.length;
}

function rangeOffsetsInRunSpan(range, runSpan) {
  if (!range.intersectsNode(runSpan)) return null;
  const runRange = document.createRange();
  runRange.selectNodeContents(runSpan);
  const length = textLengthOfRunSpan(runSpan);

  let start = 0;
  if (range.compareBoundaryPoints(Range.START_TO_START, runRange) > 0) {
    const before = document.createRange();
    before.selectNodeContents(runSpan);
    before.setEnd(range.startContainer, range.startOffset);
    start = Math.max(0, Math.min(length, before.toString().length));
    before.detach();
  }

  let end = length;
  if (range.compareBoundaryPoints(Range.END_TO_END, runRange) < 0) {
    const selected = document.createRange();
    selected.selectNodeContents(runSpan);
    selected.setStart(runSpan, 0);
    selected.setEnd(range.endContainer, range.endOffset);
    end = Math.max(start, Math.min(length, selected.toString().length));
    selected.detach();
  }

  runRange.detach();
  return { start, end };
}

function setTextSpacePreserve(textNode, text) {
  if (/^\s|\s$/.test(text)) {
    textNode.setAttribute("xml:space", "preserve");
  } else {
    textNode.removeAttribute("xml:space");
  }
}

function cloneRunWithText(sourceRun, text, shouldBold) {
  const doc = sourceRun.ownerDocument || state.documentXml;
  const run = doc.createElementNS(wordNs, "w:r");
  const properties = qName(sourceRun, "rPr")[0];
  if (properties) run.append(properties.cloneNode(true));

  const textNode = doc.createElementNS(wordNs, "w:t");
  textNode.textContent = text;
  setTextSpacePreserve(textNode, text);
  run.append(textNode);
  setRunBold(run, shouldBold);
  return { run, textNode };
}

function runSpanAttributes(sourceSpan, textId, shouldBold) {
  const attrs = [
    `class="docx-run"`,
    `data-text-id="${escapeHtml(textId)}"`,
    `data-bold="${shouldBold}"`,
    `data-original-bold="${shouldBold}"`,
  ];
  const style = sourceSpan.getAttribute("style");
  if (style) attrs.push(`style="${escapeHtml(style)}"`);
  return attrs.join(" ");
}

function replaceRunSpanWithParts(sourceSpan, parts) {
  const html = parts
    .map((part) => `<span ${runSpanAttributes(sourceSpan, part.textId, part.bold)}>${escapeHtml(part.text)}</span>`)
    .join("");
  sourceSpan.insertAdjacentHTML("beforebegin", html);
  sourceSpan.remove();
}

function explicitBoldValue(element) {
  if (!(element instanceof HTMLElement)) return null;
  if (element.matches("b, strong")) return true;

  const weight = element.style.fontWeight;
  if (!weight) return null;
  if (weight === "bold" || weight === "bolder") return true;
  if (weight === "normal" || weight === "lighter") return false;

  const numericWeight = Number(weight);
  if (!Number.isNaN(numericWeight)) return numericWeight >= 600;
  return null;
}

function textNodeBoldValue(textNode, runSpan) {
  let isBold = runSpan.dataset.bold === "true";
  const path = [];
  let node = textNode.parentElement;

  while (node && node !== runSpan) {
    path.unshift(node);
    node = node.parentElement;
  }

  path.forEach((element) => {
    const value = explicitBoldValue(element);
    if (value !== null) isBold = value;
  });

  return isBold;
}

function collectNativeBoldParts(runSpan) {
  const parts = [];
  const walker = document.createTreeWalker(runSpan, NodeFilter.SHOW_TEXT);
  let textNode = walker.nextNode();

  while (textNode) {
    const text = textNode.textContent;
    if (text) {
      const bold = textNodeBoldValue(textNode, runSpan);
      const previous = parts[parts.length - 1];
      if (previous && previous.bold === bold) {
        previous.text += text;
      } else {
        parts.push({ text, bold });
      }
    }
    textNode = walker.nextNode();
  }

  return parts;
}

function replaceDocxRunWithBoldParts(runSpan, parts) {
  const xmlTextNode = state.textNodes.get(runSpan.dataset.textId);
  const sourceRun = xmlTextNode?.parentNode;
  const parent = sourceRun?.parentNode;
  if (!sourceRun || !parent || !parts.length) return false;

  if (parts.length === 1) {
    runSpan.textContent = parts[0].text;
    runSpan.dataset.bold = String(parts[0].bold);
    runSpan.dataset.boldDirty = "true";
    setRunBold(sourceRun, parts[0].bold);
    return true;
  }

  const fragment = sourceRun.ownerDocument.createDocumentFragment();
  const htmlParts = parts.map((part) => {
    const { run, textNode } = cloneRunWithText(sourceRun, part.text, part.bold);
    const textId = `t-${state.nextTextId++}`;
    state.textNodes.set(textId, textNode);
    fragment.append(run);
    return { ...part, textId };
  });

  state.textNodes.delete(runSpan.dataset.textId);
  parent.insertBefore(fragment, sourceRun);
  sourceRun.remove();
  replaceRunSpanWithParts(runSpan, htmlParts);
  return true;
}

function syncNativeBoldMarkupFromEditor() {
  Array.from(els.editor.querySelectorAll(".docx-run")).forEach((runSpan) => {
    if (!runSpan.querySelector("b, strong, [style*='font-weight']")) return;

    const parts = collectNativeBoldParts(runSpan);
    if (!parts.length) return;

    const originalText = runSpan.textContent;
    const originalBold = runSpan.dataset.bold === "true";
    const changed = parts.length > 1 || parts[0].text !== originalText || parts[0].bold !== originalBold;
    if (changed) replaceDocxRunWithBoldParts(runSpan, parts);
  });
}

function splitDocxRunForBold(runSpan, start, end, shouldBold) {
  const text = runSpan.textContent;
  const xmlTextNode = state.textNodes.get(runSpan.dataset.textId);
  const sourceRun = xmlTextNode?.parentNode;
  const parent = sourceRun?.parentNode;
  if (!text || !sourceRun || !parent || start >= end) return false;

  const segments = [
    { text: text.slice(0, start), bold: runSpan.dataset.bold === "true" },
    { text: text.slice(start, end), bold: shouldBold },
    { text: text.slice(end), bold: runSpan.dataset.bold === "true" },
  ].filter((segment) => segment.text);

  const fragment = sourceRun.ownerDocument.createDocumentFragment();
  const parts = segments.map((segment) => {
    const { run, textNode } = cloneRunWithText(sourceRun, segment.text, segment.bold);
    const textId = `t-${state.nextTextId++}`;
    state.textNodes.set(textId, textNode);
    fragment.append(run);
    return { ...segment, textId };
  });

  state.textNodes.delete(runSpan.dataset.textId);
  parent.insertBefore(fragment, sourceRun);
  sourceRun.remove();
  replaceRunSpanWithParts(runSpan, parts);
  return true;
}

function toggleWordBoldSelection() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return false;

  const range = selection.getRangeAt(0);
  const runs = selectedDocxRuns();
  if (!runs.length) return false;

  const shouldBold = !runs.every((run) => run.dataset.bold === "true");
  runs.forEach((run) => {
    if (!range.collapsed) {
      const offsets = rangeOffsetsInRunSpan(range, run);
      const runLength = textLengthOfRunSpan(run);
      if (offsets && (offsets.start > 0 || offsets.end < runLength)) {
        splitDocxRunForBold(run, offsets.start, offsets.end, shouldBold);
        return;
      }
    }

    run.dataset.bold = String(shouldBold);
    run.dataset.boldDirty = "true";
    const xmlTextNode = state.textNodes.get(run.dataset.textId);
    setRunBold(xmlTextNode?.parentNode, shouldBold);
  });

  markDocumentDirty();
  setStatus(shouldBold ? "ההדגשה תישמר ב-Word" : "ההדגשה הוסרה ותישמר ב-Word", "dirty");
  return true;
}

function applyEditorAlignment(command) {
  const block = selectedEditorBlock();
  const alignment = {
    justifyRight: "right",
    justifyCenter: "center",
    justifyLeft: "left",
    justifyFull: "justify",
  }[command];

  document.execCommand(command, false, null);
  if (!alignment || !block) return;

  block.style.textAlign = alignment;
  block.dataset.wordAlign = paragraphAlignmentFromBlock(block);
  const paragraphIndex = Number(block.dataset.paragraphIndex);
  if (Number.isInteger(paragraphIndex)) {
    setParagraphAlignment(state.paragraphNodes[paragraphIndex], paragraphAlignmentFromBlock(block));
  }
}

function runCommand(command, value = null) {
  els.editor.focus();
  restoreEditorSelection();
  if (command === "bold" && state.isDocx && toggleWordBoldSelection()) return;

  if (command.startsWith("justify")) {
    applyEditorAlignment(command);
  } else {
    document.execCommand(command, false, value);
  }

  syncParagraphFormattingFromEditor();
  markDocumentDirty();
  setStatus("השינוי הוחל", "ready");
}

function selectedEditorBlock() {
  const selection = window.getSelection();
  if (!selection || !selection.anchorNode) return null;

  const node = selection.anchorNode.nodeType === Node.TEXT_NODE ? selection.anchorNode.parentElement : selection.anchorNode;
  return node?.closest?.(".docx-paragraph, p, h1, h2, h3") || null;
}

function fallbackStyleIdForFormat(format) {
  return {
    h1: "Heading1",
    h2: "Heading2",
    h3: "Heading3",
  }[format] || "";
}

function ensureParagraphProperties(paragraph) {
  let properties = qName(paragraph, "pPr")[0];
  if (properties) return properties;

  const doc = paragraph.ownerDocument || state.documentXml;
  properties = doc.createElementNS(wordNs, "w:pPr");
  paragraph.insertBefore(properties, paragraph.firstChild);
  return properties;
}

function setParagraphWordStyle(paragraph, format) {
  if (!paragraph || !state.documentXml) return;

  const properties = ensureParagraphProperties(paragraph);
  let pStyle = qName(properties, "pStyle")[0];

  if (format === "p") {
    pStyle?.remove();
    return;
  }

  if (!pStyle) {
    pStyle = state.documentXml.createElementNS(wordNs, "w:pStyle");
    properties.insertBefore(pStyle, properties.firstChild);
  }

  const styleId = state.headingStyleIds[format] || fallbackStyleIdForFormat(format);
  pStyle.setAttributeNS(wordNs, "w:val", styleId);
}

function updateBlockStyleMetadata(block, format) {
  if (!block) return;

  block.classList.remove("word-style-p", "word-style-h1", "word-style-h2", "word-style-h3");
  block.classList.add(`word-style-${format}`);

  if (format === "p") {
    block.removeAttribute("data-style-level");
    block.removeAttribute("data-word-style-id");
    block.removeAttribute("data-word-style-name");
    return;
  }

  const styleId = state.headingStyleIds[format] || fallbackStyleIdForFormat(format);
  const styleInfo = state.paragraphStyleMap.get(styleId);
  block.dataset.styleLevel = format;
  block.dataset.wordStyleId = styleId;
  block.dataset.wordStyleName = styleInfo?.name || styleId;
}

function applyParagraphFormat(format) {
  const beforeBlock = selectedEditorBlock();
  els.editor.focus();
  document.execCommand("formatBlock", false, format);
  const block = selectedEditorBlock() || beforeBlock;
  const paragraphIndex = Number(block?.dataset.paragraphIndex);

  updateBlockStyleMetadata(block, format);
  if (Number.isInteger(paragraphIndex)) {
    setParagraphWordStyle(state.paragraphNodes[paragraphIndex], format);
  }

  buildHeadingNav();
  markDocumentDirty();
}

function setActiveFootnote(id, shouldScroll = true) {
  if (!id || !state.footnoteMap.has(id) || state.activeFootnoteId === id) return;

  state.activeFootnoteId = id;
  els.footnotesList.querySelectorAll(".footnote-card").forEach((card) => {
    card.classList.toggle("active", card.dataset.footnoteId === id);
  });
  els.editor.querySelectorAll(".footnote-ref").forEach((ref) => {
    ref.classList.toggle("active", ref.dataset.footnoteRef === id);
  });
  els.activeFootnoteLabel.textContent = `הערה ${id}`;

  if (shouldScroll) {
    const card = els.footnotesList.querySelector(`[data-footnote-id="${CSS.escape(id)}"]`);
    card?.scrollIntoView({ block: "start", behavior: "smooth" });
  }
}

function syncFootnoteToVisibleParagraph() {
  if (!state.footnoteMap.size || els.footnotesPane.hidden) return;

  const editorBox = els.editor.getBoundingClientRect();
  const centerY = editorBox.top + editorBox.height * 0.42;
  const paragraphs = Array.from(els.editor.querySelectorAll(".docx-paragraph, p, h1, h2, h3"));
  let best = null;
  let bestDistance = Number.POSITIVE_INFINITY;

  paragraphs.forEach((paragraph) => {
    const ref = paragraph.querySelector(".footnote-ref");
    if (!ref) return;

    const box = paragraph.getBoundingClientRect();
    if (box.bottom < editorBox.top || box.top > editorBox.bottom) return;

    const distance = Math.abs(box.top + box.height / 2 - centerY);
    if (distance < bestDistance) {
      bestDistance = distance;
      best = ref.dataset.footnoteRef;
    }
  });

  if (best) setActiveFootnote(best);
}

function setFootnotesHeight(height) {
  const shellBox = els.editorShell.getBoundingClientRect();
  const min = 150;
  const max = Math.max(min, shellBox.height - 220);
  const nextHeight = Math.min(Math.max(height, min), max);
  els.editorShell.style.setProperty("--footnotes-height", `${nextHeight}px`);
}

function markDocumentDirty() {
  buildHeadingNav();
  updateDocumentPositionSlider();
  updateDocumentStats();
  saveLocalDraft();
  setStatus(state.isDocx ? "יש שינויים שלא נשמרו" : "נערך", "dirty");
  scheduleAutoSave();
}

function sharedFootnoteZoomValue(editorZoomValue) {
  return Math.max(78, Math.round(Number(editorZoomValue) * 0.86));
}

function applyEditorZoom() {
  const editorValue = Number(els.zoomInput.value);
  document.documentElement.style.setProperty("--editor-zoom", `${editorValue}%`);
  els.zoomOutput.textContent = `${editorValue}%`;

  if (zoomMode === "shared") {
    const footnoteValue = sharedFootnoteZoomValue(editorValue);
    els.footnoteZoomInput.value = String(footnoteValue);
    document.documentElement.style.setProperty("--footnote-zoom", `${footnoteValue}%`);
    els.footnoteZoomOutput.textContent = `${footnoteValue}%`;
  }
}

function applyFootnoteZoom() {
  const value = Number(els.footnoteZoomInput.value);
  document.documentElement.style.setProperty("--footnote-zoom", `${value}%`);
  els.footnoteZoomOutput.textContent = `${value}%`;
}

function setZoomMode(mode) {
  zoomMode = mode;
  document.querySelectorAll("[data-zoom-mode]").forEach((button) => {
    button.classList.toggle("active", button.dataset.zoomMode === mode);
  });

  els.footnoteZoomControl.hidden = mode === "shared";
  if (mode === "shared") {
    applyEditorZoom();
  } else {
    applyFootnoteZoom();
  }
}

function closeToolMenus(exceptMenu = null) {
  document.querySelectorAll(".tool-menu[open]").forEach((menu) => {
    if (menu !== exceptMenu) menu.removeAttribute("open");
  });
}

function headingTitle(heading) {
  return heading.textContent.replace(/\s+/g, " ").trim() || "כותרת ללא טקסט";
}

function headingLevel(heading) {
  if (heading.matches("h1, .word-style-h1")) return 1;
  if (heading.matches("h2, .word-style-h2")) return 2;
  return 3;
}

function buildHeadingNav() {
  const headings = Array.from(els.editor.querySelectorAll("h1, h2, h3, [data-style-level='h1'], [data-style-level='h2'], [data-style-level='h3']"));
  const uniqueHeadings = headings.filter((heading, index, all) => all.indexOf(heading) === index);

  if (!uniqueHeadings.length) {
    els.headingNav.innerHTML = `<p class="nav-empty">אין כותרות במסמך</p>`;
    return;
  }

  els.headingNav.innerHTML = uniqueHeadings
    .map((heading, index) => {
      heading.dataset.navIndex = String(index);
      const level = headingLevel(heading);
      return `<button class="heading-nav-item level-${level}" type="button" data-nav-index="${index}">${escapeHtml(headingTitle(heading))}</button>`;
    })
    .join("");
}

function scrollToHeading(index) {
  const heading = els.editor.querySelector(`[data-nav-index="${CSS.escape(String(index))}"]`);
  if (!heading) return;

  heading.scrollIntoView({ block: "start", behavior: "smooth" });
  heading.classList.add("nav-flash");
  setTimeout(() => heading.classList.remove("nav-flash"), 900);
}

function updateDocumentPositionSlider() {
  if (state.isDraggingDocumentPosition) return;

  const maxScroll = Math.max(0, els.editor.scrollHeight - els.editor.clientHeight);
  const percent = maxScroll ? Math.round((els.editor.scrollTop / maxScroll) * 100) : 0;
  els.documentPositionInput.value = String(percent);
  els.documentPositionOutput.textContent = `${percent}%`;

  const headings = Array.from(els.editor.querySelectorAll("[data-nav-index]"));
  let active = null;
  let activeDistance = Number.POSITIVE_INFINITY;
  const editorTop = els.editor.getBoundingClientRect().top;

  headings.forEach((heading) => {
    const distance = Math.abs(heading.getBoundingClientRect().top - editorTop - 24);
    if (distance < activeDistance) {
      activeDistance = distance;
      active = heading.dataset.navIndex;
    }
  });

  els.headingNav.querySelectorAll(".heading-nav-item").forEach((button) => {
    button.classList.toggle("active", button.dataset.navIndex === active);
  });
}

function scrollToDocumentPercent(percent) {
  const maxScroll = Math.max(0, els.editor.scrollHeight - els.editor.clientHeight);
  els.editor.scrollTop = maxScroll * (Number(percent) / 100);
  els.documentPositionOutput.textContent = `${Math.round(Number(percent))}%`;
}

function hideEditorContextMenu() {
  els.editorContextMenu.hidden = true;
}

function showEditorContextMenu(event) {
  saveEditorSelection();
  els.editorContextMenu.hidden = false;
  const menuBox = els.editorContextMenu.getBoundingClientRect();
  const left = Math.min(event.clientX, window.innerWidth - menuBox.width - 10);
  const top = Math.min(event.clientY, window.innerHeight - menuBox.height - 10);
  els.editorContextMenu.style.left = `${Math.max(10, left)}px`;
  els.editorContextMenu.style.top = `${Math.max(10, top)}px`;
}

function isInsertFootnoteShortcut(event) {
  return event.ctrlKey && event.altKey && !event.shiftKey && event.key.toLowerCase() === "f";
}

function setupInstallPrompt() {
  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    state.deferredInstallPrompt = event;
    els.installAppButton.hidden = false;
  });

  window.addEventListener("appinstalled", () => {
    state.deferredInstallPrompt = null;
    els.installAppButton.hidden = true;
    setStatus("האפליקציה הותקנה", "ready");
  });
}

async function installApp() {
  if (!state.deferredInstallPrompt) {
    setStatus("התקנה זמינה אחרי פתיחה דרך שרת מקומי או אתר מאובטח", "error");
    return;
  }

  state.deferredInstallPrompt.prompt();
  const choice = await state.deferredInstallPrompt.userChoice;
  state.deferredInstallPrompt = null;
  els.installAppButton.hidden = true;
  if (choice.outcome === "accepted") setStatus("האפליקציה הותקנה", "ready");
}

function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  if (location.protocol === "file:") return;

  navigator.serviceWorker.register("service-worker.js").catch((error) => {
    console.warn("Service worker registration failed", error);
  });
}

els.openFileButton.addEventListener("click", async () => {
  if (!window.showOpenFilePicker) {
    els.fileInput.click();
    return;
  }

  try {
    const [fileHandle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "Word document",
          accept: {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"],
          },
        },
      ],
    });
    const file = await fileHandle.getFile();
    await openDocx(file, fileHandle);
  } catch (error) {
    if (error.name !== "AbortError") {
      console.error(error);
      setStatus("לא הצלחתי לפתוח את קובץ Word", "error");
    }
  }
});

els.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  try {
    await openDocx(file, null);
  } catch (error) {
    console.error(error);
    setStatus(error.message, "error");
  } finally {
    event.target.value = "";
  }
});

document.querySelectorAll(".tool-menu").forEach((menu) => {
  menu.addEventListener("toggle", () => {
    if (!menu.open) return;
    closeToolMenus(menu);
  });
});

document.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === "b" && state.isDocx) {
    event.preventDefault();
    els.editor.focus();
    toggleWordBoldSelection();
    return;
  }

  if (isInsertFootnoteShortcut(event)) {
    event.preventDefault();
    hideEditorContextMenu();
    insertFootnote();
    return;
  }

  if (event.key === "Escape") {
    closeToolMenus();
    hideEditorContextMenu();
  }
});

document.addEventListener("selectionchange", saveEditorSelection);
document.addEventListener("click", (event) => {
  if (!event.target.closest("#editorContextMenu")) hideEditorContextMenu();
  if (!event.target.closest(".tool-menu")) {
    closeToolMenus();
  }
});

els.newDocButton.addEventListener("click", newDocument);
els.installAppButton.addEventListener("click", () => {
  installApp().catch((error) => {
    console.error(error);
    setStatus("לא הצלחתי לפתוח התקנה", "error");
  });
});
els.saveButton.addEventListener("click", () => {
  saveDocx().catch((error) => {
    console.error(error);
    setStatus("השמירה נכשלה. כדאי לנסות שוב או לשמור עותק מהטקסט.", "error");
  });
});
els.autoSaveButton.addEventListener("click", () => {
  toggleAutoSave().catch((error) => {
    console.error(error);
    setStatus("לא הצלחתי להפעיל שמירה אוטומטית", "error");
  });
});

document.querySelectorAll("[data-command]").forEach((button) => {
  button.addEventListener("click", () => runCommand(button.dataset.command));
});

els.insertFootnoteButton.addEventListener("pointerdown", (event) => {
  event.preventDefault();
  saveEditorSelection();
});

els.insertFootnoteButton.addEventListener("click", () => {
  insertFootnote();
});

els.contextInsertFootnoteButton.addEventListener("click", () => {
  hideEditorContextMenu();
  insertFootnote();
});

els.formatSelect.addEventListener("change", () => {
  applyParagraphFormat(els.formatSelect.value);
  els.formatSelect.value = "p";
});

document.querySelectorAll("[data-width]").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll("[data-width]").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    els.editorShell.classList.remove("split", "wide", "focus");
    els.editorShell.classList.add(button.dataset.width);
  });
});

els.zoomInput.addEventListener("input", () => {
  applyEditorZoom();
});

els.footnoteZoomInput.addEventListener("input", () => {
  zoomMode = "separate";
  setZoomMode("separate");
  applyFootnoteZoom();
});

document.querySelectorAll("[data-zoom-mode]").forEach((button) => {
  button.addEventListener("click", () => {
    setZoomMode(button.dataset.zoomMode);
  });
});

els.headingNav.addEventListener("click", (event) => {
  const button = event.target.closest(".heading-nav-item");
  if (!button) return;
  scrollToHeading(button.dataset.navIndex);
  button.closest(".tool-menu")?.removeAttribute("open");
});

els.documentPositionInput.addEventListener("pointerdown", () => {
  state.isDraggingDocumentPosition = true;
});

els.documentPositionInput.addEventListener("input", () => {
  scrollToDocumentPercent(els.documentPositionInput.value);
});

els.documentPositionInput.addEventListener("change", () => {
  state.isDraggingDocumentPosition = false;
  scrollToDocumentPercent(els.documentPositionInput.value);
});

els.toggleFootnotesButton.addEventListener("click", () => {
  els.footnotesPane.hidden = !els.footnotesPane.hidden;
  const hidden = els.footnotesPane.hidden;
  els.toggleFootnotesButton.textContent = hidden ? "הצג חלונית" : "הסתר חלונית";
  els.toggleFootnotesButton.classList.toggle("active", !hidden);
  els.toggleFootnotesButton.setAttribute("aria-pressed", String(!hidden));
  if (!hidden) setTimeout(syncFootnoteToVisibleParagraph, 0);
});

els.editor.addEventListener("input", markDocumentDirty);
function caretRangeAtPoint(x, y) {
  if (document.caretRangeFromPoint) return document.caretRangeFromPoint(x, y);
  const position = document.caretPositionFromPoint?.(x, y);
  if (!position) return null;
  const range = document.createRange();
  range.setStart(position.offsetNode, position.offset);
  range.collapse(true);
  return range;
}

function moveFootnoteReferenceToPoint(id, x, y) {
  const range = caretRangeAtPoint(x, y);
  const targetNode = range?.startContainer?.nodeType === Node.TEXT_NODE ? range.startContainer.parentElement : range?.startContainer;
  const targetBlock = targetNode?.closest?.(".docx-paragraph, p, h1, h2, h3");
  if (!range || !targetBlock || !els.editor.contains(targetBlock)) return false;
  const targetRun = targetNode.closest?.(".docx-run[data-text-id]");

  Array.from(qName(state.documentXml, "footnoteReference"))
    .filter((reference) => attr(reference, "id") === id)
    .forEach((reference) => {
      const run = reference.parentNode;
      reference.remove();
      if (run?.localName === "r" && !qName(run, "t").length) run.remove();
    });
  els.editor.querySelectorAll(`.footnote-ref[data-footnote-ref="${CSS.escape(id)}"]`).forEach((ref) => ref.remove());

  const inserted = targetRun
    ? insertReferenceIntoDocxRun(range, targetRun, id)
    : appendReferenceToDocxParagraph(targetBlock, id);
  if (!inserted) return false;
  setActiveFootnote(id, false);
  markDocumentDirty();
  return true;
}

els.editor.addEventListener("dragstart", (event) => {
  if (event.target.closest?.(".footnote-ref")) event.preventDefault();
});
els.editor.addEventListener("pointerdown", (event) => {
  const ref = event.target.closest?.(".footnote-ref[data-footnote-ref]");
  if (!ref) return;
  state.footnotePointerMove = { id: ref.dataset.footnoteRef, x: event.clientX, y: event.clientY, moved: false };
  try { ref.setPointerCapture?.(event.pointerId); } catch (_) { /* Synthetic events may not own a pointer. */ }
});
els.editor.addEventListener("pointermove", (event) => {
  const move = state.footnotePointerMove;
  if (!move) return;
  if (Math.hypot(event.clientX - move.x, event.clientY - move.y) > 5) move.moved = true;
  if (move.moved) event.preventDefault();
});
els.editor.addEventListener("pointerup", (event) => {
  const move = state.footnotePointerMove;
  state.footnotePointerMove = null;
  if (!move?.moved) return;
  event.preventDefault();
  moveFootnoteReferenceToPoint(move.id, event.clientX, event.clientY);
});
els.editor.addEventListener("pointercancel", () => {
  state.footnotePointerMove = null;
});
els.editor.addEventListener("paste", (event) => {
  const html = event.clipboardData?.getData("text/html") || "";
  if (!html || !/mso-|MsoFootnote|MsoEndnote/i.test(html)) return;
  event.preventDefault();
  const doc = parser.parseFromString(html, "text/html");
  doc.querySelectorAll('[style*="mso-element:footnote" i], [style*="mso-element:endnote" i], .MsoFootnoteText, .MsoEndnoteText, [id^="ftn"], [id^="edn"]').forEach((node) => node.remove());
  doc.querySelectorAll('a[href^="#_ftn"], a[href^="#_edn"]').forEach((node) => node.remove());
  document.execCommand("insertHTML", false, doc.body.innerHTML);
  markDocumentDirty();
});
els.footnotesList.addEventListener("input", markDocumentDirty);
els.footnotesList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-delete-footnote]");
  if (!button) return;
  const id = button.dataset.deleteFootnote;
  if (window.confirm(`למחוק את הערת השוליים ${id} ואת כל ההפניות אליה?`)) removeFootnote(id);
});
els.editor.addEventListener("scroll", () => {
  hideEditorContextMenu();
  syncFootnoteToVisibleParagraph();
  updateDocumentPositionSlider();
});
els.editor.addEventListener("contextmenu", (event) => {
  const target = event.target.closest(".editor");
  if (!target) return;
  event.preventDefault();
  showEditorContextMenu(event);
});
els.editor.addEventListener("click", (event) => {
  const ref = event.target.closest(".footnote-ref");
  if (ref) setActiveFootnote(ref.dataset.footnoteRef);
});

els.footnoteResizeHandle.addEventListener("pointerdown", (event) => {
  state.isResizingFootnotes = true;
  els.footnoteResizeHandle.setPointerCapture(event.pointerId);
  document.body.classList.add("resizing-footnotes");
});

els.footnoteResizeHandle.addEventListener("pointermove", (event) => {
  if (!state.isResizingFootnotes) return;
  const shellBottom = els.editorShell.getBoundingClientRect().bottom;
  setFootnotesHeight(shellBottom - event.clientY);
});

els.footnoteResizeHandle.addEventListener("pointerup", (event) => {
  state.isResizingFootnotes = false;
  els.footnoteResizeHandle.releasePointerCapture(event.pointerId);
  document.body.classList.remove("resizing-footnotes");
});

els.footnoteResizeHandle.addEventListener("keydown", (event) => {
  if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
  event.preventDefault();
  const currentHeight = els.footnotesPane.getBoundingClientRect().height;
  setFootnotesHeight(currentHeight + (event.key === "ArrowUp" ? 24 : -24));
});

setupInstallPrompt();
registerServiceWorker();
restoreLocalDraft();
