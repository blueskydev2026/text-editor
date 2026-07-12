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
  activeFootnoteId: null,
  isResizingFootnotes: false,
  isDraggingDocumentPosition: false,
  deferredInstallPrompt: null,
  savedEditorRange: null,
};

const els = {
  toolsToggleButton: document.querySelector("#toolsToggleButton"),
  closeToolsButton: document.querySelector("#closeToolsButton"),
  toolsDrawer: document.querySelector("#toolsDrawer"),
  drawerBackdrop: document.querySelector("#drawerBackdrop"),
  openFileButton: document.querySelector("#openFileButton"),
  fileInput: document.querySelector("#fileInput"),
  newDocButton: document.querySelector("#newDocButton"),
  saveButton: document.querySelector("#saveButton"),
  exportTextButton: document.querySelector("#exportTextButton"),
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

function createFootnoteReferenceRun(id) {
  const run = createWordElement("r");
  const properties = createWordElement("rPr");
  const style = createWordElement("rStyle");
  setWordAttr(style, "val", "FootnoteReference");
  properties.append(style);
  run.append(properties);

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
  const referenceRun = createWordElement("r", doc);
  const referenceProperties = createWordElement("rPr", doc);
  const referenceStyle = createWordElement("rStyle", doc);
  setWordAttr(referenceStyle, "val", "FootnoteReference");
  referenceProperties.append(referenceStyle);
  referenceRun.append(referenceProperties, createWordElement("footnoteRef", doc));

  const textRun = createWordElement("r", doc);
  const text = createWordElement("t", doc);
  text.textContent = textNode;
  textRun.append(text);
  paragraph.append(referenceRun, textRun);
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

function renderRun(run) {
  const chunks = [];
  const children = Array.from(run.childNodes);
  children.forEach((child) => {
    if (child.localName === "t") {
      const textId = `t-${state.nextTextId++}`;
      state.textNodes.set(textId, child);
      chunks.push(
        `<span class="docx-run" data-text-id="${textId}" data-bold="${isBoldRun(run)}">${escapeHtml(child.textContent)}</span>`
      );
    }

    if (child.localName === "tab") chunks.push("	");
    if (child.localName === "br") chunks.push("<br>");
    if (child.localName === "footnoteReference") {
      const id = attr(child, "id");
      chunks.push(`<sup class="footnote-ref" contenteditable="false" data-footnote-ref="${escapeHtml(id)}">${escapeHtml(id)}</sup>`);
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
  const style = align === "center" ? "text-align:center" : align === "right" ? "text-align:right" : "";
  const styleId = styleInfo.id ? ` data-word-style-id="${escapeHtml(styleInfo.id)}"` : "";
  const styleName = styleInfo.name ? ` data-word-style-name="${escapeHtml(styleInfo.name)}"` : "";
  const styleLevel = tag !== "p" ? ` data-style-level="${tag}"` : "";
  return `<${tag} class="docx-paragraph word-style-${tag}" data-paragraph-index="${index}"${styleId}${styleName}${styleLevel} style="${style}">${content || "<br>"}</${tag}>`;
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
  if (!paragraph || !state.documentXml) return;

  const properties = ensureParagraphProperties(paragraph);
  removeWordChildren(properties, "jc");

  if (!alignment) return;

  const jc = state.documentXml.createElementNS(wordNs, "w:jc");
  setWordAttr(jc, "val", alignment);
  properties.append(jc);
}

function paragraphAlignmentFromBlock(block) {
  const value = (block.style.textAlign || "").toLowerCase();
  if (value === "center") return "center";
  if (value === "left") return "left";
  if (value === "right" || !value) return "right";
  if (value === "justify" || value === "start") return "both";
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

    setParagraphWordStyle(paragraph, formatFromBlock(block));
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
    const hasTrackedText = body?.querySelector("[data-footnote-text-id]");
    if (hasTrackedText) return;

    const textNodes = Array.from(qName(note, "t"));
    if (!textNodes.length) return;

    textNodes[0].textContent = visibleFootnoteText(card);
    textNodes.slice(1).forEach((node) => {
      node.textContent = "";
    });
  });
}

function updateXmlFromEditor() {
  state.textNodes.forEach((xmlNode, id) => {
    const htmlNode = els.editor.querySelector(`[data-text-id="${CSS.escape(id)}"]`);
    if (htmlNode) {
      xmlNode.textContent = htmlNode.textContent;
      setRunBold(xmlNode.parentNode, htmlNode.dataset.bold === "true");
    }
  });

  syncParagraphFormattingFromEditor();
  syncFootnotesFromEditor();
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

function exportTextFile() {
  const text = plainEditorText();
  if (!text) {
    setStatus("אין טקסט להורדה", "error");
    return;
  }

  const cleanName = (state.fileName || els.documentName.textContent || "מסמך")
    .replace(/\.docx$/i, "")
    .replace(/[\\/:*?"<>|]/g, "-")
    .trim() || "מסמך";
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  downloadBlob(blob, `${cleanName}.txt`);
  setStatus("קובץ הטקסט ירד למחשב", "ready");
}

async function createDocxBlob() {
  if (!state.zip || !state.documentXml) return;

  updateXmlFromEditor();
  await ensureDocxFootnotePackageParts();
  state.zip.file("word/document.xml", serializer.serializeToString(state.documentXml));
  if (state.footnotesXml) {
    state.zip.file("word/footnotes.xml", serializer.serializeToString(state.footnotesXml));
  }

  return state.zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
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
  if (!state.autoSaveEnabled || !state.openedFileHandle || state.isAutoSaving) return;

  state.isAutoSaving = true;
  try {
    setStatus("שומר אוטומטית...", "busy");
    const blob = await createDocxBlob();
    const writable = await state.openedFileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
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
  }
}

function scheduleAutoSave() {
  if (!state.autoSaveEnabled) return;
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

function runCommand(command, value = null) {
  els.editor.focus();
  document.execCommand(command, false, value);
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

  properties = state.documentXml.createElementNS(wordNs, "w:pPr");
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

function setToolsDrawerOpen(isOpen) {
  els.toolsDrawer.classList.toggle("open", isOpen);
  els.toolsDrawer.setAttribute("aria-hidden", String(!isOpen));
  els.toolsToggleButton.setAttribute("aria-expanded", String(isOpen));
  els.drawerBackdrop.hidden = !isOpen;
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

els.toolsToggleButton.addEventListener("click", () => {
  setToolsDrawerOpen(!els.toolsDrawer.classList.contains("open"));
});

els.closeToolsButton.addEventListener("click", () => {
  setToolsDrawerOpen(false);
});

els.drawerBackdrop.addEventListener("click", () => {
  setToolsDrawerOpen(false);
});

document.addEventListener("keydown", (event) => {
  if (isInsertFootnoteShortcut(event)) {
    event.preventDefault();
    hideEditorContextMenu();
    insertFootnote();
    return;
  }

  if (event.key === "Escape") {
    setToolsDrawerOpen(false);
    hideEditorContextMenu();
  }
});

document.addEventListener("selectionchange", saveEditorSelection);
document.addEventListener("click", (event) => {
  if (!event.target.closest("#editorContextMenu")) hideEditorContextMenu();
});

els.newDocButton.addEventListener("click", newDocument);
els.exportTextButton.addEventListener("click", exportTextFile);
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
  setToolsDrawerOpen(false);
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
els.footnotesList.addEventListener("input", markDocumentDirty);
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
