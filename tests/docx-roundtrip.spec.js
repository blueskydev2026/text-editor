const { test, expect } = require("@playwright/test");
const JSZip = require("../vendor/jszip.min.js");
const { makeDocxFixture } = require("./docx-fixture");

test("opens, edits paragraph structure and produces a readable DOCX", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({
    name: "fixture.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: fixture,
  });

  await expect(page.locator("#editor")).toContainText("פסקה ראשונה");
  await expect(page.locator("#footnotesList")).toContainText("הערת בדיקה");

  await page.locator("#editor").evaluate((editor) => {
    const blocks = Array.from(editor.children);
    blocks[1].textContent = "פסקה ראשונה נערכה";
    const added = document.createElement("p");
    added.textContent = "פסקה חדשה";
    blocks[1].after(added);
    blocks[2].remove();
    editor.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText" }));
  });

  const downloadPromise = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const download = await downloadPromise;
  const path = await download.path();
  const output = await require("node:fs/promises").readFile(path);
  const zip = await JSZip.loadAsync(output);
  const documentXml = await zip.file("word/document.xml").async("text");

  expect(documentXml).toContain("פסקה ראשונה נערכה");
  expect(documentXml).toContain("פסקה חדשה");
  expect(documentXml).not.toContain("פסקה שנייה");
  expect(await zip.file("word/media/preserved.txt").async("text")).toBe("חלק שאסור למחוק");
  expect(zip.file("word/footnotes.xml")).toBeTruthy();
});

test("deletes a footnote and all of its references", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({
    name: "footnote.docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: fixture,
  });

  page.on("dialog", (dialog) => dialog.accept());
  await page.locator('[data-delete-footnote="1"]').click();
  await expect(page.locator('[data-footnote-id="1"]')).toHaveCount(0);
  await expect(page.locator('[data-footnote-ref="1"]')).toHaveCount(0);

  const downloadPromise = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const download = await downloadPromise;
  const output = await require("node:fs/promises").readFile(await download.path());
  const zip = await JSZip.loadAsync(output);
  const documentXml = await zip.file("word/document.xml").async("text");
  const footnotesXml = await zip.file("word/footnotes.xml").async("text");

  expect(documentXml).not.toMatch(/footnoteReference[^>]+w:id="1"/);
  expect(footnotesXml).not.toMatch(/<w:footnote[^>]+w:id="1"/);
  expect(footnotesXml).toContain('w:id="-1"');
  expect(footnotesXml).toContain('w:id="0"');
});

test("saves partial text deletion and whole paragraph deletion", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({ name: "delete.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: fixture });
  await page.locator("#editor").evaluate((editor) => {
    const blocks = Array.from(editor.children);
    blocks[1].querySelector("[data-text-id]").textContent = "פסקה";
    blocks[2].remove();
    editor.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward" }));
  });
  const pending = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const output = await require("node:fs/promises").readFile(await (await pending).path());
  const xml = await (await JSZip.loadAsync(output)).file("word/document.xml").async("text");
  expect(xml).toContain(">פסקה<");
  expect(xml).not.toContain("פסקה ראשונה");
  expect(xml).not.toContain("פסקה שנייה");
});

test("moving a footnote reference updates its Word paragraph", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({ name: "move.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: fixture });
  await page.locator("#editor").evaluate((editor) => {
    const blocks = Array.from(editor.children);
    blocks[1].append(blocks[2].querySelector(".footnote-ref"));
    editor.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText" }));
  });
  const pending = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const output = await require("node:fs/promises").readFile(await (await pending).path());
  const xml = await (await JSZip.loadAsync(output)).file("word/document.xml").async("text");
  expect(xml.indexOf("footnoteReference")).toBeLessThan(xml.indexOf("פסקה שנייה"));
});

test("deleting a reference from the document removes its orphaned footnote", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({ name: "orphan.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: fixture });
  await page.locator(".footnote-ref").evaluate((node) => node.remove());
  await page.locator("#editor").dispatchEvent("input");
  const pending = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const output = await require("node:fs/promises").readFile(await (await pending).path());
  const zip = await JSZip.loadAsync(output);
  expect(await zip.file("word/document.xml").async("text")).not.toContain("footnoteReference");
  expect(await zip.file("word/footnotes.xml").async("text")).not.toMatch(/<w:footnote[^>]+w:id="1"/);
});

test("Word paste omits the appended footnote body", async ({ page }) => {
  await page.goto("/");
  await page.locator("#editor").click();
  await page.locator("#editor").evaluate((editor) => {
    const data = new DataTransfer();
    data.setData("text/html", '<p class="MsoNormal">גוף מועתק<a href="#_ftn1">1</a></p><div style="mso-element:footnote-list"><div style="mso-element:footnote" id="ftn1"><p class="MsoFootnoteText">הערה מועתקת</p></div></div>');
    editor.dispatchEvent(new ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data }));
  });
  await expect(page.locator("#editor")).toContainText("גוף מועתק");
  await expect(page.locator("#editor")).not.toContainText("הערה מועתקת");
});

test("autosave persists body, footnote text and reference position", async ({ page }) => {
  const fixture = await makeDocxFixture();
  const base64 = fixture.toString("base64");
  await page.addInitScript(({ base64 }) => {
    const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
    window.__autoSaveWrites = [];
    window.__autoSaveAttempts = 0;
    const file = new File([bytes], "autosave.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    window.showOpenFilePicker = async () => [{
      getFile: async () => file,
      createWritable: async () => {
        window.__autoSaveAttempts += 1;
        if (window.__autoSaveAttempts === 1) throw new DOMException("changed on disk", "InvalidStateError");
        return {
          write: async (blob) => window.__autoSaveWrites.push(blob),
          close: async () => {},
        };
      },
    }];
  }, { base64 });

  await page.goto("/");
  await page.locator("#openFileButton").click();
  await expect(page.locator("#editor")).toContainText("פסקה ראשונה");
  await page.locator("#autoSaveButton").click();
  await expect.poll(() => page.evaluate(() => window.__autoSaveWrites.length)).toBeGreaterThan(0);

  await page.locator("#editor").evaluate((editor) => {
    const blocks = Array.from(editor.children);
    blocks[1].querySelector("[data-text-id]").textContent = "גוף שנשמר אוטומטית";
    blocks[1].append(blocks[2].querySelector(".footnote-ref"));
    editor.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText" }));
  });
  await page.locator('[data-footnote-id="1"] .footnote-body').evaluate((body) => {
    body.append(document.createTextNode(" ותוספת חדשה"));
    body.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText" }));
  });

  await expect(page.locator("#saveState")).toHaveText("נשמר אוטומטית בקובץ המקורי", { timeout: 5000 });
  const savedBytes = await page.evaluate(async () => Array.from(new Uint8Array(await window.__autoSaveWrites.at(-1).arrayBuffer())));
  const zip = await JSZip.loadAsync(Buffer.from(savedBytes));
  const documentXml = await zip.file("word/document.xml").async("text");
  const footnotesXml = await zip.file("word/footnotes.xml").async("text");
  expect(documentXml).toContain("גוף שנשמר אוטומטית");
  expect(documentXml.indexOf("footnoteReference")).toBeLessThan(documentXml.indexOf("פסקה שנייה"));
  expect(footnotesXml).toContain("ותוספת חדשה");
  expect(await page.locator("#autoSaveButton").getAttribute("aria-pressed")).toBe("true");
});

test("controlled footnote drag moves only the reference and preserves text", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({ name: "controlled-drag.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: fixture });
  await expect(page.locator("#editor")).toContainText("פסקה ראשונה");
  const before = await page.locator("#editor").evaluate((editor) => {
    const copy = editor.cloneNode(true);
    copy.querySelectorAll(".footnote-ref").forEach((node) => node.remove());
    return copy.innerText;
  });
  await page.locator("#editor").evaluate((editor) => {
    const ref = editor.querySelector(".footnote-ref");
    const target = editor.children[1].querySelector(".docx-run");
    const box = target.getBoundingClientRect();
    const start = ref.getBoundingClientRect();
    ref.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, cancelable: true, pointerId: 1, clientX: start.left, clientY: start.top }));
    ref.dispatchEvent(new PointerEvent("pointermove", { bubbles: true, cancelable: true, pointerId: 1, clientX: box.left + 2, clientY: box.top + box.height / 2 }));
    ref.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, cancelable: true, pointerId: 1, clientX: box.left + 2, clientY: box.top + box.height / 2 }));
  });
  const after = await page.locator("#editor").evaluate((editor) => {
    const copy = editor.cloneNode(true);
    copy.querySelectorAll(".footnote-ref").forEach((node) => node.remove());
    return copy.innerText;
  });
  expect(after.replace(/\s+/g, " ")).toBe(before.replace(/\s+/g, " "));
  await expect(page.locator("#editor").locator(".footnote-ref")).toHaveCount(1);
  const pending = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const output = await require("node:fs/promises").readFile(await (await pending).path());
  const zip = await JSZip.loadAsync(output);
  expect(await zip.file("word/document.xml").async("text")).toContain("פסקה ראשונה");
});

test("drag-style paragraph duplication is normalized instead of blocking save", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({ name: "drag.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: fixture });
  await page.locator("#editor").evaluate((editor) => {
    const source = editor.children[2];
    const duplicate = source.cloneNode(true);
    duplicate.querySelectorAll("[data-text-id]").forEach((node) => node.remove());
    duplicate.textContent = "פסקה שנוצרה בזמן גרירה";
    duplicate.append(source.querySelector(".footnote-ref"));
    source.after(duplicate);
    editor.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertFromDrop" }));
  });
  const pending = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const output = await require("node:fs/promises").readFile(await (await pending).path());
  const zip = await JSZip.loadAsync(output);
  expect(await zip.file("word/document.xml").async("text")).toContain("פסקה שנוצרה בזמן גרירה");
  expect(await zip.file("word/footnotes.xml").async("text")).toMatch(/<w:footnote[^>]+w:id="1"/);
});

test("new footnotes are renumbered by their order in the document", async ({ page }) => {
  const fixture = await makeDocxFixture();
  await page.goto("/");
  await page.locator("#fileInput").setInputFiles({ name: "numbering.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: fixture });
  await page.locator("#editor").evaluate((editor) => {
    const text = editor.children[0].querySelector("[data-text-id]").firstChild;
    const range = document.createRange();
    range.setStart(text, 1);
    range.collapse(true);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));
  });
  await page.locator("#insertFootnoteButton").click();
  const pending = page.waitForEvent("download");
  await page.locator("#saveButton").click();
  const output = await require("node:fs/promises").readFile(await (await pending).path());
  const zip = await JSZip.loadAsync(output);
  const documentXml = await zip.file("word/document.xml").async("text");
  const footnotesXml = await zip.file("word/footnotes.xml").async("text");
  const referenceIds = [...documentXml.matchAll(/footnoteReference[^>]+w:id="(\d+)"/g)].map((match) => match[1]);
  expect(referenceIds).toEqual(["1", "2"]);
  expect(footnotesXml).toMatch(/<w:footnote[^>]+w:id="1"/);
  expect(footnotesXml).toMatch(/<w:footnote[^>]+w:id="2"/);
  await expect(page.locator('[data-footnote-id="1"]')).toHaveCount(1);
  await expect(page.locator('[data-footnote-id="2"]')).toHaveCount(1);
});
