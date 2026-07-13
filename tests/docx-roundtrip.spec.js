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
