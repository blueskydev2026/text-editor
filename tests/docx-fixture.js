const JSZip = require("../vendor/jszip.min.js");

async function makeDocxFixture() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
    </Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`);
  zip.folder("word").file("document.xml", `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:bidi/></w:pPr><w:r><w:t>כותרת בדיקה</w:t></w:r></w:p>
      <w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>פסקה ראשונה</w:t></w:r></w:p>
      <w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>פסקה שנייה</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
      <w:sectPr/>
    </w:body></w:document>`);
  zip.folder("word").file("styles.xml", `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
    </w:styles>`);
  zip.folder("word").file("footnotes.xml", `<?xml version="1.0"?>
    <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:footnote w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
      <w:footnote w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
      <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>הערת בדיקה</w:t></w:r></w:p></w:footnote>
    </w:footnotes>`);
  zip.folder("word").folder("_rels").file("document.xml.rels", `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
    </Relationships>`);
  zip.file("word/media/preserved.txt", "חלק שאסור למחוק");
  return zip.generateAsync({ type: "nodebuffer" });
}

module.exports = { makeDocxFixture };
