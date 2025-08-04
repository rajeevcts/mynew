const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const synopsisText = fs.readFileSync("project_synopsis.txt", "utf8");

const paragraphs = synopsisText.split(/\r?\n\r?\n/).map(
  (para) => new Paragraph({ children: [new TextRun(para)] })
);

const doc = new Document({
  sections: [
    {
      properties: {},
      children: paragraphs,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("project_synopsis.docx", buffer);
  console.log("Word document created: project_synopsis.docx");
});
