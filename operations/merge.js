const { PDFDocument } = require('pdf-lib');
const fs = require('fs');

module.exports = async (files, outPath) => {
  const output = await PDFDocument.create();

  for (const file of files) {
    const data = fs.readFileSync(file);
    const doc = await PDFDocument.load(data);
    const pages = await output.copyPages(doc, doc.getPageIndices());
    pages.forEach((page) => output.addPage(page));
  }

  const bytes = await output.save();
  fs.writeFileSync(outPath, bytes);
};
