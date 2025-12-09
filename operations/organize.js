const { PDFDocument, degrees } = require('pdf-lib');
const fs = require('fs');

/**
 * Organize PDF pages - reorder, rotate, delete
 * @param {string} inputPath - Path to input PDF
 * @param {Array} operations - Array of page operations: 
 *   [{pageIndex: 0, rotation: 90, delete: false, newPosition: 0}, ...]
 * @param {string} outputPath - Path to save organized PDF
 */
module.exports = async (inputPath, operations, outputPath) => {
  const data = fs.readFileSync(inputPath);
  const sourceDoc = await PDFDocument.load(data);
  const newDoc = await PDFDocument.create();

  // Filter out deleted pages and sort by new position
  const activeOps = operations
    .filter((op) => !op.delete)
    .sort((a, b) => a.newPosition - b.newPosition);

  // Copy and modify pages
  for (const op of activeOps) {
    const [copiedPage] = await newDoc.copyPages(sourceDoc, [op.pageIndex]);

    // Apply rotation if specified
    if (op.rotation && op.rotation !== 0) {
      const currentRotation = copiedPage.getRotation().angle;
      copiedPage.setRotation(degrees(currentRotation + op.rotation));
    }

    newDoc.addPage(copiedPage);
  }

  const bytes = await newDoc.save();
  fs.writeFileSync(outputPath, bytes);
};
