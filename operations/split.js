const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');
const { handlePdfError } = require('./errorHandler');

/**
 * Split PDF into multiple files based on ranges
 * @param {string} inputPath - Path to input PDF
 * @param {Array} ranges - Array of ranges: [{start: 1, end: 3, name: 'part1'}, ...]
 * @param {string} outputDir - Directory to save split PDFs
 */
module.exports = async (inputPath, ranges, outputDir) => {
  try {
    // Validate input file exists
    if (!fs.existsSync(inputPath)) {
      throw new Error(`Input file not found: ${path.basename(inputPath)}`);
    }

    // Validate ranges array
    if (!Array.isArray(ranges) || ranges.length === 0) {
      throw new Error('At least one page range is required');
    }

    // Create output directory if it doesn't exist
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const data = fs.readFileSync(inputPath);
    const sourceDoc = await PDFDocument.load(data, { 
      ignoreEncryption: true,
      updateMetadata: false 
    });
    const totalPages = sourceDoc.getPageCount();

    const results = [];

    for (const range of ranges) {
      const newDoc = await PDFDocument.create();
      const { start, end, name } = range;

      // Validate range
      const startIdx = Math.max(0, start - 1); // Convert to 0-based
      const endIdx = Math.min(totalPages - 1, end - 1);

      if (startIdx > endIdx || startIdx >= totalPages) {
        throw new Error(`Invalid range: pages ${start}-${end}`);
      }

      // Copy pages
      const pageIndices = [];
      for (let i = startIdx; i <= endIdx; i++) {
        pageIndices.push(i);
      }

      const copiedPages = await newDoc.copyPages(sourceDoc, pageIndices);
      copiedPages.forEach((page) => newDoc.addPage(page));

      // Save
      const outputPath = path.join(outputDir, `${name}.pdf`);
      const bytes = await newDoc.save();
      fs.writeFileSync(outputPath, bytes);

      results.push(outputPath);
    }

    return results;
  } catch (error) {
    throw handlePdfError(error, inputPath, '', 'Split');
  }
};
