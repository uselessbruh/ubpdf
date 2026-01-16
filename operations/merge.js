const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');
const { handlePdfError } = require('./errorHandler');

module.exports = async (files, outPath) => {
  try {
    const output = await PDFDocument.create();

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const fileName = path.basename(file);
      
      try {
        const data = fs.readFileSync(file);
        
        // Try loading with ignoreEncryption option for better compatibility
        const doc = await PDFDocument.load(data, { 
          ignoreEncryption: true,
          updateMetadata: false 
        });
        
        const pages = await output.copyPages(doc, doc.getPageIndices());
        pages.forEach((page) => output.addPage(page));
        
      } catch (error) {
        // Handle file-specific error with context
        const handledError = handlePdfError(error, file, '', 'Merge');
        throw new Error(`Failed to merge "${fileName}" (file ${i + 1} of ${files.length}): ${handledError.message}`);
      }
    }

    const bytes = await output.save();
    fs.writeFileSync(outPath, bytes);
  } catch (error) {
    // Handle output file errors
    throw handlePdfError(error, '', outPath, 'Merge');
  }
};
