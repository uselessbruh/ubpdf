const { PDFDocument } = require('pdf-lib');
const fs = require('fs');

/**
 * Read metadata from a PDF file
 * @param {string} inputPath - Path to input PDF
 * @returns {Object} PDF metadata
 */
async function readMetadata(inputPath) {
  try {
    const existingPdfBytes = fs.readFileSync(inputPath);
    const pdfDoc = await PDFDocument.load(existingPdfBytes);

    return {
      title: pdfDoc.getTitle() || '',
      author: pdfDoc.getAuthor() || '',
      subject: pdfDoc.getSubject() || '',
      keywords: pdfDoc.getKeywords() || '',
      creator: pdfDoc.getCreator() || '',
      producer: pdfDoc.getProducer() || '',
      creationDate: pdfDoc.getCreationDate(),
      modificationDate: pdfDoc.getModificationDate()
    };
  } catch (error) {
    throw new Error(`Failed to read PDF metadata: ${error.message}`);
  }
}

/**
 * Update metadata in a PDF file
 * @param {string} inputPath - Path to input PDF
 * @param {Object} metadata - New metadata values
 * @param {string} outputPath - Path to save updated PDF
 */
async function updateMetadata(inputPath, metadata, outputPath) {
  try {
    const existingPdfBytes = fs.readFileSync(inputPath);
    const pdfDoc = await PDFDocument.load(existingPdfBytes);

    // Update metadata fields
    if (metadata.title !== undefined) {
      pdfDoc.setTitle(metadata.title);
    }
    if (metadata.author !== undefined) {
      pdfDoc.setAuthor(metadata.author);
    }
    if (metadata.subject !== undefined) {
      pdfDoc.setSubject(metadata.subject);
    }
    if (metadata.keywords !== undefined) {
      pdfDoc.setKeywords(Array.isArray(metadata.keywords) ? metadata.keywords : [metadata.keywords]);
    }
    if (metadata.creator !== undefined) {
      pdfDoc.setCreator(metadata.creator);
    }
    if (metadata.producer !== undefined) {
      pdfDoc.setProducer(metadata.producer);
    }

    // Update modification date
    pdfDoc.setModificationDate(new Date());

    // Save the updated PDF
    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);

    return { success: true };
  } catch (error) {
    throw new Error(`Failed to update PDF metadata: ${error.message}`);
  }
}

module.exports = { readMetadata, updateMetadata };
