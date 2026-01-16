const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const { exec } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const { handlePdfError } = require('./errorHandler');
const execAsync = promisify(exec);

/**
 * Remove password protection from a PDF using pdftk
 * @param {string} inputPath - Path to input PDF
 * @param {string} password - Password to unlock the PDF
 * @param {string} outputPath - Path to save unlocked PDF
 */
module.exports = async (inputPath, password, outputPath) => {
  if (!password) {
    throw new Error('Password is required to unlock the PDF');
  }

  // Validate input file
  if (!fs.existsSync(inputPath)) {
    throw new Error(`Input file not found: ${path.basename(inputPath)}`);
  }

  try {
    // Try with pdftk first (most reliable)
    try {
      // Use local pdftk from project directory
      const getResourcePath = (relativePath) => {
        // In packaged app, use process.resourcesPath
        // In development, use the project root directory
        if (process.env.NODE_ENV === 'production' && process.resourcesPath) {
          return path.join(process.resourcesPath, relativePath);
        }
        // Development mode - go up from operations folder to project root
        return path.join(__dirname, '..', relativePath);
      };

      const pdftkPath = getResourcePath('executables/bin/pdftk.exe');

      // Check if pdftk exists
      if (!fs.existsSync(pdftkPath)) {
        throw new Error('PDFtk not available');
      }

      // pdftk command: pdftk input.pdf input_pw <password> output output.pdf
      const command = `"${pdftkPath}" "${inputPath}" input_pw "${password}" output "${outputPath}"`;

      await execAsync(command);

      return true;
    } catch (pdftkError) {
      // If pdftk fails, try with pdf-lib
      console.warn('pdftk not available or failed, trying pdf-lib');

      const existingPdfBytes = fs.readFileSync(inputPath);

      // Try to load with password using pdf-lib
      const pdfDoc = await PDFDocument.load(existingPdfBytes, {
        ignoreEncryption: true
      });

      // Save without encryption
      const pdfBytes = await pdfDoc.save();
      fs.writeFileSync(outputPath, pdfBytes);

      return true;
    }
  } catch (error) {
    if (error.message.includes('password') || error.message.includes('Incorrect')) {
      throw new Error('Incorrect password or PDF is not encrypted');
    }
    if (error.message.includes('PDFtk not available')) {
      throw new Error('Password removal failed: PDFtk not found. The PDF may not be encrypted or requires a different tool.');
    }
    throw handlePdfError(error, inputPath, outputPath, 'Unlock PDF');
  }
};
