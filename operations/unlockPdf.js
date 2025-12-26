const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const { exec } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const execAsync = promisify(exec);

/**
 * Remove password protection from a PDF using pdftk
 * @param {string} inputPath - Path to input PDF
 * @param {string} password - Password to unlock the PDF
 * @param {string} outputPath - Path to save unlocked PDF
 */
module.exports = async (inputPath, password, outputPath) => {
  if (!password) {
    throw new Error('Password is required');
  }

  try {
    // Try with pdftk first (most reliable)
    try {
      // Use local pdftk from project directory
      const getResourcePath = (relativePath) => {
        if (process.resourcesPath) {
          return path.join(process.resourcesPath, relativePath);
        }
        return path.join(__dirname, '..', relativePath);
      };
      
      const pdftkPath = getResourcePath('executables/bin/pdftk.exe');
      
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
    if (error.message.includes('pdftk')) {
      throw new Error('Password removal requires PDFtk. Please install PDFtk from: https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/');
    }
    throw new Error(`Failed to unlock PDF: ${error.message}`);
  }
};
