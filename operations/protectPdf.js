const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const { exec } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const execAsync = promisify(exec);

/**
 * Add password protection to a PDF using pdftk
 * @param {string} inputPath - Path to input PDF
 * @param {Object} options - Protection options
 *   - userPassword: string - Password to open the PDF (required)
 *   - ownerPassword: string - Password for permissions (optional)
 *   - allowPrinting: boolean - Allow printing
 *   - allowModification: boolean - Allow modification
 *   - allowCopying: boolean - Allow copying text
 * @param {string} outputPath - Path to save protected PDF
 */
module.exports = async (inputPath, options, outputPath) => {
  const {
    userPassword,
    ownerPassword,
    allowPrinting = true,
    allowModification = false,
    allowCopying = false
  } = options;

  if (!userPassword) {
    throw new Error('User password is required');
  }

  // Use local pdftk from project directory
  const getResourcePath = (relativePath) => {
    if (process.resourcesPath) {
      return path.join(process.resourcesPath, relativePath);
    }
    return path.join(__dirname, '..', relativePath);
  };
  
  const pdftkPath = getResourcePath('executables/bin/pdftk.exe');

  try {
    // Use pdftk for encryption (most reliable cross-platform solution)
    // Build permissions string
    const permissions = [];
    if (allowPrinting) permissions.push('Printing');
    if (allowModification) permissions.push('ModifyContents');
    if (allowCopying) permissions.push('CopyContents');
    
    const permString = permissions.length > 0 ? permissions.join(' ') : 'none';
    
    // PDFtk requires different passwords or no owner password
    // If owner password is same as user password or not provided, omit it
    let command;
    if (ownerPassword && ownerPassword.trim() && ownerPassword.trim() !== userPassword) {
      // Use both passwords if they are different
      command = `"${pdftkPath}" "${inputPath}" output "${outputPath}" user_pw "${userPassword}" owner_pw "${ownerPassword.trim()}" allow ${permString} encrypt_128bit`;
    } else {
      // Omit owner password if not provided or same as user password
      command = `"${pdftkPath}" "${inputPath}" output "${outputPath}" user_pw "${userPassword}" allow ${permString} encrypt_128bit`;
    }
    
    await execAsync(command);
    
    return true;
  } catch (error) {
    // Check if pdftk executable exists
    if (!fs.existsSync(pdftkPath)) {
      throw new Error(`PDFtk not found at: ${pdftkPath}`);
    }
    throw new Error(`Failed to protect PDF: ${error.message}`);
  }
};
