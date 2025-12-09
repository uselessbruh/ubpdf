const { execFile } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');
const path = require('path');

const execFileAsync = promisify(execFile);

// Path to Ghostscript in project directory
const GHOSTSCRIPT_PATH = path.join(__dirname, '..', 'executables', 'gswin64c.exe');

/**
 * Compress a PDF using Ghostscript for real image compression
 * @param {string} inputPath - Path to input PDF
 * @param {Object} options - Compression options
 *   - compressionLevel: 'low' | 'medium' | 'high'
 *   - imageQuality: number (10-100)
 *   - removeMetadata: boolean
 *   - compressImages: boolean
 * @param {string} outputPath - Path to save compressed PDF
 */
module.exports = async (inputPath, options, outputPath) => {
  const {
    compressionLevel = 'medium',
    imageQuality = 60,
    removeMetadata = true,
    compressImages = true
  } = options;

  try {
    // Get original file size
    const originalSize = fs.statSync(inputPath).size;

    // Map quality to Ghostscript settings
    let pdfSettings;
    if (compressionLevel === 'high') {
      pdfSettings = '/screen'; // 72 DPI - maximum compression
    } else if (compressionLevel === 'medium') {
      pdfSettings = '/ebook'; // 150 DPI - balanced
    } else {
      pdfSettings = '/printer'; // 300 DPI - minimal compression
    }

    // Prepare Ghostscript arguments
    const gsArgs = [
      '-sDEVICE=pdfwrite',
      '-dCompatibilityLevel=1.4',
      `-dPDFSETTINGS=${pdfSettings}`,
      '-dNOPAUSE',
      '-dQUIET',
      '-dBATCH',
      '-dCompressFonts=true',
      '-dCompressPages=true',
      '-dDetectDuplicateImages=true',
      `-sOutputFile=${outputPath}`,
      inputPath
    ];

    // Add image optimization if enabled
    if (compressImages) {
      // Calculate DPI based on imageQuality (10-100 maps to 72-300 DPI)
      const dpi = Math.round(72 + (imageQuality / 100) * 228);
      
      gsArgs.splice(7, 0,
        '-dDownsampleColorImages=true',
        `-dColorImageResolution=${dpi}`,
        '-dDownsampleGrayImages=true',
        `-dGrayImageResolution=${dpi}`,
        '-dDownsampleMonoImages=true',
        `-dMonoImageResolution=${dpi}`
      );
    }

    // Execute Ghostscript
    await execFileAsync(GHOSTSCRIPT_PATH, gsArgs);

    // Get compressed file size
    const compressedSize = fs.statSync(outputPath).size;
    const compressionRatio = ((1 - compressedSize / originalSize) * 100).toFixed(1);

    return {
      originalSize,
      compressedSize,
      compressionRatio,
      saved: originalSize - compressedSize
    };
  } catch (error) {
    throw new Error(`Failed to compress PDF: ${error.message}`);
  }
};
