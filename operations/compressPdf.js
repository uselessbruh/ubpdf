const { execFile } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');
const path = require('path');
const { handlePdfError } = require('./errorHandler');

const execFileAsync = promisify(execFile);

// Path to Ghostscript - works in both dev and packaged app
const getResourcePath = (relativePath) => {
  // In packaged app, use process.resourcesPath
  // In development, use the project root directory
  if (process.env.NODE_ENV === 'production' && process.resourcesPath) {
    return path.join(process.resourcesPath, relativePath);
  }
  // Development mode - go up from operations folder to project root
  return path.join(__dirname, '..', relativePath);
};

// Detect system architecture and use appropriate Ghostscript version
const is64Bit = process.arch === 'x64';
const gsFolder = is64Bit ? 'gs64' : 'gs32';
const gsExe = is64Bit ? 'gswin64c.exe' : 'gswin32c.exe';
const GHOSTSCRIPT_PATH = getResourcePath(`executables/${gsFolder}/bin/${gsExe}`);

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
    // Check if Ghostscript exists
    if (!fs.existsSync(GHOSTSCRIPT_PATH)) {
      throw new Error(`Compression failed: Ghostscript not found at ${GHOSTSCRIPT_PATH}`);
    }

    throw handlePdfError(error, inputPath, outputPath, 'Compress');
  }
};
