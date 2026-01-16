const { execFile } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const fs = require('fs');
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
 * Convert PDF pages to images using Ghostscript
 * @param {string} inputPath - Path to input PDF
 * @param {Object} options - Conversion options
 *   - format: string - 'png', 'jpeg', 'tiff', 'bmp', 'webp'
 *   - quality: number - JPEG/WebP quality (1-100)
 *   - dpi: number - Output DPI (72-600)
 *   - pageRange: string - Page range (e.g., '1-5', 'all')
 *   - colorMode: string - 'color', 'grayscale', 'monochrome'
 *   - naming: string - 'default', 'page', 'padded', 'custom'
 *   - customPrefix: string - Custom prefix for naming
 *   - outputFolder: string - Folder to save images
 * @returns {Promise<Array>} - Array of created image file paths
 */
module.exports = async (inputPath, options) => {
  const {
    format = 'png',
    quality = 90,
    dpi = 150,
    pageRange = 'all',
    colorMode = 'color',
    naming = 'default',
    customPrefix = 'page',
    outputFolder
  } = options;

  if (!outputFolder) {
    throw new Error('Output folder is required');
  }

  // Check if Ghostscript exists
  if (!fs.existsSync(GHOSTSCRIPT_PATH)) {
    throw new Error(`Ghostscript not found at: ${GHOSTSCRIPT_PATH}`);
  }

  // Create output folder if it doesn't exist
  if (!fs.existsSync(outputFolder)) {
    fs.mkdirSync(outputFolder, { recursive: true });
  }

  // Get PDF name without extension for output naming
  const pdfName = path.basename(inputPath, path.extname(inputPath));
  
  // Determine output prefix based on naming option
  let outputPrefix = pdfName;
  if (naming === 'custom') {
    outputPrefix = customPrefix;
  } else if (naming === 'page') {
    outputPrefix = `${pdfName}-page`;
  } else if (naming === 'padded') {
    outputPrefix = pdfName; // Will handle padding in post-processing
  }

  // Map format to Ghostscript device
  const deviceMap = {
    'png': 'png16m',      // 24-bit RGB PNG
    'jpeg': 'jpeg',       // JPEG
    'jpg': 'jpeg',
    'tiff': 'tiff24nc',   // 24-bit RGB TIFF
    'tif': 'tiff24nc'
  };

  const device = deviceMap[format.toLowerCase()] || 'png16m';
  const ext = format.toLowerCase() === 'jpg' ? 'jpeg' : format.toLowerCase();

  // Output file pattern for Ghostscript
  const outputPattern = path.join(outputFolder, `${outputPrefix}-%d.${ext}`);

  // Build Ghostscript arguments
  const gsArgs = [
    '-dNOPAUSE',
    '-dBATCH',
    '-dSAFER',
    '-sDEVICE=' + device,
    `-r${dpi}`,                    // Set resolution
    '-dTextAlphaBits=4',           // Anti-aliasing for text
    '-dGraphicsAlphaBits=4',       // Anti-aliasing for graphics
  ];

  // Add JPEG quality if applicable
  if (format.toLowerCase() === 'jpeg' || format.toLowerCase() === 'jpg') {
    gsArgs.push(`-dJPEGQ=${quality}`);
  }

  // Add grayscale or monochrome
  if (colorMode === 'grayscale') {
    gsArgs[3] = '-sDEVICE=pnggray';  // Override device for grayscale
  } else if (colorMode === 'monochrome') {
    gsArgs[3] = '-sDEVICE=pngmono';  // Override device for monochrome
  }

  // Add page range if specified
  if (pageRange !== 'all') {
    // Parse page range
    const ranges = pageRange.split(',').map(r => r.trim());
    const pages = [];
    
    for (const range of ranges) {
      if (range.includes('-')) {
        const [start, end] = range.split('-').map(n => parseInt(n.trim()));
        for (let i = start; i <= end; i++) {
          if (!pages.includes(i)) pages.push(i);
        }
      } else {
        const pageNum = parseInt(range.trim());
        if (!pages.includes(pageNum)) pages.push(pageNum);
      }
    }
    
    if (pages.length > 0) {
      pages.sort((a, b) => a - b);
      gsArgs.push(`-dFirstPage=${pages[0]}`);
      gsArgs.push(`-dLastPage=${pages[pages.length - 1]}`);
    }
  }

  gsArgs.push(`-sOutputFile=${outputPattern}`);
  gsArgs.push(inputPath);

  try {
    // Execute Ghostscript
    await execFileAsync(GHOSTSCRIPT_PATH, gsArgs);

    // Get list of created files
    const fileExtension = `.${ext}`;
    let files = fs.readdirSync(outputFolder)
      .filter(file => file.startsWith(outputPrefix) && file.endsWith(fileExtension))
      .map(file => path.join(outputFolder, file))
      .sort();

    // Handle padded naming (001, 002, etc.)
    if (naming === 'padded' && files.length > 0) {
      const paddingLength = files.length.toString().length;
      files = files.map((file, index) => {
        const currentExt = path.extname(file);
        const pageNum = (index + 1).toString().padStart(paddingLength, '0');
        const newName = `${pdfName}-${pageNum}${currentExt}`;
        const newPath = path.join(outputFolder, newName);
        
        // Rename file
        fs.renameSync(file, newPath);
        return newPath;
      });
    }

    return files;
  } catch (error) {
    if (!fs.existsSync(GHOSTSCRIPT_PATH)) {
      throw new Error(`PDF to images conversion failed: Ghostscript not found at ${GHOSTSCRIPT_PATH}`);
    }
    throw handlePdfError(error, inputPath, '', 'PDF to images conversion');
  }
};
