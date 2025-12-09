const poppler = require('pdf-poppler');
const path = require('path');
const fs = require('fs');

/**
 * Convert PDF pages to images
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

  // Configure pdf-poppler options
  // Note: TIFF, BMP, WebP are not reliably supported by pdf-poppler, so we convert to PNG first
  const actualFormat = (format === 'jpeg') ? 'jpeg' : 'png';
  
  const popplerOptions = {
    format: actualFormat,
    out_dir: outputFolder,
    out_prefix: outputPrefix,
    page: null // Will be set based on pageRange
  };

  // Add format-specific options
  if (actualFormat === 'jpeg') {
    popplerOptions.jpegopt = {
      quality: quality,
      progressive: false
    };
  } else {
    popplerOptions.pngopt = {
      compression: 6
    };
  }

  // Set resolution using scale_to (DPI scaling)
  popplerOptions.scale_to = dpi;
  
  // Set color mode
  if (colorMode === 'grayscale') {
    popplerOptions.gray = true;
  } else if (colorMode === 'monochrome') {
    popplerOptions.mono = true;
  }

  // Parse page range
  if (pageRange !== 'all') {
    // Handle formats like "1-5", "1,3,5", "1-3,5-7"
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
    
    // pdf-poppler expects first and last page
    if (pages.length > 0) {
      pages.sort((a, b) => a - b);
      popplerOptions.page = pages[0];
      popplerOptions.last_page = pages[pages.length - 1];
    }
  }

  try {
    // Convert PDF to images
    await poppler.convert(inputPath, popplerOptions);

    // Get list of created files
    const actualExtension = actualFormat === 'jpeg' ? '.jpg' : '.png';
    let files = fs.readdirSync(outputFolder)
      .filter(file => file.startsWith(outputPrefix) && (file.endsWith('.png') || file.endsWith('.jpg')))
      .map(file => path.join(outputFolder, file))
      .sort();

    // Handle padded naming (001, 002, etc.)
    if (naming === 'padded' && files.length > 0) {
      const paddingLength = files.length.toString().length;
      files = files.map((file, index) => {
        const ext = path.extname(file);
        const pageNum = (index + 1).toString().padStart(paddingLength, '0');
        const newName = `${pdfName}-${pageNum}${ext}`;
        const newPath = path.join(outputFolder, newName);
        
        // Rename file
        fs.renameSync(file, newPath);
        return newPath;
      });
    }

    return files;
  } catch (error) {
    throw new Error(`Failed to convert PDF to images: ${error.message}`);
  }
};
