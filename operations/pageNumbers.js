const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const fs = require('fs');
const path = require('path');
const { handlePdfError } = require('./errorHandler');

/**
 * Convert number to different styles
 */
function convertNumberStyle(num, style) {
  switch (style) {
    case 'upperRoman':
      return toRoman(num).toUpperCase();
    case 'lowerRoman':
      return toRoman(num).toLowerCase();
    case 'upperAlpha':
      return toAlpha(num).toUpperCase();
    case 'lowerAlpha':
      return toAlpha(num).toLowerCase();
    case 'decimal':
    default:
      return num.toString();
  }
}

function toRoman(num) {
  const romanNumerals = [
    ['M', 1000], ['CM', 900], ['D', 500], ['CD', 400],
    ['C', 100], ['XC', 90], ['L', 50], ['XL', 40],
    ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]
  ];
  let result = '';
  for (const [roman, value] of romanNumerals) {
    while (num >= value) {
      result += roman;
      num -= value;
    }
  }
  return result;
}

function toAlpha(num) {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

/**
 * Add page numbers to PDF pages
 * @param {string} inputPath - Path to input PDF
 * @param {Object} options - Page number options
 *   - position: string - 'top', 'bottom', 'topLeft', 'topRight', 'bottomLeft', 'bottomRight'
 *   - fontSize: number - Font size
 *   - color: object - {r, g, b} (0-1)
 *   - style: string - 'decimal', 'upperRoman', 'lowerRoman', 'upperAlpha', 'lowerAlpha'
 *   - format: string - 'number', 'pageOfTotal', 'pageX', 'brackets', 'dashes'
 *   - startFrom: number - Starting page number
 *   - margin: number - Margin from edge in pixels
 *   - skipPages: array - Page numbers to skip (1-indexed)
 * @param {string} outputPath - Path to save PDF with page numbers
 */
module.exports = async (inputPath, options, outputPath) => {
  try {
    // Validate input file
    if (!fs.existsSync(inputPath)) {
      throw new Error(`Input file not found: ${path.basename(inputPath)}`);
    }

    const data = fs.readFileSync(inputPath);
    const pdfDoc = await PDFDocument.load(data, { 
      ignoreEncryption: true,
      updateMetadata: false 
    });
    const pages = pdfDoc.getPages();
    const totalPages = pages.length;

    const {
      position = 'bottom',
      fontSize = 12,
      color = { r: 0, g: 0, b: 0 },
      style = 'decimal',
      format = 'number',
      startFrom = 1,
      margin = 30,
      skipPages = [],
      fontStyle = 'Helvetica'
    } = options;

    // Map font style to StandardFonts
    const fontMap = {
      'Helvetica': StandardFonts.Helvetica,
      'Helvetica-Bold': StandardFonts.HelveticaBold,
      'Helvetica-Oblique': StandardFonts.HelveticaOblique,
      'Helvetica-BoldOblique': StandardFonts.HelveticaBoldOblique,
      'Times-Roman': StandardFonts.TimesRoman,
      'Times-Bold': StandardFonts.TimesBold,
      'Times-Italic': StandardFonts.TimesItalic,
      'Times-BoldItalic': StandardFonts.TimesBoldItalic,
      'Courier': StandardFonts.Courier,
      'Courier-Bold': StandardFonts.CourierBold,
      'Courier-Oblique': StandardFonts.CourierOblique,
      'Courier-BoldOblique': StandardFonts.CourierBoldOblique,
      'Symbol': StandardFonts.Symbol,
      'ZapfDingbats': StandardFonts.ZapfDingbats,
    };

    // Embed font
    const font = await pdfDoc.embedFont(fontMap[fontStyle] || StandardFonts.Helvetica);

    // Calculate adjusted numbering based on excluded pages
    let adjustedPageNumbers = [];
    let currentNumber = startFrom;
    
    pages.forEach((page, index) => {
      const actualPageNumber = index + 1;
      const excludePages = options.excludePages || [];
    
      if (excludePages.includes(actualPageNumber)) {
        adjustedPageNumbers.push(null); // Excluded completely
      } else {
        adjustedPageNumbers.push(currentNumber);
        currentNumber++;
      }
    });

    pages.forEach((page, index) => {
      const { width, height } = page.getSize();
      const actualPageNumber = index + 1; // 1-indexed page number
      const currentPageNumber = adjustedPageNumbers[index];
      
      // Scale font size based on page width relative to A4 (595pt reference)
      const referencePage = 595; // A4 width in points
      const scaleFactor = width / referencePage;
      const scaledFontSize = fontSize * scaleFactor;
      const scaledMargin = margin * scaleFactor;
      
      // Skip if page is excluded from counting
      if (currentPageNumber === null) {
        return;
      }
      
      // Skip this page if it's in the skip list (hide number but keep counting)
      if (skipPages.includes(actualPageNumber)) {
        return;
      }
      
      // Convert number to selected style
      const styledNumber = convertNumberStyle(currentPageNumber, style);
      const styledTotal = convertNumberStyle(totalPages, style);
      
      // Generate page number text based on format
      let pageText;
      switch (format) {
        case 'pageOfTotal':
          pageText = `Page ${styledNumber} of ${styledTotal}`;
          break;
        case 'pageX':
          pageText = `Page ${styledNumber}`;
          break;
        case 'brackets':
          pageText = `[${styledNumber}]`;
          break;
        case 'dashes':
          pageText = `- ${styledNumber} -`;
          break;
        case 'number':
        default:
          pageText = `${styledNumber}`;
      }
      
      // Calculate text width for positioning
      const textWidth = font.widthOfTextAtSize(pageText, scaledFontSize);

      // Determine position
      let x, y;
      switch (position) {
        case 'top':
          x = (width - textWidth) / 2;
          y = height - scaledMargin;
          break;
        case 'topLeft':
          x = scaledMargin;
          y = height - scaledMargin;
          break;
        case 'topRight':
          x = width - scaledMargin - textWidth;
          y = height - scaledMargin;
          break;
        case 'bottomLeft':
          x = scaledMargin;
          y = scaledMargin;
          break;
        case 'bottomRight':
          x = width - scaledMargin - textWidth;
          y = scaledMargin;
          break;
        case 'bottom':
        default:
          x = (width - textWidth) / 2;
          y = scaledMargin;
      }

      // Draw page number
      page.drawText(pageText, {
        x,
        y,
        size: scaledFontSize,
        font,
        color: rgb(color.r, color.g, color.b),
      });
    });

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
  } catch (error) {
    throw handlePdfError(error, inputPath, outputPath, 'Add page numbers');
  }
};
