const { PDFDocument, rgb, degrees, StandardFonts } = require('pdf-lib');
const fs = require('fs');

/**
 * Add watermark to PDF pages
 * @param {string} inputPath - Path to input PDF
 * @param {Object} options - Watermark options
 *   - type: 'text' or 'image'
 *   - text: string - Watermark text (for text type)
 *   - imagePath: string - Path to image (for image type)
 *   - opacity: number - 0 to 1
 *   - fontSize: number (for text)
 *   - rotation: number - degrees
 *   - color: object - {r, g, b} (0-1)
 *   - position: string - 'center', 'diagonal', 'top', 'bottom', 'custom'
 *   - x: number - custom X position (percentage)
 *   - y: number - custom Y position (percentage)
 *   - scale: number - image scale (0-1)
 * @param {string} outputPath - Path to save watermarked PDF
 */
module.exports = async (inputPath, options, outputPath) => {
  const data = fs.readFileSync(inputPath);
  const pdfDoc = await PDFDocument.load(data);
  const pages = pdfDoc.getPages();

  const {
    type = 'text',
    text = 'WATERMARK',
    imagePath = null,
    opacity = 0.3,
    fontSize = 48,
    fontName = 'Helvetica',
    rotation = 45,
    color = { r: 0.5, g: 0.5, b: 0.5 },
    position = 'diagonal',
    x = 50,
    y = 50,
    scale = 0.3
  } = options;

  let imageEmbed = null;
  let imageDims = null;
  let font = null;

  // Font mapping for all 14 standard PDF fonts
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

  // Load font for text measurements
  if (type === 'text') {
    const selectedFont = fontMap[fontName] || StandardFonts.Helvetica;
    font = await pdfDoc.embedFont(selectedFont);
  }

  // Load image if type is image
  if (type === 'image' && imagePath) {
    const imageBytes = fs.readFileSync(imagePath);
    const ext = imagePath.toLowerCase();

    if (ext.endsWith('.png')) {
      imageEmbed = await pdfDoc.embedPng(imageBytes);
    } else if (ext.endsWith('.jpg') || ext.endsWith('.jpeg')) {
      imageEmbed = await pdfDoc.embedJpg(imageBytes);
    } else {
      throw new Error('Unsupported image format. Use PNG or JPG');
    }

    imageDims = imageEmbed.scale(scale);
  }

  pages.forEach((page) => {
    const { width, height } = page.getSize();
    const margin = 50;

    let centerX, centerY;

    // Determine the center point where watermark should be placed
    if (position === 'custom') {
      centerX = (width * x) / 100;
      centerY = (height * y) / 100;
    } else {
      switch (position) {
        case 'center':
        case 'diagonal':
          centerX = width / 2;
          centerY = height / 2;
          break;
        case 'top':
          centerX = width / 2;
          centerY = height - margin;
          break;
        case 'bottom':
          centerX = width / 2;
          centerY = margin;
          break;
        case 'topLeft':
          centerX = margin;
          centerY = height - margin;
          break;
        case 'topRight':
          centerX = width - margin;
          centerY = height - margin;
          break;
        case 'bottomLeft':
          centerX = margin;
          centerY = margin;
          break;
        case 'bottomRight':
          centerX = width - margin;
          centerY = margin;
          break;
        default:
          centerX = width / 2;
          centerY = height / 2;
      }
    }

    if (type === 'image' && imageEmbed) {
      // For images, pdf-lib rotates around the bottom-left corner of the image
      // We need to place the image so its center is at centerX, centerY
      page.drawImage(imageEmbed, {
        x: centerX - imageDims.width / 2,
        y: centerY - imageDims.height / 2,
        width: imageDims.width,
        height: imageDims.height,
        opacity,
        rotate: degrees(-rotation),
      });
    } else {
      // For text, pdf-lib rotates around the x,y point (bottom-left of text)
      // To rotate around center, we need to:
      // 1. Calculate text dimensions
      // 2. Position so that when rotation happens at x,y, the center stays at centerX, centerY

      const textWidth = font.widthOfTextAtSize(text, fontSize);
      const textHeight = fontSize;

      // Convert rotation to radians for calculation
      const radians = (-rotation * Math.PI) / 180;

      // Calculate offset from center to bottom-left corner
      const offsetX = -textWidth / 2;
      const offsetY = -textHeight / 2;

      // Apply rotation to the offset to find where bottom-left should be
      const rotatedOffsetX = offsetX * Math.cos(radians) - offsetY * Math.sin(radians);
      const rotatedOffsetY = offsetX * Math.sin(radians) + offsetY * Math.cos(radians);

      // Position text so after rotation, its center is at centerX, centerY
      page.drawText(text, {
        x: centerX + rotatedOffsetX,
        y: centerY + rotatedOffsetY,
        size: fontSize,
        font: font,
        color: rgb(color.r, color.g, color.b),
        opacity,
        rotate: degrees(-rotation),
      });
    }
  });

  const bytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, bytes);
};
