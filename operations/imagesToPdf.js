const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');

/**
 * Convert multiple images to a single PDF
 * @param {Array<string>} imagePaths - Array of image file paths
 * @param {Object} options - Conversion options
 *   - pageSize: string - 'fit', 'a4', 'letter', 'legal'
 *   - orientation: string - 'portrait', 'landscape'
 *   - margin: number - Margin in points (default: 0)
 *   - quality: string - 'original', 'compressed'
 * @param {string} outputPath - Path to save the PDF
 */
module.exports = async (imagePaths, options, outputPath) => {
  const {
    pageSize = 'fit',
    orientation = 'portrait',
    margin = 0,
    quality = 'original'
  } = options;

  // Create a new PDF document
  const pdfDoc = await PDFDocument.create();

  // Page size presets (in points: 1 inch = 72 points)
  const pageSizes = {
    a4: { width: 595, height: 842 },      // 210 x 297 mm
    letter: { width: 612, height: 792 },   // 8.5 x 11 inches
    legal: { width: 612, height: 1008 }    // 8.5 x 14 inches
  };

  for (const imagePath of imagePaths) {
    try {
      const imageBytes = fs.readFileSync(imagePath);
      const ext = path.extname(imagePath).toLowerCase();

      let image;
      if (ext === '.png') {
        image = await pdfDoc.embedPng(imageBytes);
      } else if (ext === '.jpg' || ext === '.jpeg') {
        image = await pdfDoc.embedJpg(imageBytes);
      } else {
        console.warn(`Skipping unsupported image format: ${imagePath}`);
        continue;
      }

      const imageDims = image.scale(1);
      let pageWidth, pageHeight;

      // Determine page dimensions
      if (pageSize === 'fit') {
        // Fit page to image size
        if (orientation === 'landscape') {
          pageWidth = Math.max(imageDims.width, imageDims.height);
          pageHeight = Math.min(imageDims.width, imageDims.height);
        } else {
          pageWidth = imageDims.width;
          pageHeight = imageDims.height;
        }
      } else {
        // Use preset page size
        const preset = pageSizes[pageSize];
        if (orientation === 'landscape') {
          pageWidth = preset.height;
          pageHeight = preset.width;
        } else {
          pageWidth = preset.width;
          pageHeight = preset.height;
        }
      }

      // Create page with determined dimensions
      const page = pdfDoc.addPage([pageWidth, pageHeight]);

      // Calculate image dimensions with margin
      const availableWidth = pageWidth - (2 * margin);
      const availableHeight = pageHeight - (2 * margin);

      let imageWidth = imageDims.width;
      let imageHeight = imageDims.height;

      // Scale image to fit within available space while maintaining aspect ratio
      if (pageSize !== 'fit') {
        const widthRatio = availableWidth / imageWidth;
        const heightRatio = availableHeight / imageHeight;
        const scale = Math.min(widthRatio, heightRatio, 1); // Don't upscale

        imageWidth = imageWidth * scale;
        imageHeight = imageHeight * scale;
      }

      // Center image on page
      const x = (pageWidth - imageWidth) / 2;
      const y = (pageHeight - imageHeight) / 2;

      // Draw image on page
      page.drawImage(image, {
        x: x,
        y: y,
        width: imageWidth,
        height: imageHeight,
      });

    } catch (error) {
      console.error(`Error processing image ${imagePath}:`, error.message);
      throw new Error(`Failed to process image: ${path.basename(imagePath)}`);
    }
  }

  // Save PDF
  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, pdfBytes);

  return outputPath;
};
