const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const merge = require('./operations/merge');
const split = require('./operations/split');
const organize = require('./operations/organize');
const watermark = require('./operations/watermark');
const pageNumbers = require('./operations/pageNumbers');
const pdfToImages = require('./operations/pdfToImages');
const imagesToPdf = require('./operations/imagesToPdf');
const protectPdf = require('./operations/protectPdf');
const unlockPdf = require('./operations/unlockPdf');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    icon: path.join(__dirname, 'assets', 'logo.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

app.whenReady().then(createWindow);

ipcMain.handle('merge', async (_, files, outPath) => {
  await merge(files, outPath);
  return true;
});

ipcMain.handle('showSaveDialog', async (_, options = {}) => {
  const defaultOptions = {
    title: 'Save Merged PDF',
    defaultPath: 'merged.pdf',
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  };

  const dialogOptions = { ...defaultOptions, ...options };
  const result = await dialog.showSaveDialog(mainWindow, dialogOptions);

  return result.canceled ? null : result.filePath;
});

ipcMain.handle('openFile', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select PDF File',
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
    properties: ['openFile'],
  });

  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('selectFolder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Output Folder',
    properties: ['openDirectory'],
  });

  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('selectImage', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Watermark Image',
    filters: [
      { name: 'Images', extensions: ['png', 'jpg', 'jpeg'] }
    ],
    properties: ['openFile'],
  });

  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('split', async (_, inputPath, ranges, outputDir) => {
  const results = await split(inputPath, ranges, outputDir);
  return results;
});

ipcMain.handle('getPDFPageCount', async (_, filePath) => {
  const fs = require('fs');
  const path = require('path');
  const { PDFDocument } = require('pdf-lib');

  try {
    const data = fs.readFileSync(filePath);
    const doc = await PDFDocument.load(data, {
      ignoreEncryption: true,
      updateMetadata: false
    });
    return doc.getPageCount();
  } catch (error) {
    const fileName = path.basename(filePath);
    throw new Error(
      `Failed to read PDF "${fileName}": ${error.message}. ` +
      `This PDF file appears to be corrupted or has an invalid structure. ` +
      `Please verify the file is valid by opening it in another PDF viewer.`
    );
  }
});

ipcMain.handle('getPDFPageDimensions', async (_, filePath) => {
  const fs = require('fs');
  const path = require('path');
  const { PDFDocument } = require('pdf-lib');

  try {
    const data = fs.readFileSync(filePath);
    const doc = await PDFDocument.load(data, {
      ignoreEncryption: true,
      updateMetadata: false
    });
    const firstPage = doc.getPage(0);
    const { width, height } = firstPage.getSize();
    return { width, height };
  } catch (error) {
    const fileName = path.basename(filePath);
    throw new Error(
      `Failed to read PDF dimensions "${fileName}": ${error.message}.`
    );
  }
});

ipcMain.handle('getFileSize', async (_, filePath) => {
  const fs = require('fs');
  const stats = fs.statSync(filePath);
  return stats.size;
});

ipcMain.handle('organize', async (_, inputPath, operations, outputPath) => {
  await organize(inputPath, operations, outputPath);
  return true;
});

ipcMain.handle('watermark', async (_, inputPath, options, outputPath) => {
  await watermark(inputPath, options, outputPath);
  return true;
});

ipcMain.handle('pageNumbers', async (_, inputPath, options, outputPath) => {
  await pageNumbers(inputPath, options, outputPath);
  return true;
});

ipcMain.handle('pdfToImages', async (_, inputPath, options) => {
  const files = await pdfToImages(inputPath, options);
  return files;
});

ipcMain.handle('imagesToPdf', async (_, imagePaths, options, outputPath) => {
  await imagesToPdf(imagePaths, options, outputPath);
  return true;
});

ipcMain.handle('selectImages', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Images',
    filters: [
      { name: 'Images', extensions: ['png', 'jpg', 'jpeg'] }
    ],
    properties: ['openFile', 'multiSelections'],
  });

  return result.canceled ? null : result.filePaths;
});

ipcMain.handle('protectPdf', async (_, inputPath, options, outputPath) => {
  await protectPdf(inputPath, options, outputPath);
  return true;
});

ipcMain.handle('unlockPdf', async (_, inputPath, password, outputPath) => {
  await unlockPdf(inputPath, password, outputPath);
  return true;
});

ipcMain.handle('compressPdf', async (_, inputPath, options, outputPath) => {
  const compressPdf = require('./operations/compressPdf');
  const result = await compressPdf(inputPath, options, outputPath);
  return result;
});

ipcMain.handle('readMetadata', async (_, inputPath) => {
  const { readMetadata } = require('./operations/editMetadata');
  const metadata = await readMetadata(inputPath);
  return metadata;
});

ipcMain.handle('updateMetadata', async (_, inputPath, metadata, outputPath) => {
  const { updateMetadata } = require('./operations/editMetadata');
  const result = await updateMetadata(inputPath, metadata, outputPath);
  return result;
});

ipcMain.handle('convert', async (_, conversionType, inputPath, outputPath) => {
  const convert = require('./operations/convert');
  const result = await convert(conversionType, inputPath, outputPath);
  return result;
});

ipcMain.handle('getPDFThumbnails', async (_, filePath) => {
  const fs = require('fs');
  const os = require('os');
  const path = require('path');
  const poppler = require('pdf-poppler');

  try {
    // Create temporary directory for thumbnails
    const tempDir = path.join(os.tmpdir(), `pdf-thumbnails-${Date.now()}`);
    fs.mkdirSync(tempDir, { recursive: true });

    const opts = {
      format: 'jpeg',
      out_dir: tempDir,
      out_prefix: 'page',
      page: null, // All pages
      scale: 150, // Low resolution for thumbnails
    };

    await poppler.convert(filePath, opts);

    // Read generated images
    const files = fs.readdirSync(tempDir).sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)?.[0] || '0');
      const numB = parseInt(b.match(/\d+/)?.[0] || '0');
      return numA - numB;
    });

    const thumbnails = files.map(file => {
      const imgPath = path.join(tempDir, file);
      const imgData = fs.readFileSync(imgPath);
      const base64 = imgData.toString('base64');
      fs.unlinkSync(imgPath); // Clean up
      return `data:image/jpeg;base64,${base64}`;
    });

    // Clean up temp directory
    fs.rmdirSync(tempDir);

    return thumbnails;
  } catch (error) {
    console.error('Error generating thumbnails:', error);
    return [];
  }
});
