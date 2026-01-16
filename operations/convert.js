const { execFile } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const fs = require('fs');

const execFileAsync = promisify(execFile);

// Path to converter executable - works in both dev and packaged app
const getResourcePath = (relativePath) => {
    // In packaged app, resources are in app.asar.unpacked or resources folder
    if (process.resourcesPath) {
        // Packaged app - use process.resourcesPath which points to resources folder
        return path.join(process.resourcesPath, relativePath);
    }
    // Development mode - go up from operations folder to project root
    return path.join(__dirname, '..', relativePath);
};

// Detect system architecture and use appropriate converter version
const is64Bit = process.arch === 'x64';
const converterExe = is64Bit ? 'converter64.exe' : 'converter32.exe';
const CONVERTER_PATH = getResourcePath(`executables/${converterExe}`);

/**
 * Convert files between different formats
 * @param {string} conversionType - Type of conversion (pdf-to-word, word-to-pdf, etc.)
 * @param {string} inputPath - Path to input file
 * @param {string} outputPath - Path to save converted file
 * @returns {Promise<Object>} Conversion result
 */
async function convert(conversionType, inputPath, outputPath) {
    try {
        // Check if converter exists
        if (!fs.existsSync(CONVERTER_PATH)) {
            throw new Error(`Converter not found at: ${CONVERTER_PATH}. Please build converter.exe first.`);
        }

        // Validate input file exists
        if (!fs.existsSync(inputPath)) {
            throw new Error(`Input file not found: ${path.basename(inputPath)}`);
        }

        // Validate conversion type
        const validTypes = ['pdf-to-word', 'word-to-pdf', 'pdf-to-excel', 'excel-to-pdf',
            'html-to-pdf', 'ppt-to-pdf', 'pdf-to-html', 'pdf-to-ppt',
            'pdf-to-png', 'pdf-to-jpeg', 'pdf-to-jpg', 'pdf-to-tiff'];
        if (!validTypes.includes(conversionType)) {
            throw new Error(`Invalid conversion type: ${conversionType}`);
        }

        // Execute converter
        const { stdout, stderr } = await execFileAsync(CONVERTER_PATH, [
            conversionType,
            inputPath,
            outputPath
        ]);

        // Check if output file was created
        if (!fs.existsSync(outputPath)) {
            throw new Error('Conversion failed: Output file not created');
        }

        // Get file sizes
        const inputSize = fs.statSync(inputPath).size;
        const outputSize = fs.statSync(outputPath).size;

        return {
            success: true,
            inputSize,
            outputSize,
            outputPath,
            message: stdout || 'Conversion completed successfully'
        };
    } catch (error) {
        console.error('Conversion error:', error);

        // Handle permission errors
        if (error.code === 'EACCES' || error.code === 'EPERM' || error.code === 'EBUSY' ||
            error.message.includes('Permission denied') || error.message.includes('Errno 13')) {
            const fileName = path.basename(outputPath);
            throw new Error(`Cannot save ${fileName}: File is open in another program or you don't have write permissions. Please close the file and try again.`);
        }

        // Handle corrupted file errors
        if (error.message.includes('parse') || error.message.includes('corrupted') ||
            error.message.includes('invalid') || error.message.includes('damaged') ||
            error.message.includes('format') || error.message.includes('reading')) {
            throw new Error(`The file "${path.basename(inputPath)}" appears to be corrupted or invalid. Please check the file and try again.`);
        }

        // Handle file not found errors
        if (error.message.includes('not found') || error.code === 'ENOENT') {
            throw error;
        }

        // Handle password protected files
        if (error.message.includes('password') || error.message.includes('encrypted')) {
            throw new Error('The file is password-protected. Please unlock it first using the Unlock PDF tool.');
        }

        // Handle Ghostscript errors
        if (error.message.includes('Ghostscript')) {
            throw new Error('Conversion failed: Required dependencies are missing.');
        }

        throw new Error(`Failed to convert file: ${error.message}`);
    }
}

module.exports = convert;
