const { execFile } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const fs = require('fs');

const execFileAsync = promisify(execFile);

// Path to converter executable
const CONVERTER_PATH = path.join(__dirname, '..', 'executables', 'converter.exe');

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
            throw new Error(`Input file not found: ${inputPath}`);
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
        throw new Error(`Failed to convert file: ${error.message}`);
    }
}

module.exports = convert;
