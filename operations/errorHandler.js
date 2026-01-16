const path = require('path');

/**
 * Handle common PDF operation errors with user-friendly messages
 * @param {Error} error - The error object
 * @param {string} inputPath - Path to the input file
 * @param {string} outputPath - Path to the output file
 * @param {string} operation - Name of the operation (e.g., 'merge', 'compress')
 * @returns {Error} - A new error with user-friendly message
 */
function handlePdfError(error, inputPath = '', outputPath = '', operation = 'operation') {
    const inputFileName = inputPath ? path.basename(inputPath) : 'file';
    const outputFileName = outputPath ? path.basename(outputPath) : 'file';

    // Permission errors (file locked/open in another program)
    if (error.code === 'EACCES' || error.code === 'EPERM' || error.code === 'EBUSY' ||
        error.message.includes('Permission denied') || error.message.includes('Errno 13') ||
        error.message.includes('EACCES') || error.message.includes('EPERM')) {

        if (outputPath) {
            return new Error(`Cannot save "${outputFileName}": File is already open in another program. Please close it and try again.`);
        } else if (inputPath) {
            return new Error(`Cannot access "${inputFileName}": File is locked or you don't have permission. Please check file permissions.`);
        }
        return new Error('Permission denied: The file is locked or open in another program. Please close it and try again.');
    }

    // Corrupted or invalid file errors
    if (error.message.includes('parse') || error.message.includes('corrupt') ||
        error.message.includes('invalid') || error.message.includes('damaged') ||
        error.message.includes('malformed') || error.message.includes('reading') ||
        error.message.includes('Failed to parse') || error.message.includes('bad format')) {
        return new Error(`"${inputFileName}" appears to be corrupted or invalid. The file may be damaged. Try opening it in a PDF reader to verify.`);
    }

    // File not found errors
    if (error.code === 'ENOENT' || error.message.includes('not found') ||
        error.message.includes('does not exist') || error.message.includes('ENOENT')) {
        return new Error(`File not found: "${inputFileName}". The file may have been moved or deleted.`);
    }

    // Password protected / encrypted files
    if (error.message.includes('password') || error.message.includes('encrypt') ||
        error.message.includes('decrypt') || error.message.includes('locked')) {
        return new Error(`"${inputFileName}" is password-protected. Please unlock it first using the Unlock PDF tool.`);
    }

    // Out of memory errors
    if (error.message.includes('memory') || error.message.includes('heap') ||
        error.message.includes('ENOMEM')) {
        return new Error(`${operation} failed: File is too large or system is out of memory. Try closing other applications.`);
    }

    // Disk space errors
    if (error.message.includes('ENOSPC') || error.message.includes('disk') ||
        error.message.includes('space')) {
        return new Error(`${operation} failed: Not enough disk space. Please free up some space and try again.`);
    }

    // Empty or zero-size file
    if (error.message.includes('empty') || error.message.includes('zero') ||
        error.message.includes('no pages')) {
        return new Error(`"${inputFileName}" is empty or has no pages. Please check the file.`);
    }

    // Invalid PDF version or unsupported features
    if (error.message.includes('version') || error.message.includes('unsupported')) {
        return new Error(`"${inputFileName}" uses unsupported PDF features or an incompatible version.`);
    }

    // Default error message
    return new Error(`${operation} failed: ${error.message}`);
}

module.exports = { handlePdfError };
