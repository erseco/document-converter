/**
 * ZetaJS Document Converter for GitHub Pages
 * 
 * This module handles iframe communication with WordPress Playground
 * and manages document conversion using ZetaJS/LibreOffice WASM.
 * 
 * Message Protocol:
 * - Receives: { type: 'convert', file: ArrayBuffer, format: string, requestId: any }
 * - Sends: { type: 'ready' } when ZetaJS is initialized
 * - Sends: { type: 'result', blob: Blob, format: string, requestId: any } on success
 * - Sends: { type: 'error', error: string, requestId: any } on failure
 */

// ZetaJS CDN configuration - pinned version for stability
const ZETAJS_VERSION = '1.2.0';
const ZETAJS_BASE_URL = `https://cdn.jsdelivr.net/npm/zetajs@${ZETAJS_VERSION}`;

// Get allowed origins from URL parameter or allow all by default
// Usage: ?origins=https://example.com,https://another.com
const urlParams = new URLSearchParams(window.location.search);
const allowedOriginsParam = urlParams.get('origins');
const allowedOrigins = allowedOriginsParam 
    ? allowedOriginsParam.split(',').map(o => o.trim())
    : null; // null means allow all origins

let zeta = null;
let moduleReady = false;

// Format configuration maps
const formatExtensions = {
    'pdf': 'pdf',
    'docx': 'docx',
    'xlsx': 'xlsx',
    'pptx': 'pptx',
    'odt': 'odt',
    'ods': 'ods',
    'odp': 'odp',
    'html': 'html',
    'txt': 'txt',
    'rtf': 'rtf',
    'png': 'png',
    'jpg': 'jpg'
};

const filterMap = {
    'pdf': 'writer_pdf_Export',
    'docx': 'MS Word 2007 XML',
    'xlsx': 'Calc MS Excel 2007 XML',
    'pptx': 'Impress MS PowerPoint 2007 XML',
    'odt': 'writer8',
    'ods': 'calc8',
    'odp': 'impress8',
    'html': 'HTML (StarWriter)',
    'txt': 'Text',
    'rtf': 'Rich Text Format',
    'png': 'writer_png_Export',
    'jpg': 'writer_jpg_Export'
};

const mimeTypes = {
    'pdf': 'application/pdf',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'odt': 'application/vnd.oasis.opendocument.text',
    'ods': 'application/vnd.oasis.opendocument.spreadsheet',
    'odp': 'application/vnd.oasis.opendocument.presentation',
    'html': 'text/html',
    'txt': 'text/plain',
    'rtf': 'application/rtf',
    'png': 'image/png',
    'jpg': 'image/jpeg'
};

// DOM elements
const statusText = document.getElementById('statusText');
const spinner = document.getElementById('spinner');

/**
 * Validate if origin is allowed to communicate
 */
function isOriginAllowed(origin) {
    if (!allowedOrigins) return true; // Allow all if no whitelist
    if (origin === 'null') return true; // Allow null origin for file:// or sandboxed iframes
    return allowedOrigins.includes(origin);
}

/**
 * Get safe target origin for postMessage
 */
function getTargetOrigin(eventOrigin) {
    // For null origins (file:// or sandboxed iframes), use '*'
    if (eventOrigin === 'null' || !eventOrigin) return '*';
    return eventOrigin;
}

/**
 * Update the status display
 */
function updateStatus(message, showSpinner = true) {
    if (statusText) statusText.textContent = message;
    if (spinner) spinner.classList.toggle('hidden', !showSpinner);
}

/**
 * Notify parent window about status
 */
function notifyParent(message, targetOrigin = '*') {
    if (window.parent !== window) {
        window.parent.postMessage(message, targetOrigin);
    }
}

/**
 * Initialize ZetaJS module
 */
export async function initZetaJS() {
    try {
        updateStatus('Loading ZetaJS module...');
        
        // Import ZetaJS from CDN with pinned version
        const { createZetaModule } = await import(`${ZETAJS_BASE_URL}/zetajs.mjs`);
        
        updateStatus('Initializing LibreOffice WASM...');
        
        zeta = await createZetaModule({
            locateFile: (path) => `${ZETAJS_BASE_URL}/${path}`
        });
        
        moduleReady = true;
        updateStatus('Ready for document conversion', false);
        
        // Notify parent that we're ready
        notifyParent({ type: 'ready' });
        
        return true;
    } catch (error) {
        updateStatus(`Error initializing: ${error.message}`, false);
        console.error('ZetaJS initialization error:', error);
        notifyParent({ type: 'error', error: error.message });
        return false;
    }
}

/**
 * Convert document using ZetaJS
 * @param {ArrayBuffer} arrayBuffer - The document data
 * @param {string} outputFormat - Target format (pdf, docx, etc.)
 * @returns {Promise<Blob>} - The converted document
 */
export async function convertDocument(arrayBuffer, outputFormat = 'pdf') {
    if (!moduleReady || !zeta) {
        throw new Error('ZetaJS not initialized');
    }

    updateStatus('Converting document...');

    try {
        const inputData = new Uint8Array(arrayBuffer);

        // Write input file to virtual filesystem
        const inputPath = '/tmp/input';
        try {
            zeta.FS.writeFile(inputPath, inputData);
        } catch (fsError) {
            throw new Error(`Failed to write input file to filesystem: ${fsError.message}`);
        }

        const extension = formatExtensions[outputFormat.toLowerCase()] || 'pdf';
        const outputPath = `/tmp/output.${extension}`;
        const filterName = filterMap[outputFormat.toLowerCase()] || 'writer_pdf_Export';

        // Load the document
        const desktop = zeta.getDesktop();
        const doc = await desktop.loadComponentFromURL(
            `file://${inputPath}`,
            '_blank',
            0,
            []
        );

        if (!doc) {
            throw new Error('Failed to load document');
        }
        
        // Export to the desired format
        const exportProps = [
            { Name: 'FilterName', Value: filterName },
            { Name: 'Overwrite', Value: true }
        ];
        
        await doc.storeToURL(`file://${outputPath}`, exportProps);
        doc.close(true);
        
        // Read the output file
        const outputData = zeta.FS.readFile(outputPath);
        
        // Clean up temporary files
        try {
            zeta.FS.unlink(inputPath);
            zeta.FS.unlink(outputPath);
        } catch (e) {
            // Ignore cleanup errors
        }
        
        const mimeType = mimeTypes[extension] || 'application/octet-stream';
        const blob = new Blob([outputData], { type: mimeType });
        
        updateStatus('Conversion complete', false);
        
        return blob;
    } catch (error) {
        updateStatus(`Conversion error: ${error.message}`, false);
        throw error;
    }
}

/**
 * Check if ZetaJS is ready
 */
export function isReady() {
    return moduleReady;
}

/**
 * Handle incoming messages from parent iframe (WordPress)
 */
window.addEventListener('message', async (event) => {
    const { data, origin, source } = event;
    
    // Validate origin if whitelist is configured
    if (!isOriginAllowed(origin)) {
        console.warn(`Rejected message from unauthorized origin: ${origin}`);
        return;
    }

    const targetOrigin = getTargetOrigin(origin);
    
    // Handle conversion request
    // Supports both 'buffer' (original) and 'file' (WordPress) property names
    if (data && data.type === 'convert') {
        try {
            const inputBuffer = data.buffer || data.file;
            if (!inputBuffer) {
                throw new Error('No document buffer provided');
            }

            // Convert to ArrayBuffer with explicit type validation
            let arrayBuffer;
            if (inputBuffer instanceof ArrayBuffer) {
                arrayBuffer = inputBuffer;
            } else if (inputBuffer instanceof Uint8Array || ArrayBuffer.isView(inputBuffer)) {
                arrayBuffer = inputBuffer.buffer;
            } else if (typeof inputBuffer === 'object' && inputBuffer.buffer instanceof ArrayBuffer) {
                arrayBuffer = inputBuffer.buffer;
            } else {
                throw new Error('Invalid buffer type: expected ArrayBuffer or TypedArray');
            }

            const outputFormat = data.format || 'pdf';
            const blob = await convertDocument(arrayBuffer, outputFormat);

            // Send the result back to parent
            source.postMessage({
                type: 'result',
                blob: blob,
                format: outputFormat,
                requestId: data.requestId
            }, targetOrigin);
            
        } catch (error) {
            source.postMessage({
                type: 'error',
                error: error.message,
                requestId: data.requestId
            }, targetOrigin);
        }
    }
    
    // Handle ping/status check
    if (data && data.type === 'ping') {
        source.postMessage({
            type: 'pong',
            ready: moduleReady,
            requestId: data.requestId
        }, targetOrigin);
    }
});
