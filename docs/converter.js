/**
 * ZetaJS Document Converter for GitHub Pages
 * 
 * This module handles iframe communication with WordPress Playground
 * and manages document conversion using ZetaJS/LibreOffice WASM.
 * 
 * Uses locally hosted WASM files downloaded via GitHub Actions.
 * 
 * Message Protocol:
 * - Receives: { type: 'convert', file: ArrayBuffer, format: string, requestId: any }
 * - Sends: { type: 'ready' } when ZetaJS is initialized
 * - Sends: { type: 'result', blob: Blob, format: string, requestId: any } on success
 * - Sends: { type: 'error', error: string, requestId: any } on failure
 */

// Local paths for WASM files (downloaded by GitHub Actions)
const WASM_BASE_URL = './wasm';
const ZETAJS_BASE_URL = './vendor/zetajs';

// Get allowed origins from URL parameter or allow all by default
// Usage: ?origins=https://example.com,https://another.com
const urlParams = new URLSearchParams(window.location.search);
const allowedOriginsParam = urlParams.get('origins');
const allowedOrigins = allowedOriginsParam 
    ? allowedOriginsParam.split(',').map(o => o.trim())
    : null; // null means allow all origins

let zHM = null;
let moduleReady = false;

// Store pending conversion requests
const pendingRequests = new Map();

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
 * Create office thread script as blob URL
 */
function createOfficeThreadBlob() {
    const zetajsUrl = new URL(`${ZETAJS_BASE_URL}/zetaHelper.js`, window.location.href).href;
    
    const code = `
        import { ZetaHelperThread } from '${zetajsUrl}';
        
        const zHT = new ZetaHelperThread();
        const zetajs = zHT.zetajs;
        const css = zHT.css;
        
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
        
        const bean_hidden = new css.beans.PropertyValue({Name: 'Hidden', Value: true});
        const bean_overwrite = new css.beans.PropertyValue({Name: 'Overwrite', Value: true});
        
        zHT.thrPort.onmessage = (e) => {
            switch (e.data.cmd) {
                case 'convert':
                    try {
                        const { from, to, format, requestId } = e.data;
                        const filterName = filterMap[format.toLowerCase()] || 'writer_pdf_Export';
                        const bean_filter = new css.beans.PropertyValue({Name: 'FilterName', Value: filterName});
                        
                        const xModel = zHT.desktop.loadComponentFromURL('file://' + from, '_blank', 0, [bean_hidden]);
                        if (!xModel) {
                            throw new Error('Failed to load document');
                        }
                        xModel.storeToURL('file://' + to, [bean_overwrite, bean_filter]);
                        xModel.close(true);
                        
                        zHT.thrPort.postMessage({cmd: 'converted', from, to, format, requestId});
                    } catch (err) {
                        const exc = zetajs.catchUnoException(err);
                        const errMsg = exc ? exc.Message : (err.message || String(err));
                        zHT.thrPort.postMessage({cmd: 'error', error: errMsg, requestId: e.data.requestId});
                    }
                    break;
                default:
                    console.warn('Unknown command in office thread:', e.data.cmd);
            }
        };
        
        zHT.thrPort.postMessage({cmd: 'ready'});
    `;
    const blob = new Blob([code], { type: 'text/javascript' });
    return URL.createObjectURL(blob);
}

/**
 * Handle messages from the worker thread
 */
function handleWorkerMessage(e) {
    switch (e.data.cmd) {
        case 'ready':
            moduleReady = true;
            updateStatus('Ready for document conversion', false);
            notifyParent({ type: 'ready' });
            break;
            
        case 'converted':
            {
                const { from, to, format, requestId } = e.data;
                const request = pendingRequests.get(requestId);
                if (request) {
                    try {
                        // Read the output file from the virtual filesystem
                        const outputData = zHM.FS.readFile(to);
                        
                        // Clean up temporary files
                        try {
                            zHM.FS.unlink(from);
                            zHM.FS.unlink(to);
                        } catch (cleanupErr) {
                            // Ignore cleanup errors
                        }
                        
                        const mimeType = mimeTypes[format] || 'application/octet-stream';
                        const blob = new Blob([outputData], { type: mimeType });
                        
                        updateStatus('Conversion complete', false);
                        
                        request.source.postMessage({
                            type: 'result',
                            blob: blob,
                            format: format,
                            requestId: requestId
                        }, request.targetOrigin);
                    } catch (err) {
                        request.source.postMessage({
                            type: 'error',
                            error: err.message,
                            requestId: requestId
                        }, request.targetOrigin);
                    }
                    pendingRequests.delete(requestId);
                }
            }
            break;
            
        case 'error':
            {
                const { error, requestId } = e.data;
                const request = pendingRequests.get(requestId);
                if (request) {
                    updateStatus(`Conversion error: ${error}`, false);
                    request.source.postMessage({
                        type: 'error',
                        error: error,
                        requestId: requestId
                    }, request.targetOrigin);
                    pendingRequests.delete(requestId);
                }
            }
            break;
    }
}

/**
 * Initialize ZetaJS module using local WASM files
 */
export async function initZetaJS() {
    try {
        updateStatus('Loading ZetaJS module...');
        
        // Import ZetaHelperMain from local vendor folder
        const { ZetaHelperMain } = await import(`${ZETAJS_BASE_URL}/zetaHelper.js`);
        
        updateStatus('Initializing LibreOffice WASM...');
        
        // Get absolute URL for WASM files
        const wasmUrl = new URL(WASM_BASE_URL + '/', window.location.href).href;
        
        // Create ZetaHelperMain with local WASM files
        zHM = new ZetaHelperMain(null, {
            wasmPkg: 'url:' + wasmUrl,
            blockPageScroll: false
        });

        // Set up message handler BEFORE starting
        const setupMessageHandler = () => {
            if (!zHM.thrPort) {
                // thrPort not ready yet, try again
                setTimeout(setupMessageHandler, 100);
                return;
            }

            zHM.thrPort.onmessage = (e) => {
                switch (e.data.cmd) {
                    case 'ZetaHelper::thr_started':
                        // Now inject our conversion logic
                        zHM.thrPort.postMessage({
                            cmd: 'ZetaHelper::run_thr_script',
                            threadJs: createOfficeThreadBlob(),
                            threadJsType: 'module'
                        });
                        break;
                    default:
                        handleWorkerMessage(e);
                }
            };
        };

        // Start ZetaOffice and setup handler
        zHM.start(() => {
            setupMessageHandler();
        });

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
 * @param {string} requestId - Request identifier
 * @param {MessageEventSource} source - Message source for response
 * @param {string} targetOrigin - Target origin for response
 */
export async function convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin) {
    if (!moduleReady || !zHM) {
        throw new Error('ZetaJS not initialized');
    }

    updateStatus('Converting document...');

    try {
        const inputData = new Uint8Array(arrayBuffer);
        const inputPath = '/tmp/input';
        const extension = formatExtensions[outputFormat.toLowerCase()] || 'pdf';
        const outputPath = `/tmp/output.${extension}`;
        
        // Write input file to virtual filesystem
        zHM.FS.writeFile(inputPath, inputData);
        
        // Store the request for later
        pendingRequests.set(requestId, { source, targetOrigin, format: outputFormat });
        
        // Send conversion request to worker
        zHM.thrPort.postMessage({
            cmd: 'convert',
            from: inputPath,
            to: outputPath,
            format: outputFormat,
            requestId: requestId
        });
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
            const requestId = data.requestId || Date.now().toString();
            
            await convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin);
            
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
