/**
 * =============================================================================
 * ZetaJS Document Converter - Browser-based Document Conversion Service
 * =============================================================================
 *
 * This module provides document conversion capabilities using ZetaJS, which is
 * a JavaScript wrapper around LibreOffice compiled to WebAssembly (WASM).
 *
 * ARCHITECTURE OVERVIEW:
 * ----------------------
 * The converter runs entirely in the browser and uses a multi-threaded architecture:
 *
 * 1. MAIN THREAD (this file):
 *    - Handles iframe communication via postMessage API
 *    - Manages the Emscripten virtual filesystem (FS)
 *    - Coordinates conversion requests and responses
 *
 * 2. WORKER THREAD (created dynamically via Blob URL):
 *    - Runs the actual LibreOffice conversion logic
 *    - Has access to the UNO API via ZetaHelperThread
 *    - Operates on files in the virtual filesystem
 *
 * WHY THIS ARCHITECTURE:
 * ----------------------
 * - LibreOffice WASM requires SharedArrayBuffer for multi-threading
 * - SharedArrayBuffer requires Cross-Origin Isolation (COOP/COEP headers)
 * - The worker thread prevents UI blocking during heavy conversions
 * - Blob URLs allow dynamic script creation without external files
 *
 * COMMUNICATION FLOW:
 * -------------------
 * Parent Window                    Converter (iframe)                 Worker Thread
 *      |                                  |                                 |
 *      |-- postMessage({type:'convert'}) ->|                                |
 *      |                                  |-- FS.writeFile() -------------->|
 *      |                                  |-- thrPort.postMessage('convert') ->|
 *      |                                  |                                 |-- loadComponentFromURL()
 *      |                                  |                                 |-- storeToURL()
 *      |                                  |<- thrPort.postMessage('converted') -|
 *      |                                  |-- FS.readFile() --------------->|
 *      |<- postMessage({type:'result'}) --|                                 |
 *
 * MESSAGE PROTOCOL:
 * -----------------
 * Incoming (from parent window):
 *   - { type: 'convert', buffer: ArrayBuffer, format: string, requestId: any }
 *   - { type: 'ping', requestId: any }
 *
 * Outgoing (to parent window):
 *   - { type: 'ready' } - Sent when LibreOffice is fully initialized
 *   - { type: 'result', blob: Blob, format: string, requestId: any } - Conversion success
 *   - { type: 'error', error: string, requestId: any } - Conversion failure
 *   - { type: 'pong', ready: boolean, requestId: any } - Health check response
 *
 * DEPLOYMENT NOTES:
 * -----------------
 * - WASM files are downloaded by GitHub Actions and served locally
 * - Cross-Origin Isolation is enabled via coi-serviceworker.js
 * - Works on GitHub Pages, Vercel, or any static hosting with proper headers
 *
 * @author Document Converter Team
 * @license MIT
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

/**
 * Base URL for WASM files (soffice.js, soffice.wasm, soffice.data)
 * These files are downloaded from ZetaOffice CDN by GitHub Actions workflow
 * and stored locally to avoid CORS issues and reduce external dependencies.
 */
const WASM_BASE_URL = './wasm';

/**
 * Base URL for ZetaJS library files (zeta.js, zetaHelper.js)
 * Copied from the zetajs npm package by GitHub Actions workflow.
 */
const ZETAJS_BASE_URL = './vendor/zetajs';

// =============================================================================
// SECURITY: ORIGIN WHITELIST
// =============================================================================

/**
 * Optional origin whitelist for postMessage security.
 *
 * Usage: Add ?origins=https://example.com,https://another.com to the URL
 *
 * When configured, only messages from whitelisted origins will be processed.
 * If not configured (default), messages from any origin are accepted.
 * This is useful when embedding the converter in specific trusted domains.
 */
const urlParams = new URLSearchParams(window.location.search);
const allowedOriginsParam = urlParams.get('origins');
const allowedOrigins = allowedOriginsParam
    ? allowedOriginsParam.split(',').map(o => o.trim())
    : null; // null = allow all origins (open mode)

// =============================================================================
// STATE MANAGEMENT
// =============================================================================

/**
 * Reference to the ZetaHelperMain instance.
 * This object provides:
 * - FS: Emscripten virtual filesystem for reading/writing files
 * - thrPort: MessagePort for communicating with the worker thread
 * - start(): Method to initialize LibreOffice
 */
let zHM = null;

/**
 * Flag indicating whether LibreOffice is fully initialized and ready.
 * Set to true when the worker thread sends the 'ready' message.
 */
let moduleReady = false;

/**
 * Map of pending conversion requests.
 * Key: requestId (string)
 * Value: { source: MessageEventSource, targetOrigin: string, format: string }
 *
 * This allows us to route responses back to the correct parent window
 * when a conversion completes asynchronously.
 */
const pendingRequests = new Map();

// =============================================================================
// FORMAT CONFIGURATION
// =============================================================================

/**
 * Maps output format names to file extensions.
 * Used to generate appropriate output file paths in the virtual filesystem.
 */
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

/**
 * Maps output format names to LibreOffice export filter names.
 * These are the internal filter identifiers used by LibreOffice's UNO API
 * to determine the output format when calling storeToURL().
 *
 * Reference: https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
 */
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

/**
 * Maps output format names to MIME types.
 * Used when creating Blob objects for the converted documents.
 */
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

// =============================================================================
// DOM ELEMENTS
// =============================================================================

/**
 * Reference to the status text element in index.html.
 * Used to display loading progress and conversion status.
 */
const statusText = document.getElementById('statusText');

/**
 * Reference to the loading spinner element in index.html.
 * Shown during initialization and hidden when ready.
 */
const spinner = document.getElementById('spinner');

// =============================================================================
// SECURITY UTILITIES
// =============================================================================

/**
 * Validates if a message origin is allowed to communicate with the converter.
 *
 * @param {string} origin - The origin of the incoming message
 * @returns {boolean} True if the origin is allowed, false otherwise
 *
 * Security considerations:
 * - If no whitelist is configured, all origins are allowed (for development/public use)
 * - The 'null' origin is always allowed (for file:// URLs and sandboxed iframes)
 * - In production, consider using ?origins= parameter to restrict access
 */
function isOriginAllowed(origin) {
    if (!allowedOrigins) return true; // No whitelist = allow all
    if (origin === 'null') return true; // Allow null origin (file:// or sandboxed iframes)
    return allowedOrigins.includes(origin);
}

/**
 * Determines the appropriate target origin for postMessage responses.
 *
 * @param {string} eventOrigin - The origin from the incoming message event
 * @returns {string} The target origin to use in postMessage
 *
 * For security, we echo back the origin of the sender.
 * For null origins (file:// or sandboxed iframes), we must use '*'.
 */
function getTargetOrigin(eventOrigin) {
    if (eventOrigin === 'null' || !eventOrigin) return '*';
    return eventOrigin;
}

// =============================================================================
// UI UTILITIES
// =============================================================================

/**
 * Updates the status display in the converter iframe.
 *
 * @param {string} message - The status message to display
 * @param {boolean} showSpinner - Whether to show the loading spinner (default: true)
 */
function updateStatus(message, showSpinner = true) {
    if (statusText) statusText.textContent = message;
    if (spinner) spinner.classList.toggle('hidden', !showSpinner);
}

/**
 * Sends a message to the parent window (if embedded in an iframe).
 *
 * @param {Object} message - The message object to send
 * @param {string} targetOrigin - The target origin (default: '*' for any origin)
 *
 * This is used to notify the parent when:
 * - The converter is ready
 * - A conversion completes successfully
 * - An error occurs
 */
function notifyParent(message, targetOrigin = '*') {
    if (window.parent !== window) {
        window.parent.postMessage(message, targetOrigin);
    }
}

// =============================================================================
// WORKER THREAD SCRIPT CREATION
// =============================================================================

/**
 * Creates the worker thread script as a Blob URL.
 *
 * WHY A BLOB URL:
 * ---------------
 * ZetaJS requires a separate script to run in the worker thread where LibreOffice
 * actually executes. Instead of maintaining a separate file, we dynamically create
 * the script as a Blob URL. This has several advantages:
 *
 * 1. Single-file deployment: All conversion logic is in this file
 * 2. Dynamic URL generation: We can inject the correct absolute URLs
 * 3. No CORS issues: Blob URLs are same-origin by definition
 *
 * WHAT THE WORKER SCRIPT DOES:
 * ----------------------------
 * 1. Imports ZetaHelperThread from the zetajs library
 * 2. Sets up message handlers for conversion commands
 * 3. Uses LibreOffice's UNO API to load and convert documents
 * 4. Posts results back to the main thread
 *
 * THE UNO API:
 * ------------
 * LibreOffice's Universal Network Objects (UNO) API is used to:
 * - loadComponentFromURL(): Open a document from the virtual filesystem
 * - storeToURL(): Save the document in a different format
 * - close(): Release the document resources
 *
 * @returns {string} A Blob URL pointing to the worker script
 */
function createOfficeThreadBlob() {
    // Get absolute URL for the zetaHelper module (required for ES module import)
    const zetajsUrl = new URL(`${ZETAJS_BASE_URL}/zetaHelper.js`, window.location.href).href;
    console.log('converter.js: Creating thread script with zetaHelper URL:', zetajsUrl);

    // The worker thread script as a string
    // This will be executed in a separate thread by ZetaJS
    const code = `
        console.log('Worker thread: Script starting...');

        // Import ZetaHelperThread - this provides access to LibreOffice's UNO API
        import { ZetaHelperThread } from '${zetajsUrl}';

        console.log('Worker thread: ZetaHelperThread imported, initializing...');

        // Initialize the helper thread
        // This gives us access to:
        // - zHT.zetajs: The zetajs library
        // - zHT.css: LibreOffice's CSS (Component Service Set) for UNO types
        // - zHT.desktop: The LibreOffice desktop instance
        // - zHT.thrPort: MessagePort for communication with main thread
        const zHT = new ZetaHelperThread();
        const zetajs = zHT.zetajs;
        const css = zHT.css;

        console.log('Worker thread: ZetaHelperThread initialized');

        // LibreOffice export filter names (duplicated here as this runs in a separate context)
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

        // UNO property beans for document operations
        // Hidden=true: Don't show any UI when loading the document
        // Overwrite=true: Overwrite the output file if it exists
        const bean_hidden = new css.beans.PropertyValue({Name: 'Hidden', Value: true});
        const bean_overwrite = new css.beans.PropertyValue({Name: 'Overwrite', Value: true});

        // Handle messages from the main thread
        zHT.thrPort.onmessage = (e) => {
            console.log('Worker thread: Received message:', e.data.cmd);
            switch (e.data.cmd) {
                case 'convert':
                    try {
                        const { from, to, format, requestId } = e.data;
                        console.log('Worker thread: Converting', from, 'to', format);

                        // Get the appropriate filter for the output format
                        const filterName = filterMap[format.toLowerCase()] || 'writer_pdf_Export';
                        const bean_filter = new css.beans.PropertyValue({Name: 'FilterName', Value: filterName});

                        // Load the document from the virtual filesystem
                        // 'file://' prefix is required for Emscripten FS paths
                        // '_blank' opens in a new frame (required for conversion)
                        // 0 = no special flags
                        // [bean_hidden] = load invisibly
                        const xModel = zHT.desktop.loadComponentFromURL('file://' + from, '_blank', 0, [bean_hidden]);

                        if (!xModel) {
                            throw new Error('Failed to load document');
                        }

                        // Convert and save to the output path
                        // The filter determines the output format
                        xModel.storeToURL('file://' + to, [bean_overwrite, bean_filter]);

                        // Close the document to free resources
                        xModel.close(true);

                        console.log('Worker thread: Conversion complete');
                        // Notify main thread of success
                        zHT.thrPort.postMessage({cmd: 'converted', from, to, format, requestId});

                    } catch (err) {
                        // Handle both UNO exceptions and regular JavaScript errors
                        const exc = zetajs.catchUnoException(err);
                        const errMsg = exc ? exc.Message : (err.message || String(err));
                        console.error('Worker thread: Conversion error:', errMsg);
                        zHT.thrPort.postMessage({cmd: 'error', error: errMsg, requestId: e.data.requestId});
                    }
                    break;

                default:
                    console.warn('Worker thread: Unknown command:', e.data.cmd);
            }
        };

        // Signal to main thread that we're ready to accept conversion requests
        console.log('Worker thread: Sending ready message');
        zHT.thrPort.postMessage({cmd: 'ready'});
    `;

    // Create a Blob from the script code and return its URL
    const blob = new Blob([code], { type: 'text/javascript' });
    const blobUrl = URL.createObjectURL(blob);
    console.log('converter.js: Created thread script blob URL:', blobUrl);
    return blobUrl;
}

// =============================================================================
// WORKER MESSAGE HANDLING
// =============================================================================

/**
 * Handles messages received from the worker thread.
 *
 * Message types:
 * - 'ready': LibreOffice is initialized and ready for conversions
 * - 'converted': A conversion completed successfully
 * - 'error': A conversion failed
 *
 * @param {MessageEvent} e - The message event from the worker
 */
function handleWorkerMessage(e) {
    switch (e.data.cmd) {
        case 'ready':
            // LibreOffice and the worker thread are fully initialized
            moduleReady = true;
            updateStatus('Ready for document conversion', false);
            // Notify parent window that we're ready to accept conversion requests
            notifyParent({ type: 'ready' });
            break;

        case 'converted':
            {
                const { from, to, format, requestId } = e.data;
                const request = pendingRequests.get(requestId);

                if (request) {
                    try {
                        // Read the converted file from the virtual filesystem
                        // This returns a Uint8Array of the file contents
                        const outputData = zHM.FS.readFile(to);

                        // Clean up temporary files from the virtual filesystem
                        // We do this to prevent memory accumulation over multiple conversions
                        try {
                            zHM.FS.unlink(from);  // Delete input file
                            zHM.FS.unlink(to);    // Delete output file
                        } catch (cleanupErr) {
                            // Ignore cleanup errors - files might already be gone
                        }

                        // Create a Blob with the appropriate MIME type
                        const mimeType = mimeTypes[format] || 'application/octet-stream';
                        const blob = new Blob([outputData], { type: mimeType });

                        updateStatus('Conversion complete', false);

                        // Send the result back to the requesting window
                        request.source.postMessage({
                            type: 'result',
                            blob: blob,
                            format: format,
                            requestId: requestId
                        }, request.targetOrigin);

                    } catch (err) {
                        // Error reading the output file
                        request.source.postMessage({
                            type: 'error',
                            error: err.message,
                            requestId: requestId
                        }, request.targetOrigin);
                    }

                    // Remove the completed request from pending
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

                    // Forward the error to the requesting window
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

// =============================================================================
// INITIALIZATION
// =============================================================================

/**
 * Initializes the ZetaJS/LibreOffice WASM module.
 *
 * INITIALIZATION SEQUENCE:
 * ------------------------
 * 1. Import ZetaHelperMain from the local zetajs vendor folder
 * 2. Create the worker thread script as a Blob URL
 * 3. Create ZetaHelperMain with the thread script and WASM configuration
 * 4. Call start() to begin loading LibreOffice WASM (~150MB)
 * 5. Set up message handler to receive 'ready' signal from worker
 *
 * WHY THREAD SCRIPT IS PASSED TO CONSTRUCTOR:
 * -------------------------------------------
 * ZetaHelperMain needs to know about the worker script at construction time
 * so it can properly initialize the worker thread during start(). If we tried
 * to inject the script later (via message), we would miss the initialization
 * window and get "No threadJs given" errors.
 *
 * The threadJsType: 'module' option tells ZetaJS that our worker script uses
 * ES module syntax (import/export) rather than classic script syntax.
 *
 * @returns {Promise<boolean>} True if initialization succeeded, false otherwise
 */
export async function initZetaJS() {
    try {
        updateStatus('Loading ZetaJS module...');
        console.log('converter.js: Loading ZetaJS module from', `${ZETAJS_BASE_URL}/zetaHelper.js`);

        // Dynamically import ZetaHelperMain from local vendor folder
        // Using dynamic import() because this is an ES module
        const { ZetaHelperMain } = await import(`${ZETAJS_BASE_URL}/zetaHelper.js`);
        console.log('converter.js: ZetaHelperMain imported successfully');

        updateStatus('Initializing LibreOffice WASM...');

        // Convert relative WASM path to absolute URL
        // ZetaJS needs absolute URLs to properly fetch the WASM files
        const wasmUrl = new URL(WASM_BASE_URL + '/', window.location.href).href;
        console.log('converter.js: WASM URL:', wasmUrl);

        // Create the worker thread script BEFORE initializing ZetaHelperMain
        // This is critical - the script must exist when ZetaHelperMain starts
        const threadJsUrl = createOfficeThreadBlob();

        // Create ZetaHelperMain with all required configuration
        //
        // Parameters:
        // - threadJsUrl: URL to the worker script (created above as Blob URL)
        // - options object:
        //   - wasmPkg: Location of WASM files ('url:' prefix indicates a URL)
        //   - blockPageScroll: Don't interfere with page scrolling
        //   - threadJsType: 'module' for ES module syntax in worker script
        console.log('converter.js: Creating ZetaHelperMain with threadJs:', threadJsUrl);
        zHM = new ZetaHelperMain(threadJsUrl, {
            wasmPkg: 'url:' + wasmUrl,
            blockPageScroll: false,
            threadJsType: 'module'
        });
        console.log('converter.js: ZetaHelperMain created');

        /**
         * Sets up the message handler for worker thread communication.
         *
         * This function polls for zHM.thrPort availability because the
         * MessagePort is created asynchronously during LibreOffice loading.
         * Once available, it attaches our message handler.
         */
        const setupMessageHandler = () => {
            if (!zHM.thrPort) {
                // thrPort not ready yet, poll again in 100ms
                console.log('converter.js: thrPort not ready yet, polling...');
                setTimeout(setupMessageHandler, 100);
                return;
            }

            console.log('converter.js: thrPort is ready, attaching message handler');
            // Attach our message handler to receive worker responses
            zHM.thrPort.onmessage = (e) => {
                console.log('converter.js: Received message from worker:', e.data);
                handleWorkerMessage(e);
            };
        };

        // Start LibreOffice WASM loading
        // This triggers the download and initialization of ~150MB of WASM files
        console.log('converter.js: Starting ZetaOffice...');
        zHM.start(() => {
            console.log('converter.js: ZetaOffice start() callback fired');
            setupMessageHandler();
        });

        // Also set up handler immediately in case start() callback timing varies
        // This provides redundancy to ensure we don't miss messages
        setupMessageHandler();

        return true;

    } catch (error) {
        updateStatus(`Error initializing: ${error.message}`, false);
        console.error('ZetaJS initialization error:', error);
        notifyParent({ type: 'error', error: error.message });
        return false;
    }
}

// =============================================================================
// CONVERSION API
// =============================================================================

/**
 * Converts a document to the specified output format.
 *
 * CONVERSION FLOW:
 * ----------------
 * 1. Write the input document to Emscripten's virtual filesystem
 * 2. Store the request in pendingRequests map for later response routing
 * 3. Send a 'convert' message to the worker thread
 * 4. Worker thread loads document with LibreOffice and converts it
 * 5. Worker sends 'converted' message back
 * 6. handleWorkerMessage() reads the output and sends it to the parent
 *
 * @param {ArrayBuffer} arrayBuffer - The document data as an ArrayBuffer
 * @param {string} outputFormat - Target format (pdf, docx, xlsx, etc.)
 * @param {string} requestId - Unique identifier for this request
 * @param {MessageEventSource} source - The window that sent the request
 * @param {string} targetOrigin - Origin to use when sending the response
 *
 * @throws {Error} If ZetaJS is not initialized
 */
export async function convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin) {
    if (!moduleReady || !zHM) {
        throw new Error('ZetaJS not initialized');
    }

    updateStatus('Converting document...');

    try {
        // Convert ArrayBuffer to Uint8Array for filesystem operations
        const inputData = new Uint8Array(arrayBuffer);

        // File paths in Emscripten's virtual filesystem
        // We use /tmp because it's always writable
        const inputPath = '/tmp/input';
        const extension = formatExtensions[outputFormat.toLowerCase()] || 'pdf';
        const outputPath = `/tmp/output.${extension}`;

        // Write the input document to the virtual filesystem
        // This makes it accessible to LibreOffice running in the worker thread
        zHM.FS.writeFile(inputPath, inputData);

        // Store request details for routing the response later
        pendingRequests.set(requestId, { source, targetOrigin, format: outputFormat });

        // Send conversion command to worker thread
        // The worker will load the file, convert it, and save to outputPath
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
 * Checks if the converter is ready to accept conversion requests.
 *
 * @returns {boolean} True if LibreOffice is fully initialized
 */
export function isReady() {
    return moduleReady;
}

// =============================================================================
// POSTMESSAGE EVENT LISTENER
// =============================================================================

/**
 * Handles incoming postMessage events from parent windows.
 *
 * This is the main entry point for external communication. When the converter
 * is embedded in an iframe, parent windows can send messages to request
 * conversions or check the converter's status.
 *
 * SUPPORTED MESSAGE TYPES:
 * ------------------------
 *
 * 1. Convert Request:
 *    {
 *      type: 'convert',
 *      buffer: ArrayBuffer,  // or 'file' for WordPress compatibility
 *      format: 'pdf',        // output format (optional, defaults to 'pdf')
 *      requestId: 'unique-id' // for correlating responses (optional)
 *    }
 *
 * 2. Health Check:
 *    {
 *      type: 'ping',
 *      requestId: 'ping-id'
 *    }
 *    Response: { type: 'pong', ready: boolean, requestId: 'ping-id' }
 */
window.addEventListener('message', async (event) => {
    const { data, origin, source } = event;

    // Security: Validate the message origin against whitelist (if configured)
    if (!isOriginAllowed(origin)) {
        console.warn(`Rejected message from unauthorized origin: ${origin}`);
        return;
    }

    // Determine the origin to use for response messages
    const targetOrigin = getTargetOrigin(origin);

    // Handle conversion request
    // Supports both 'buffer' (standard) and 'file' (WordPress) property names
    if (data && data.type === 'convert') {
        try {
            // Get the document data from either property name
            const inputBuffer = data.buffer || data.file;
            if (!inputBuffer) {
                throw new Error('No document buffer provided');
            }

            // Normalize the input to an ArrayBuffer
            // postMessage can receive ArrayBuffer, TypedArray, or objects with buffer property
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

            // Use provided format or default to PDF
            const outputFormat = data.format || 'pdf';
            // Generate a request ID if not provided
            const requestId = data.requestId || Date.now().toString();

            // Start the conversion
            await convertDocument(arrayBuffer, outputFormat, requestId, source, targetOrigin);

        } catch (error) {
            // Send error response to the requesting window
            source.postMessage({
                type: 'error',
                error: error.message,
                requestId: data.requestId
            }, targetOrigin);
        }
    }

    // Handle ping/health check request
    // Useful for parent windows to check if converter is ready before sending files
    if (data && data.type === 'ping') {
        source.postMessage({
            type: 'pong',
            ready: moduleReady,
            requestId: data.requestId
        }, targetOrigin);
    }
});
