# Document Converter

A static microservice for Vercel that uses ZetaJS (LibreOffice WASM) to convert documents in the browser.

## Features

- **Client-side document conversion** using LibreOffice WASM via ZetaJS
- **Iframe-compatible** with proper CORS and security headers for SharedArrayBuffer support
- **Multiple format support**: PDF, DOCX, XLSX, PPTX, ODT, ODS, ODP, HTML, TXT, RTF, PNG, JPG
- **Origin whitelist** optional security feature to restrict which domains can use the converter

## Deployment

Deploy to Vercel:

```bash
vercel
```

Or click the deploy button:

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/erseco/document-converter)

## Usage

### Embedding in an iframe

```html
<iframe id="converter" src="https://your-vercel-deployment.vercel.app"></iframe>
```

### With origin whitelist (optional security)

To restrict which domains can communicate with the converter, add the `origins` query parameter:

```html
<iframe id="converter" src="https://your-vercel-deployment.vercel.app?origins=https://example.com,https://another.com"></iframe>
```

### Sending documents for conversion

```javascript
const iframe = document.getElementById('converter');

// Wait for the converter to be ready
window.addEventListener('message', (event) => {
    if (event.data.type === 'ready') {
        console.log('Converter is ready!');
    }
    
    if (event.data.type === 'result') {
        // Handle the converted blob
        const blob = event.data.blob;
        const url = URL.createObjectURL(blob);
        // Download or display the result
    }
    
    if (event.data.type === 'error') {
        console.error('Conversion error:', event.data.error);
    }
});

// Convert a document
async function convertDocument(file, format = 'pdf') {
    const arrayBuffer = await file.arrayBuffer();
    
    iframe.contentWindow.postMessage({
        type: 'convert',
        buffer: arrayBuffer,
        format: format,  // 'pdf', 'docx', 'xlsx', etc.
        requestId: Date.now()
    }, '*');
}
```

## Security Headers

The `vercel.json` configuration includes required headers for SharedArrayBuffer support:

- `Cross-Origin-Opener-Policy: same-origin`
- `Cross-Origin-Embedder-Policy: require-corp`
- `Content-Security-Policy` with `frame-ancestors *` to allow embedding

## Supported Formats

### Input formats
Any format supported by LibreOffice (DOC, DOCX, ODT, XLS, XLSX, ODS, PPT, PPTX, ODP, etc.)

### Output formats
- **PDF** - Portable Document Format
- **DOCX** - Microsoft Word
- **XLSX** - Microsoft Excel
- **PPTX** - Microsoft PowerPoint
- **ODT** - OpenDocument Text
- **ODS** - OpenDocument Spreadsheet
- **ODP** - OpenDocument Presentation
- **HTML** - Web page
- **TXT** - Plain text
- **RTF** - Rich Text Format
- **PNG** - Image
- **JPG** - Image

## License

MIT