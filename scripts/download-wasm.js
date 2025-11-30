#!/usr/bin/env node
/**
 * Script to download ZetaOffice WASM files from the CDN
 * These files are required for LibreOffice WASM to work
 */

import https from 'https';
import fs from 'fs';
import path from 'path';
import zlib from 'zlib';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ZetaOffice CDN base URL
const ZETAOFFICE_CDN = 'https://cdn.zetaoffice.net/zetaoffice_latest';

// Files to download from ZetaOffice CDN
const WASM_FILES = [
    'soffice.js',
    'soffice.wasm',
    'soffice.data'
];

// Target directory for GitHub Pages
const TARGET_DIR = path.join(__dirname, '..', 'docs', 'wasm');

/**
 * Download a file from URL to local path
 */
function downloadFile(url, destPath) {
    return new Promise((resolve, reject) => {
        console.log(`Downloading: ${url}`);

        const file = fs.createWriteStream(destPath);

        const options = {
            headers: {
                'Accept-Encoding': 'br, gzip, deflate'
            }
        };

        https.get(url, options, (response) => {
            // Handle redirects
            if (response.statusCode === 301 || response.statusCode === 302) {
                file.close();
                fs.unlinkSync(destPath);
                const redirectUrl = new URL(response.headers.location, url).href;
                downloadFile(redirectUrl, destPath)
                    .then(resolve)
                    .catch(reject);
                return;
            }

            if (response.statusCode !== 200) {
                file.close();
                fs.unlinkSync(destPath);
                reject(new Error(`Failed to download ${url}: HTTP ${response.statusCode}`));
                return;
            }

            // Handle content encoding (decompress if needed)
            const contentEncoding = response.headers['content-encoding'];
            let stream = response;

            if (contentEncoding === 'br') {
                console.log(`  Content-Encoding: brotli (decompressing)`);
                stream = response.pipe(zlib.createBrotliDecompress());
            } else if (contentEncoding === 'gzip') {
                console.log(`  Content-Encoding: gzip (decompressing)`);
                stream = response.pipe(zlib.createGunzip());
            } else if (contentEncoding === 'deflate') {
                console.log(`  Content-Encoding: deflate (decompressing)`);
                stream = response.pipe(zlib.createInflate());
            }

            const totalSize = parseInt(response.headers['content-length'], 10);
            let downloadedSize = 0;

            response.on('data', (chunk) => {
                downloadedSize += chunk.length;
                if (totalSize) {
                    const percent = ((downloadedSize / totalSize) * 100).toFixed(1);
                    process.stdout.write(`\r  Progress: ${percent}% (${(downloadedSize / 1024 / 1024).toFixed(2)} MB)`);
                }
            });

            stream.pipe(file);

            file.on('finish', () => {
                file.close();
                console.log(`\n  Saved to: ${destPath}`);
                resolve();
            });

            stream.on('error', (err) => {
                file.close();
                if (fs.existsSync(destPath)) {
                    fs.unlinkSync(destPath);
                }
                reject(err);
            });

            file.on('error', (err) => {
                file.close();
                if (fs.existsSync(destPath)) {
                    fs.unlinkSync(destPath);
                }
                reject(err);
            });
        }).on('error', (err) => {
            file.close();
            if (fs.existsSync(destPath)) {
                fs.unlinkSync(destPath);
            }
            reject(err);
        });
    });
}

async function main() {
    console.log('=== ZetaOffice WASM Downloader ===\n');
    console.log(`Source: ${ZETAOFFICE_CDN}`);
    console.log(`Target: ${TARGET_DIR}\n`);
    
    // Create target directory if it doesn't exist
    if (!fs.existsSync(TARGET_DIR)) {
        fs.mkdirSync(TARGET_DIR, { recursive: true });
        console.log(`Created directory: ${TARGET_DIR}\n`);
    }
    
    // Download each file
    for (const file of WASM_FILES) {
        const url = `${ZETAOFFICE_CDN}/${file}`;
        const destPath = path.join(TARGET_DIR, file);
        
        try {
            await downloadFile(url, destPath);
        } catch (error) {
            console.error(`\nError downloading ${file}: ${error.message}`);
            process.exit(1);
        }
    }
    
    console.log('\n=== Download complete! ===');
    
    // Show file sizes
    console.log('\nDownloaded files:');
    for (const file of WASM_FILES) {
        const filePath = path.join(TARGET_DIR, file);
        if (fs.existsSync(filePath)) {
            const stats = fs.statSync(filePath);
            const sizeMB = (stats.size / 1024 / 1024).toFixed(2);
            console.log(`  ${file}: ${sizeMB} MB`);
        }
    }
}

main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
});
