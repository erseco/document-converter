/*! coi-serviceworker v0.1.7 - Guido Zuidhof, licensed under MIT */
/*
 * This Service Worker enables Cross-Origin Isolation for environments
 * where server headers cannot be configured (like GitHub Pages).
 *
 * It intercepts all fetch requests and adds the necessary COOP/COEP headers
 * to enable SharedArrayBuffer support required by ZetaJS/LibreOffice WASM.
 *
 * Extended to support POST requests for file upload conversion.
 */

let coepCredentialless = false;

// Storage for POST data (in-memory, cleared after retrieval)
let pendingPostData = null;

if (typeof window === 'undefined') {
    self.addEventListener("install", () => self.skipWaiting());
    self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));

    self.addEventListener("message", (ev) => {
        if (!ev.data) {
            return;
        } else if (ev.data.type === "deregister") {
            self.registration
                .unregister()
                .then(() => {
                    return self.clients.matchAll();
                })
                .then((clients) => {
                    clients.forEach((client) => client.navigate(client.url));
                });
        } else if (ev.data.type === "coepCredentialless") {
            coepCredentialless = ev.data.value;
        } else if (ev.data.type === "getPostData") {
            // Client is requesting the POST data
            const data = pendingPostData;
            pendingPostData = null; // Clear after retrieval
            ev.source.postMessage({ type: "postData", data: data });
        }
    });

    self.addEventListener("fetch", function (event) {
        const r = event.request;

        // Handle POST requests to the main page
        if (r.method === "POST" && (r.url.endsWith("/") || r.url.endsWith("/index.html") || r.url.includes("document-converter"))) {
            event.respondWith(handlePostRequest(r));
            return;
        }

        if (r.cache === "only-if-cached" && r.mode !== "same-origin") {
            return;
        }

        const request =
            coepCredentialless && r.mode === "no-cors"
                ? new Request(r, {
                    credentials: "omit",
                })
                : r;

        event.respondWith(
            fetch(request)
                .then((response) => {
                    if (response.status === 0) {
                        return response;
                    }

                    const newHeaders = new Headers(response.headers);
                    newHeaders.set("Cross-Origin-Embedder-Policy",
                        coepCredentialless ? "credentialless" : "require-corp"
                    );
                    newHeaders.set("Cross-Origin-Opener-Policy", "same-origin");
                    newHeaders.set("Cross-Origin-Resource-Policy", "cross-origin");

                    return new Response(response.body, {
                        status: response.status,
                        statusText: response.statusText,
                        headers: newHeaders,
                    });
                })
                .catch((e) => console.error(e))
        );
    });

    // Handle POST request: extract file and redirect to page with postData flag
    async function handlePostRequest(request) {
        try {
            const formData = await request.formData();
            const file = formData.get("file");
            const format = formData.get("format") || "pdf";
            const download = formData.get("download") === "true";
            const fullscreen = formData.get("fullscreen") === "true";

            if (file && file instanceof File) {
                const arrayBuffer = await file.arrayBuffer();

                // Store the data for retrieval by the page
                pendingPostData = {
                    buffer: arrayBuffer,
                    filename: file.name,
                    format: format,
                    download: download,
                    fullscreen: fullscreen
                };

                // Redirect to the page with a flag indicating POST data is available
                const url = new URL(request.url);
                url.search = "?postData=true";

                // Fetch the actual page and return it
                const pageResponse = await fetch(url.pathname);
                const pageText = await pageResponse.text();

                const newHeaders = new Headers(pageResponse.headers);
                newHeaders.set("Cross-Origin-Embedder-Policy",
                    coepCredentialless ? "credentialless" : "require-corp"
                );
                newHeaders.set("Cross-Origin-Opener-Policy", "same-origin");
                newHeaders.set("Cross-Origin-Resource-Policy", "cross-origin");
                newHeaders.set("Content-Type", "text/html");

                // Inject a script to add postData=true to the URL
                const modifiedHtml = pageText.replace(
                    '</head>',
                    `<script>if(!window.location.search.includes('postData'))history.replaceState(null,'','?postData=true');</script></head>`
                );

                return new Response(modifiedHtml, {
                    status: 200,
                    statusText: "OK",
                    headers: newHeaders,
                });
            } else {
                // No file provided, redirect to normal page
                return Response.redirect(request.url.split('?')[0], 302);
            }
        } catch (e) {
            console.error("Error handling POST:", e);
            return Response.redirect(request.url.split('?')[0], 302);
        }
    }

} else {
    (() => {
        const reloadedBySelf = window.sessionStorage.getItem("coiReloadedBySelf");
        window.sessionStorage.removeItem("coiReloadedBySelf");
        const coepDegrading = (reloadedBySelf === "coepDegrade");

        // You can customize the behavior of this script by setting coi configuration
        // on the global scope before loading.
        const coi = {
            shouldRegister: () => !reloadedBySelf,
            shouldDeregister: () => false,
            coepCredentialless: () => true,
            coepDegrade: () => true,
            doReload: () => {
                window.sessionStorage.setItem("coiReloadedBySelf", 
                    coepDegrading ? "coepDegrade" : "true");
                window.location.reload();
            },
            quiet: false,
            ...window.coi
        };

        const n = navigator;

        if (coi.shouldDeregister()) {
            n.serviceWorker &&
                n.serviceWorker.controller &&
                n.serviceWorker.controller.postMessage({ type: "deregister" });
        }

        // If we're already cross-origin isolated, no need for the service worker
        if (window.crossOriginIsolated) {
            !coi.quiet && console.log("crossOriginIsolated: already enabled");
            return;
        }

        if (!coi.shouldRegister()) {
            !coi.quiet && console.log("coi-serviceworker: will not register");
            return;
        }

        if (!n.serviceWorker) {
            !coi.quiet && console.error("coi-serviceworker: ServiceWorker API not available");
            return;
        }

        n.serviceWorker.register(window.document.currentScript.src).then(
            (registration) => {
                !coi.quiet && console.log("coi-serviceworker: registered", registration.scope);

                registration.addEventListener("updatefound", () => {
                    !coi.quiet && console.log("coi-serviceworker: updatefound, reloading page");
                    coi.doReload();
                });

                // If the service worker is already active, reload immediately
                if (registration.active && !n.serviceWorker.controller) {
                    !coi.quiet && console.log("coi-serviceworker: active, reloading page");
                    coi.doReload();
                }
            },
            (err) => {
                !coi.quiet && console.error("coi-serviceworker: registration failed", err);
            }
        );
    })();
}
