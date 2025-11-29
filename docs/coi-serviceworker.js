/*! coi-serviceworker v0.1.7 - Guido Zuidhof, licensed under MIT */
/*
 * This Service Worker enables Cross-Origin Isolation for environments
 * where server headers cannot be configured (like GitHub Pages).
 * 
 * It intercepts all fetch requests and adds the necessary COOP/COEP headers
 * to enable SharedArrayBuffer support required by ZetaJS/LibreOffice WASM.
 */

let coepCredentialless = false;
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
        }
    });

    self.addEventListener("fetch", function (event) {
        const r = event.request;
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
