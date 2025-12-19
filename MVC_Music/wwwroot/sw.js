(function () {
    'use strict';

    var version = 'v1.0::CacheFirstSafe';
    var offlineUrl = "offline.html";

    function updateStaticCache() {
        return caches.open(version)
            .then(function (cache) {
                return cache.addAll([
                    offlineUrl
                ]);
            });
    }

    function addToCache(request, response) {
        if (!response.ok && response.type !== 'opaque')
            return;

        var copy = response.clone();
        caches.open(version)
            .then(function (cache) {
                cache.put(request, copy);
            });
    }

    self.addEventListener('install', function (event) {
        event.waitUntil(updateStaticCache());
    });

    self.addEventListener('activate', function (event) {
        event.waitUntil(
            caches.keys()
                .then(function (keys) {
                    return Promise.all(keys
                        .filter(function (key) {
                            return key.indexOf(version) !== 0;
                        })
                        .map(function (key) {
                            return caches.delete(key);
                        })
                    );
                })
        );
    });


    self.addEventListener('fetch', function (event) {
        var request = event.request;

        if (request.method !== 'GET' || request.url.match(/\/browserLink/ig)) {
            event.respondWith(
                fetch(request)
                    .catch(function () {
                        return caches.match(offlineUrl);
                    })
            );
            return;
        }

        if (request.url.match(/(\?|&)v=/ig)) {
            event.respondWith(
                caches.match(request)
                    .then(function (response) {
                        return response || fetch(request)
                            .then(function (response) {
                                addToCache(request, response);
                                return response || caches.match(offlineUrl);
                            })
                            .catch(function () {
                                return caches.match(offlineUrl);
                            });
                    })
            );

            return;
        }

        event.respondWith(
            fetch(request)
                .then(function (response) {
                    addToCache(request, response);
                    return response;
                })
                .catch(function () {
                    return caches.match(request)
                        .then(function (response) {
                            return response || caches.match(offlineUrl);
                        })
                        .catch(function () {
                            return caches.match(offlineUrl);
                        });
                })
        );
    });

})();