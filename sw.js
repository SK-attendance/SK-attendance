const CACHE_NAME = "checkin-pwa-v6";
const ASSETS = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", function(event){
  event.waitUntil(caches.open(CACHE_NAME).then(function(cache){ return cache.addAll(ASSETS); }));
  self.skipWaiting();
});

self.addEventListener("activate", function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k !== CACHE_NAME; }).map(function(k){ return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

// 核心修正：HTML頁面用「network-first」，永遠優先攞最新版；
// 只有喺完全無網絡嗰陣，先會退返用cache版本（確保離線仍然打得開）。
// 其他靜態資源（圖片、manifest）維持cache-first，減少流量。
self.addEventListener("fetch", function(event){
  if (event.request.method !== "GET") return;

  var isHTML = event.request.mode === "navigate" ||
               (event.request.headers.get("accept") || "").indexOf("text/html") !== -1;

  if (isHTML) {
    event.respondWith(
      fetch(event.request)
        .then(function(networkResp){
          var clone = networkResp.clone();
          caches.open(CACHE_NAME).then(function(cache){ cache.put(event.request, clone); });
          return networkResp;
        })
        .catch(function(){ return caches.match(event.request); })
    );
    return;
  }

  event.respondWith(caches.match(event.request).then(function(cached){ return cached || fetch(event.request); }));
});
