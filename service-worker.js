// 🚀 Service Worker - 快閃餐車訂購系統
// 版本號（每次更新時遞增）
const CACHE_VERSION = 'v1.0.0';
const CACHE_NAME = `food-truck-cache-${CACHE_VERSION}`;

// 需要快取的靜態資源
const STATIC_CACHE_URLS = [
  './',
  './index.html',
  './config.js',
  './app.js',
  // 可以添加其他靜態資源
];

// 需要快取的 API 請求（運行時快取）
const API_CACHE_NAME = `food-truck-api-${CACHE_VERSION}`;

// 圖片快取
const IMAGE_CACHE_NAME = `food-truck-images-${CACHE_VERSION}`;

// ====== 安裝事件 ======
self.addEventListener('install', (event) => {
  console.log('🚀 Service Worker 安裝中...');
  
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('📦 開始快取靜態資源...');
        return cache.addAll(STATIC_CACHE_URLS);
      })
      .then(() => {
        console.log('✅ 靜態資源快取完成');
        return self.skipWaiting(); // 強制激活新的 Service Worker
      })
      .catch((err) => {
        console.warn('⚠️ 快取靜態資源失敗:', err);
      })
  );
});

// ====== 激活事件 ======
self.addEventListener('activate', (event) => {
  console.log('🔄 Service Worker 激活中...');
  
  event.waitUntil(
    caches.keys()
      .then((cacheNames) => {
        // 刪除舊版本的快取
        return Promise.all(
          cacheNames.map((cacheName) => {
            if (cacheName.startsWith('food-truck-') && 
                cacheName !== CACHE_NAME && 
                cacheName !== API_CACHE_NAME &&
                cacheName !== IMAGE_CACHE_NAME) {
              console.log('🗑️ 刪除舊快取:', cacheName);
              return caches.delete(cacheName);
            }
          })
        );
      })
      .then(() => {
        console.log('✅ Service Worker 激活完成');
        return self.clients.claim(); // 立即接管所有頁面
      })
  );
});

// ====== 攔截請求事件 ======
self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);
  
  // 🔧 跳過 Chrome 擴充功能和外部請求（非同源）
  if (url.protocol === 'chrome-extension:' || 
      url.origin !== location.origin && !url.href.includes('script.google.com')) {
    return;
  }
  
  // 🖼️ 圖片請求：快取優先策略
  if (request.destination === 'image' || url.href.includes('images.unsplash.com')) {
    event.respondWith(
      caches.open(IMAGE_CACHE_NAME)
        .then((cache) => {
          return cache.match(request)
            .then((cachedResponse) => {
              if (cachedResponse) {
                return cachedResponse;
              }
              
              // 如果沒有快取，從網路取得
              return fetch(request).then((networkResponse) => {
                // 🔧 修復：確保在 clone 之前響應未被讀取
                if (networkResponse && 
                    networkResponse.status === 200 && 
                    networkResponse.bodyUsed === false) {
                  const clonedResponse = networkResponse.clone();
                  cache.put(request, clonedResponse);
                }
                return networkResponse;
              });
            });
        })
    );
    return;
  }
  
  // 📡 API 請求：網路優先，失敗時使用快取
  if (url.href.includes('script.google.com')) {
    event.respondWith(
      fetch(request)
        .then((networkResponse) => {
          // 🔧 修復：確保在 clone 之前響應未被讀取
          // API 請求成功，更新快取
          if (networkResponse && 
              networkResponse.status === 200 && 
              networkResponse.bodyUsed === false) {
            const clonedResponse = networkResponse.clone();
            caches.open(API_CACHE_NAME).then((cache) => {
              cache.put(request, clonedResponse);
            });
          }
          return networkResponse;
        })
        .catch(() => {
          // 網路失敗，嘗試從快取取得
          return caches.match(request).then((cachedResponse) => {
            if (cachedResponse) {
              console.log('📦 使用快取 API 回應:', request.url);
              return cachedResponse;
            }
            
            // 如果也沒有快取，返回離線頁面提示
            return new Response(
              JSON.stringify({ 
                ok: false, 
                msg: '網路已斷線，請稍後再試', 
                offline: true 
              }),
              { 
                headers: { 'Content-Type': 'application/json' },
                status: 503
              }
            );
          });
        })
    );
    return;
  }
  
  // 📄 HTML/JS/CSS 請求：快取優先，並在背景更新
  event.respondWith(
    caches.match(request)
      .then((cachedResponse) => {
        // 如果有快取，立即返回，同時在背景更新
        if (cachedResponse) {
          // 在背景更新快取（不阻塞返回）
          fetch(request)
            .then((networkResponse) => {
              // 🔧 修復：確保在 clone 之前響應未被讀取
              if (networkResponse && 
                  networkResponse.status === 200 && 
                  networkResponse.bodyUsed === false) {
                const clonedResponse = networkResponse.clone();
                caches.open(CACHE_NAME).then((cache) => {
                  cache.put(request, clonedResponse);
                });
              }
            })
            .catch((err) => {
              // 背景更新失敗，忽略（使用舊快取）
              console.warn('⚠️ 背景更新快取失敗:', err);
            });
          
          return cachedResponse;
        }
        
        // 如果沒有快取，從網路取得
        return fetch(request).then((networkResponse) => {
          // 🔧 修復：確保在 clone 之前響應未被讀取
          if (networkResponse && 
              networkResponse.status === 200 && 
              networkResponse.bodyUsed === false) {
            const clonedResponse = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(request, clonedResponse);
            });
          }
          return networkResponse;
        });
      })
      .catch(() => {
        // 完全離線時，返回基本的離線頁面
        if (request.destination === 'document') {
          return caches.match('./index.html');
        }
      })
  );
});

// ====== 訊息事件（與主頁面通訊） ======
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  
  // 清除所有快取
  if (event.data && event.data.type === 'CLEAR_CACHE') {
    event.waitUntil(
      caches.keys().then((cacheNames) => {
        return Promise.all(
          cacheNames.filter(name => name.startsWith('food-truck-'))
            .map(name => caches.delete(name))
        );
      }).then(() => {
        console.log('🗑️ 所有 Service Worker 快取已清除');
        // 通知主頁面
        self.clients.matchAll().then((clients) => {
          clients.forEach(client => {
            client.postMessage({ type: 'CACHE_CLEARED' });
          });
        });
      })
    );
  }
});

