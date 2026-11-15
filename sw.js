/**
 * 🔔 365메디스타약국 실시간 웹 푸시 & 백그라운드 알림 서비스 워커 (Service Worker)
 */

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

// 📩 메인 화면 앱에서 보낸 실시간 배지 세팅 메시지 수신 (백그라운드 동기화 지원)
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SET_BADGE') {
    const count = Number(event.data.count) || 0;
    try {
      if ('setAppBadge' in self.navigator) {
        if (count > 0) {
          self.navigator.setAppBadge(count).catch(() => {});
        } else if ('clearAppBadge' in self.navigator) {
          self.navigator.clearAppBadge().catch(() => {});
        }
      }
    } catch(err) {}
  }
});

// 📩 백그라운드 푸시 알림 수신 이벤트
self.addEventListener('push', (event) => {
  let data = {
    title: '📢 365메디스타약국 알림',
    body: '새로운 공지, 업무일지, 소모품 또는 약품 위치 변동사항이 도착했습니다.',
    icon: 'logo.jpg',
    badge: 'logo.jpg',
    url: './',
    tag: 'mds-notification',
    unreadCount: 1
  };

  try {
    if (event.data) {
      const payload = event.data.json();
      data = { ...data, ...payload };
    }
  } catch (e) {
    if (event.data) {
      data.body = event.data.text();
    }
  }

  const options = {
    body: data.body,
    icon: 'logo.jpg',
    badge: 'logo.jpg',
    tag: 'mds-notification-channel-v3',
    channelId: 'mds_push_channel_v3',
    renotify: true,
    silent: false,
    data: { url: data.url || './' },
    vibrate: [300, 100, 300, 100, 400, 100, 500],
    requireInteraction: true
  };

  const badgePromise = (async () => {
    try {
      if (typeof self.navigator !== 'undefined' && 'setAppBadge' in self.navigator) {
        const unreadNum = Number(data.unreadCount) || 1;
        await self.navigator.setAppBadge(unreadNum);
      }
    } catch(e) {}
  })();

  event.waitUntil(
    Promise.all([
      self.registration.showNotification(data.title, options),
      badgePromise
    ])
  );
});

// 🖱️ 스마트폰 상단 알림 배너 터치/클릭 이벤트
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const targetUrl = (event.notification.data && event.notification.data.url) || './';

  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      for (let i = 0; i < clientList.length; i++) {
        const client = clientList[i];
        if ('focus' in client) {
          return client.focus();
        }
      }
      if (clients.openWindow) {
        return clients.openWindow(targetUrl);
      }
    })
  );
});
