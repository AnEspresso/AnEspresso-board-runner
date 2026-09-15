/* AnEspresso! — Service Worker
 * Purpose: Web Push notifications ONLY.
 *
 * This worker deliberately does NOT cache anything and does NOT
 * intercept network requests. The app manages its own versioning
 * via APP_VERSION; adding caching here would fight that system.
 * Keep this file minimal and stable — it rarely needs to change.
 *
 * Works identically on iOS (16.4+ PWA) and Android (Chrome).
 */

'use strict';

var SW_VERSION = '1';

/* Activate immediately on install — no waiting for old tabs to close. */
self.addEventListener('install', function (event) {
  self.skipWaiting();
});

/* Take control of open pages as soon as the worker activates. */
self.addEventListener('activate', function (event) {
  event.waitUntil(self.clients.claim());
});

/* ----- Push: a notification arrived from the Cloud Function ----- */
self.addEventListener('push', function (event) {
  var data = {};
  try {
    data = event.data ? event.data.json() : {};
  } catch (e) {
    data = {};
  }

  var title = data.title || 'AnEspresso!';
  var options = {
    body: data.body || '',
    // tag groups notifications — a repeat for the same room replaces
    // the old one rather than stacking endlessly.
    tag: data.tag || 'anespresso',
    // renotify: re-alert (sound/vibrate) even when replacing a same-tag
    // notification. Supported on Android; harmless/ignored on iOS.
    renotify: true,
    requireInteraction: false,
    // data is read back in notificationclick below.
    data: data
  };

  event.waitUntil(self.registration.showNotification(title, options));
});

/* ----- Notification tapped: focus the app, or open it ----- */
self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  var data = (event.notification && event.notification.data) || {};

  event.waitUntil(
    self.clients
      .matchAll({ type: 'window', includeUncontrolled: true })
      .then(function (winList) {
        for (var i = 0; i < winList.length; i++) {
          var w = winList[i];
          if (w) {
            try { w.postMessage(data); } catch (e) {}
            if ('focus' in w) return w.focus();
          }
        }
        if (self.clients.openWindow) {
          return self.clients.openWindow('./');
        }
        return null;
      })
  );
});
