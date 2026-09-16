"use strict";

/**
 * AnEspresso board push — 2nd gen, us-central1.
 *
 * notifyBathroomRequest is already deployed and stays as-is (urgent breaks).
 * These add:
 *   onBreakMarked — CRNA ping when their room is marked
 *   onPushTest    — Developer "end-to-end delivery test"
 *
 * VAPID public key MUST match the client (VAPID_PUBLIC_KEY in index.html).
 * Private key: env VAPID_PRIVATE_KEY (Cloud Run / Functions secret).
 */

const { onValueWritten } = require("firebase-functions/v2/database");
const { setGlobalOptions } = require("firebase-functions/v2");
const admin = require("firebase-admin");
const webpush = require("web-push");

setGlobalOptions({ region: "us-central1", memory: "256MiB", timeoutSeconds: 60 });

if (!admin.apps.length) admin.initializeApp();
const db = admin.database();

const VAPID_PUBLIC =
  process.env.VAPID_PUBLIC_KEY ||
  "BAnJXqd7u6jbAkRqTQ3h3sja9t0FvI0iTQvXVLO_QpvkBZ05JyK01d64YokyO0zW8j2vmi0UrSpQJbKCawlHdXk";
const VAPID_SUBJECT = process.env.VAPID_SUBJECT || "mailto:peter@anespresso.com";
const STALE_MS = 4 * 60 * 60 * 1000;
const WIN_LABELS = ["Morning", "Lunch", "Afternoon", "Evening"];
const RTDB = "anespresso-auth-default-rtdb";

function vapidPrivate() {
  const k = process.env.VAPID_PRIVATE_KEY || "";
  if (!k) throw new Error("VAPID_PRIVATE_KEY is not set");
  return k;
}

function configureWebPush() {
  webpush.setVapidDetails(VAPID_SUBJECT, VAPID_PUBLIC, vapidPrivate());
}

function isStale(sub) {
  const ts = Number(sub && (sub.updatedTs || sub.createdTs)) || 0;
  if (!ts) return true;
  return Date.now() - ts > STALE_MS;
}

function subToPush(sub) {
  if (!sub || !sub.endpoint || !sub.keys || !sub.keys.p256dh || !sub.keys.auth) return null;
  return { endpoint: sub.endpoint, keys: { p256dh: sub.keys.p256dh, auth: sub.keys.auth } };
}

async function loadSubscriptions() {
  const snap = await db.ref("pushSubscriptions").once("value");
  const val = snap.val() || {};
  const out = [];
  Object.keys(val).forEach((id) => {
    out.push(Object.assign({ id }, val[id] || {}));
  });
  return out;
}

async function sendOne(sub, payload) {
  const dest = subToPush(sub);
  if (!dest) return { ok: false, reason: "bad-sub" };
  try {
    await webpush.sendNotification(dest, JSON.stringify(payload), { TTL: 300, urgency: "high" });
    return { ok: true };
  } catch (err) {
    const status = err && (err.statusCode || err.status);
    if (status === 404 || status === 410) {
      try {
        await db.ref("pushSubscriptions/" + sub.id).remove();
      } catch (e) {}
      return { ok: false, reason: "gone" };
    }
    return { ok: false, reason: String((err && err.message) || err).slice(0, 180) };
  }
}

async function writePushLog(entry) {
  try {
    await db.ref("pushLog").push(Object.assign({ ts: Date.now() }, entry));
  } catch (e) {}
}

function windowLabel(wIdx) {
  const i = Number(wIdx);
  return WIN_LABELS[i] || "Break";
}

exports.onBreakMarked = onValueWritten(
  { ref: "/boards/{boardId}/stateEvents/{wIdx}/{catId}/{room}", instance: RTDB },
  async (event) => {
    const after = event.data.after.val();
    const before = event.data.before.val();
    if (!after || !after.done) return;
    if (before && before.done) return;

    const room = event.params.room;
    const wIdx = event.params.wIdx;
    const boardId = event.params.boardId;
    if (!String(boardId || "").startsWith("runner-")) return;

    configureWebPush();
    const payload = {
      title: "Break marked",
      body: windowLabel(wIdx) + " · " + room + " — tap if this is wrong",
      tag: "break-" + room + "-" + wIdx,
      kind: "breakMarked",
      room,
      wIdx: Number(wIdx),
    };

    const subs = await loadSubscriptions();
    let sent = 0;
    let stale = 0;
    let failed = 0;
    for (const sub of subs) {
      if (sub.notifyBreakMarked === false) continue;
      if (String(sub.room || "") !== String(room)) continue;
      if (isStale(sub)) {
        stale += 1;
        continue;
      }
      const r = await sendOne(sub, payload);
      if (r.ok) sent += 1;
      else failed += 1;
    }
    await writePushLog({
      audience: "room",
      room,
      sent,
      stale,
      failed,
      kind: "breakMarked",
    });
  }
);

exports.onPushTest = onValueWritten(
  { ref: "/pushTest/{deviceId}/requestTs", instance: RTDB },
  async (event) => {
    const after = event.data.after.val();
    if (!after) return;
    const deviceId = event.params.deviceId;
    const resultRef = db.ref("pushTest/" + deviceId + "/result");

    try {
      configureWebPush();
    } catch (err) {
      await resultRef.set({ ts: Date.now(), ok: false, error: "missing-vapid" });
      return;
    }

    const snap = await db.ref("pushSubscriptions/" + deviceId).once("value");
    const sub = snap.val();
    if (!sub) {
      await resultRef.set({ ts: Date.now(), ok: false, error: "no-subscription" });
      return;
    }

    const payload = {
      title: "AnEspresso test",
      body: "If you see this, notifications can display on this device.",
      tag: "push-test",
      kind: "pushTest",
    };
    const r = await sendOne(Object.assign({ id: deviceId }, sub), payload);
    await resultRef.set({
      ts: Date.now(),
      ok: !!r.ok,
      error: r.ok ? null : r.reason || "send-failed",
    });
    await writePushLog({
      audience: "test",
      room: "pushTest",
      sent: r.ok ? 1 : 0,
      stale: 0,
      failed: r.ok ? 0 : 1,
      kind: "pushTest",
    });
  }
);
