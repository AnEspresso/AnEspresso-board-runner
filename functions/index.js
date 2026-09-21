"use strict";

/**
 * AnEspresso TEST board push — 2nd gen, us-central1.
 * Codebase: board-push. Does not replace production notifyBathroomRequest.
 *
 * Roster: /runnerPushSubscriptions  (never /pushSubscriptions)
 * Test:   /runnerPushTest
 * Log:    /runnerPushLog
 * Urgent: /boards/runner-DATE/sitePrefs/{room}/runnerUrgent
 *         (not `bathroom` — that path is what pings live phones)
 *
 * VAPID public must match the test client. Private is injected at deploy
 * (__VAPID_PRIVATE__ placeholder). Do not commit the real private key.
 * Do not rotate the production bathroom function keys.
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
  "BKMHHLDxgatVMU4rZm8MKLztG1PkGnNkedxpcj8LaWybhN8NhzxnnkY_UfKytxf0QQd5HBHkUsvMZ4eu9lej-Sw";
const VAPID_PRIVATE = process.env.VAPID_PRIVATE_KEY || "__VAPID_PRIVATE__";
const VAPID_SUBJECT = process.env.VAPID_SUBJECT || "mailto:peter@anespresso.com";
const STALE_MS = 4 * 60 * 60 * 1000;
const WIN_LABELS = ["Morning", "Lunch", "Afternoon", "Evening"];
const RTDB = "anespresso-auth-default-rtdb";
const SUBS_ROOT = "runnerPushSubscriptions";
const LOG_ROOT = "runnerPushLog";
const TEST_ROOT = "runnerPushTest";

function configureWebPush() {
  if (!VAPID_PRIVATE || VAPID_PRIVATE.indexOf("__VAPID") === 0) {
    throw new Error("VAPID_PRIVATE_KEY is not set");
  }
  webpush.setVapidDetails(VAPID_SUBJECT, VAPID_PUBLIC, VAPID_PRIVATE);
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
  const snap = await db.ref(SUBS_ROOT).once("value");
  const val = snap.val() || {};
  return Object.keys(val).map((id) => Object.assign({ id }, val[id] || {}));
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
        await db.ref(SUBS_ROOT + "/" + sub.id).remove();
      } catch (e) {}
      return { ok: false, reason: "gone" };
    }
    return { ok: false, reason: String((err && err.message) || err).slice(0, 180) };
  }
}

async function writePushLog(entry) {
  try {
    await db.ref(LOG_ROOT).push(Object.assign({ ts: Date.now(), isolate: "runner" }, entry));
  } catch (e) {}
}

function windowLabel(wIdx) {
  return WIN_LABELS[Number(wIdx)] || "Break";
}

function audienceIncludes(audience, role) {
  if (audience === "runner") return role === "runner";
  if (audience === "runner_breaker") return role === "runner" || role === "breaker";
  return true;
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
    let sent = 0,
      stale = 0,
      failed = 0;
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
    await writePushLog({ audience: "room", room, sent, stale, failed, kind: "breakMarked" });
  }
);

exports.onUrgentBreak = onValueWritten(
  { ref: "/boards/{boardId}/sitePrefs/{room}/runnerUrgent", instance: RTDB },
  async (event) => {
    const after = event.data.after.val();
    const before = event.data.before.val();
    if (!after || before) return;
    const room = event.params.room;
    const boardId = event.params.boardId;
    if (!String(boardId || "").startsWith("runner-")) return;

    configureWebPush();
    const prefSnap = await db.ref("boards/" + boardId + "/sitePrefs/" + room).once("value");
    const pref = prefSnap.val() || {};
    const audience = pref.runnerUrgentAudience || pref.bathroomAudience || "all";
    const payload = {
      title: "Urgent break",
      body: room + " needs a now-break",
      tag: "urgent-" + room,
      kind: "urgent",
      room,
    };
    const subs = await loadSubscriptions();
    let sent = 0,
      stale = 0,
      failed = 0;
    for (const sub of subs) {
      if (sub.notifyUrgent === false) continue;
      const role = sub.role || "crna";
      if (!audienceIncludes(audience, role)) continue;
      if (isStale(sub)) {
        stale += 1;
        continue;
      }
      const r = await sendOne(sub, payload);
      if (r.ok) sent += 1;
      else failed += 1;
    }
    await writePushLog({ audience, room, sent, stale, failed, kind: "urgent" });
  }
);

exports.onPushTest = onValueWritten(
  { ref: "/runnerPushTest/{deviceId}/requestTs", instance: RTDB },
  async (event) => {
    const after = event.data.after.val();
    if (!after) return;
    const deviceId = event.params.deviceId;
    const resultRef = db.ref(TEST_ROOT + "/" + deviceId + "/result");
    try {
      configureWebPush();
    } catch (err) {
      await resultRef.set({ ts: Date.now(), ok: false, error: "missing-vapid" });
      return;
    }
    const snap = await db.ref(SUBS_ROOT + "/" + deviceId).once("value");
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
