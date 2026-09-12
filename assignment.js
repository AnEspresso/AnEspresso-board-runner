/* AnEspresso assignment-sheet parser + late board. Unshipped. */
(function (g) {
  "use strict";

  var SHIFT_LATE = { M: 1, S: 1 };
  var SHIFT_DINNER = { Q: 1, W: 1, E: 1 };
  var ROOM_INDEX = null;

  function buildRoomIndex() {
    ROOM_INDEX = {};
    if (typeof CATEGORIES === "undefined") return;
    CATEGORIES.forEach(function (c) {
      (c.rooms || []).forEach(function (r) {
        ROOM_INDEX[String(r).toUpperCase()] = { cat: c.id, room: r };
      });
    });
  }

  function u16(v, i) { return v.getUint16(i, true); }
  function u32(v, i) { return v.getUint32(i, true); }

  async function inflateRaw(data) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("This phone cannot unpack Excel. Use Files in a current Safari.");
    }
    var ds = new DecompressionStream("deflate-raw");
    var stream = new Blob([data]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(buf) {
    var u = new Uint8Array(buf);
    var v = new DataView(buf);
    var eocd = -1;
    var min = Math.max(0, u.length - 22 - 65557);
    for (var i = u.length - 22; i >= min; i--) {
      if (u32(v, i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Not an Excel workbook (.xlsx)");
    var n = u16(v, eocd + 10);
    var cdOff = u32(v, eocd + 16);
    var p = cdOff;
    var files = {};
    for (var k = 0; k < n; k++) {
      if (u32(v, p) !== 0x02014b50) break;
      var method = u16(v, p + 10);
      var comp = u32(v, p + 20);
      var nameLen = u16(v, p + 28);
      var extra = u16(v, p + 30);
      var comment = u16(v, p + 32);
      var localOff = u32(v, p + 42);
      var name = new TextDecoder().decode(u.slice(p + 46, p + 46 + nameLen));
      var localNameLen = u16(v, localOff + 26);
      var localExtra = u16(v, localOff + 28);
      var dataStart = localOff + 30 + localNameLen + localExtra;
      files[name] = { method: method, data: u.slice(dataStart, dataStart + comp) };
      p += 46 + nameLen + extra + comment;
    }
    var out = {};
    var names = Object.keys(files);
    for (var j = 0; j < names.length; j++) {
      var f = files[names[j]];
      if (f.method === 0) out[names[j]] = f.data;
      else if (f.method === 8) out[names[j]] = await inflateRaw(f.data);
    }
    return out;
  }

  function xmlText(bytes) {
    return new TextDecoder("utf-8").decode(bytes);
  }

  function loadSharedStrings(xml) {
    var ss = [];
    if (!xml) return ss;
    var doc = new DOMParser().parseFromString(xml, "text/xml");
    var nodes = doc.getElementsByTagName("si");
    for (var i = 0; i < nodes.length; i++) {
      var ts = nodes[i].getElementsByTagName("t");
      var s = "";
      for (var j = 0; j < ts.length; j++) s += ts[j].textContent || "";
      ss.push(s);
    }
    return ss;
  }

  function loadCells(sheetXml, ss) {
    var doc = new DOMParser().parseFromString(sheetXml, "text/xml");
    var cells = {};
    var nodes = doc.getElementsByTagName("c");
    for (var i = 0; i < nodes.length; i++) {
      var c = nodes[i];
      var ref = c.getAttribute("r");
      if (!ref) continue;
      var t = c.getAttribute("t");
      var v = c.getElementsByTagName("v")[0];
      var isEl = c.getElementsByTagName("is")[0];
      var val = "";
      if (t === "s" && v && v.textContent) val = ss[parseInt(v.textContent, 10)] || "";
      else if (t === "inlineStr" && isEl) {
        var ts = isEl.getElementsByTagName("t");
        for (var j = 0; j < ts.length; j++) val += ts[j].textContent || "";
      } else if (v) val = v.textContent || "";
      if (val && String(val).trim()) cells[ref] = String(val).trim();
    }
    return cells;
  }

  function cell(cells, col, row) {
    return (cells[col + row] || "").trim();
  }

  function excelDate(val) {
    val = String(val || "").trim();
    var m = val.match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
    if (m) {
      var mo = parseInt(m[1], 10), d = parseInt(m[2], 10), y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      return y + "-" + mo + "-" + d;
    }
    if (/^\d+(\.0+)?$/.test(val)) {
      var serial = parseInt(val, 10);
      var dt = new Date(Date.UTC(1899, 11, 30));
      dt.setUTCDate(dt.getUTCDate() + serial);
      return dt.getUTCFullYear() + "-" + (dt.getUTCMonth() + 1) + "-" + dt.getUTCDate();
    }
    return null;
  }

  function extractFirstCase(s) {
    var m = String(s || "").match(/[@＠]\s*(?:0?(\d{1,2})[:.](\d{2})|(\d{3,4}))/);
    if (!m) return "";
    if (m[3]) {
      var d = m[3];
      if (d.length === 3) return d.charAt(0) + ":" + d.slice(1);
      return d.slice(0, 2) + ":" + d.slice(2);
    }
    return parseInt(m[1], 10) + ":" + m[2];
  }

  function cleanName(s) {
    s = String(s || "").replace(/\*/g, " ");
    s = s.replace(/\s*\+.*$/, "");
    s = s.replace(/\s*\(.*\)$/, "");
    s = s.replace(/\s*REQ\s*$/i, "");
    s = s.replace(/[@＠]\s*(?:0?\d{1,2}[:.]\d{2}|\d{3,4})/g, " ");
    s = s.replace(/[\^]/g, "");
    s = s.replace(/['`’]+/g, "");
    return s.replace(/\s+/g, " ").replace(/^[\s-]+|[\s-]+$/g, "");
  }

  function lastName(name) {
    var p = cleanName(name).split(/\s+/);
    var last = p.length ? p[p.length - 1] : "";
    if (last.length > 12 && last.indexOf("-") >= 0) {
      var bits = last.split("-");
      last = bits[bits.length - 1] || last;
    }
    return last;
  }

  function firstName(name) {
    var p = cleanName(name).split(/\s+/).filter(Boolean);
    if (p.length < 2) return "";
    return p[0];
  }

  function lastNameCounts() {
    var counts = {};
    function add(n) {
      var ln = lastName(n).toLowerCase();
      if (!ln) return;
      counts[ln] = (counts[ln] || 0) + 1;
    }
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = g.roomStaff[cat][room];
        if (rec && rec.name) add(rec.name);
      });
    });
    (g.onDeck || []).forEach(function (p) { if (p && p.name) add(p.name); });
    return counts;
  }

  function chipName(name) {
    var last = lastName(name);
    if (!last) return "";
    if ((lastNameCounts()[last.toLowerCase()] || 0) < 2) return last;
    var first = firstName(name);
    if (!first) return last;
    return first.charAt(0).toUpperCase() + ". " + last;
  }

  function isTimeLabel(s) {
    s = String(s || "").trim();
    if (!s) return false;
    if (/^\d{1,2}(:\d{2})?\s*[-–to]+\s*\d/i.test(s)) return true;
    if (/^\d{1,2}\s*-\s*\d{1,2}\s*[ap]/i.test(s)) return true;
    if (/^\d{1,2}[ap]m?\s*-\s*\d{1,2}[ap]/i.test(s)) return true;
    if (/^\d{3,4}\s*-\s*\d{3,4}/.test(s)) return true;
    if (/^(11-7a|3-11p|7-3p|7-5p|7-7p|3p-11p|11p-7a)$/i.test(s.replace(/\s/g, ""))) return true;
    return false;
  }

  function isJunkStaff(name, shift) {
    var n = String(name || "").trim();
    if (!n) return true;
    if (isTimeLabel(n)) return true;
    if (!/[A-Za-z]{3,}/.test(n)) return true;
    if (/^\d{1,2}-\d{1,2}/.test(n)) return true;
    if (/^(11-7a|3-11p)$/i.test(n.replace(/\s/g, ""))) return true;
    return false;
  }

  function normalizeShift(sh) {
    return String(sh || "").split("/").map(function (p) {
      p = p.trim();
      if (p === "d") return "d";
      if (p === "e") return "e";
      if (p === "t") return "t";
      if (p.toLowerCase() === "s") return "S";
      return p.toUpperCase();
    }).join("/");
  }

  function coreShift(shift) {
    return String(shift || "").replace(/^o\//i, "").replace(/^\*/, "").split("/")[0];
  }

  function lastShiftPart(shift) {
    var s = String(shift || "").replace(/^o\//i, "");
    var parts = s.split("/").map(function (p) {
      p = String(p || "").trim();
      return { early: p.charAt(0) === "*", letter: p.replace(/^\*/, "") };
    }).filter(function (p) { return p.letter; });
    return parts.length ? parts[parts.length - 1] : { early: false, letter: "" };
  }

  function isEarlyShift(shift) {
    var s = String(shift || "").replace(/^o\//i, "");
    return s.charAt(0) === "*";
  }

  function fmtShift(letter, early, orient) {
    if (!letter || letter === "Dr") return letter || "";
    return (orient ? "o/" : "") + (early ? "*" : "") + letter;
  }

  function breakKind(shift) {
    var day = coreShift(shift);
    if (!day || day === "Dr" || /^\d/.test(day)) return "none";
    if (SHIFT_LATE[day]) return "late";
    if (SHIFT_DINNER[day]) return "dinner";
    return "none";
  }

  function parseClockToken(tok) {
    tok = String(tok || "").trim().toLowerCase().replace(/\s+/g, "");
    var ap = /p/.test(tok) ? "p" : (/a/.test(tok) ? "a" : "");
    tok = tok.replace(/[ap]m?/g, "");
    var h = 0, m = 0;
    var mm = tok.match(/^(\d{1,2}):(\d{2})$/);
    if (mm) { h = parseInt(mm[1], 10); m = parseInt(mm[2], 10); }
    else if (/^\d{3,4}$/.test(tok)) {
      if (tok.length === 3) { h = parseInt(tok.charAt(0), 10); m = parseInt(tok.slice(1), 10); }
      else { h = parseInt(tok.slice(0, 2), 10); m = parseInt(tok.slice(2), 10); }
    } else h = parseInt(tok, 10) || 0;
    if (ap === "p" && h < 12) h += 12;
    if (ap === "a" && h === 12) h = 0;
    return { h: h, m: m || 0, ap: ap };
  }

  function compactRangeLabel(a, b) {
    function bit(x) {
      return String(x || "").replace(/\s+/g, "").replace(/am/ig, "a").replace(/pm/ig, "p").toLowerCase();
    }
    return bit(a) + "-" + bit(b);
  }

  function parseTimeRange(s) {
    var m = String(s || "").trim().match(/^(\d{1,2}(?::\d{2})?\s*[ap]m?)\s*[-–to]+\s*(\d{1,2}(?::\d{2})?\s*[ap]?m?)\s*$/i);
    if (!m) return null;
    var start = parseClockToken(m[1]);
    var end = parseClockToken(m[2]);
    if (!end.ap && start.ap === "a" && end.h > 0 && end.h <= 12) {
      if (end.h !== 12) end.h += 12;
    }
    return { startH: start.h, endH: end.h, label: compactRangeLabel(m[1], m[2]) };
  }

  function parseStaff(raw) {
    var s = String(raw || "").replace(/\s+/g, " ").trim();
    if (!s) return null;
    var student = /\^/.test(s);
    if (/^closed$/i.test(s)) return { closed: true, shift: "", name: "", kind: "none", early: false, orient: false, student: false };
    if (/^#ref/i.test(s)) return null;
    if (/rotate/i.test(s) && !/[A-Za-z]{4,}\s+[A-Za-z]{3,}/.test(s)) return null;

    var rest = s;
    var orient = false;
    var early = false;
    for (var i = 0; i < 6; i++) {
      if (/^o\//i.test(rest)) { orient = true; rest = rest.replace(/^o\//i, "").trim(); continue; }
      if (/^\*/.test(rest)) { early = true; rest = rest.replace(/^\*+/, "").trim(); continue; }
      break;
    }
    if (/[DdMmSsQqWwEeNnTt]\*/.test(s)) early = true;

    function rec(shift, name, kind) {
      return { closed: false, shift: shift || "", name: cleanName(name), kind: kind || "none", early: early, orient: orient, student: student, firstCase: extractFirstCase(s) };
    }

    var m = rest.match(/^(\d{1,2}(?::\d{2})?\s*[ap]m?)\s*[-–to]+\s*(\d{1,2}(?::\d{2})?\s*[ap]?m?)\s+(.+)$/i);
    if (m && /[A-Za-z]{2,}/.test(m[3])) {
      var tr = parseTimeRange(m[1] + "-" + m[2]);
      return rec(tr ? tr.label : compactRangeLabel(m[1], m[2]), m[3], "none");
    }
    m = rest.match(/^(Dr\.?)\s+(.+)$/i);
    if (m) return rec("Dr", m[2], "none");
    m = rest.match(/^([DdMmSsQqWwEeNnTt](?:\/[DdMmSsQqWwEeNnTt])?)\*?\s*(.+)$/);
    if (m && /[A-Za-z]/.test(m[2])) {
      var sh = normalizeShift(m[1]);
      return rec(fmtShift(sh, early, orient), m[2], breakKind(sh));
    }
    m = rest.match(/^([DdMmSsQqWwEeNnTt])\*?(?=[A-Z])([A-Z].+)$/);
    if (m) {
      sh = normalizeShift(m[1]);
      return rec(fmtShift(sh, early, orient), m[2], breakKind(sh));
    }
    return rec("", rest, "none");
  }

  function skipRoomLabel(s) {
    var t = String(s || "").replace(/\s+/g, " ").trim();
    if (!t) return true;
    if (/^#\d/.test(t)) return true;
    return /^(NORTH TOWER|SOUTH TOWER|OFFSITE|FBC|M\/N|STE\.?\s*100|WBF|ENDO A|ENDO B|RESOURCE STAFF|POC TESTING|CCS POC|2N POC|2S POC|STE POC|OB POC|CHECK HEMACUE|TRAUMA RM|SHIFT|CRNA|ASSIGNMENTS?|MIDNIGHT|LATE STAY|MN CALL|HEART CALL|NT BREAKERS|ST BREAKERS|POS|CV|3N|2N|ST 1|ST 2|STE 1|STE 2|CCS|ENDO|EP|OB)$/i.test(t);
  }

  function normalizeRoom(raw) {
    if (!ROOM_INDEX) buildRoomIndex();
    var s = String(raw || "").replace(/\s+/g, " ").trim();
    if (!s) return null;
    var ste = s.match(/^STE\.?\s*(10[1-9])$/i);
    if (ste) {
      var steRoom = "OR " + ste[1];
      if (ROOM_INDEX[steRoom]) return ROOM_INDEX[steRoom];
      return { cat: "s100", room: steRoom };
    }
    if (/^OB(@|\s|$)/i.test(s) && !/POC/i.test(s)) {
      /* OB is a real assignment on weekend sheets */
    } else if (skipRoomLabel(s)) return null;
    var su = s.toUpperCase();
    su = su.replace(/^STE\.?\s*/, "OR ");
    su = su.replace(/^ENDO\s*/, "ENDO ");
    su = su.replace("OB1", "OB 1").replace("OB2", "OB 2");
    su = su.replace("1ST MRI", "MRI 1ST");
    su = su.replace(/U\/S.*/, "US");
    if (su === "PET") su = "PET SCAN";
    su = su.replace(/@\s*3\s*P.*/i, "").trim();
    su = su.replace(/^IR\s*(\d+).*/, "IR $1");
    su = su.replace(/^BMBX?\d*.*/, "BMB");
    su = su.replace(/^CT\b.*/, "CT");
    su = su.replace("OR39/40", "OR 38-39").replace("OR 39/40", "OR 38-39");
    su = su.replace(/^MRI IC\b.*/, "MRI IC 1");
    if (su === "OB") su = "OB 1";
    var num = su.match(/^(\d{1,3})(?:\s*\(.*\))?$/);
    if (num) {
      var n = parseInt(num[1], 10);
      if ((n >= 1 && n <= 36) || (n >= 51 && n <= 66) || (n >= 71 && n <= 76) || n >= 101) su = "OR " + n;
    }
    var em = su.match(/^ENDO\s*(\d)$/i);
    if (em) su = "ENDO " + em[1];
    var epm = su.match(/^EP\s*(\d)$/);
    if (epm) su = "EP " + epm[1];
    if (ROOM_INDEX[su]) return ROOM_INDEX[su];
    var su2 = su.replace(/\s*[\(@].*$/, "").trim();
    if (ROOM_INDEX[su2]) return ROOM_INDEX[su2];
    return null;
  }

  function nameKey(n) {
    return String(n || "").toLowerCase().replace(/[^a-z]/g, "");
  }

  function shiftEndHour(shift) {
    var tr = parseTimeRange(shift);
    if (tr) return tr.endH;
    var ENDH = { D: 15, d: 17, M: 18, S: 19, Q: 21, W: 23, E: 23, e: 31, N: 31, t: 31 };
    var part = lastShiftPart(shift);
    var end = ENDH[part.letter];
    if (end == null) return 99;
    if (part.early) end -= 1;
    return end;
  }

  function shiftStartHour(shift) {
    var tr = parseTimeRange(shift);
    if (tr) return tr.startH;
    var STARTH = { D: 7, d: 7, M: 6, S: 7, Q: 7, W: 7, E: 15, e: 21, N: 23, t: 19 };
    var letter = coreShift(shift);
    var start = STARTH[letter];
    if (start == null) return 7;
    if (isEarlyShift(shift)) start -= 1;
    return start;
  }

  function isLeavingSoon(shift) {
    var end = shiftEndHour(shift);
    var h = hospitalHour();
    return h >= end - 1 && h < end + 2;
  }

  function intendedRoomOf(p) {
    if (!p) return "";
    if (p.role === "wbf" || p.role === "offsite") return "";
    if (p.intended) return p.intended;
    if (p.role === "extra") return p.lastRoom || "";
    return "";
  }

  function hospitalHour() {
    if (typeof g._hourOverride === "number") return g._hourOverride;
    try {
      if (typeof hospitalNow === "function") return hospitalNow().getHours();
      return new Date(new Date().toLocaleString("en-US", { timeZone: "America/Detroit" })).getHours();
    } catch (e) {
      return new Date().getHours();
    }
  }

  function stillInHouse(shift) {
    return hospitalHour() < shiftEndHour(shift);
  }

  function breakWindowFor(kind) {
    return kind === "dinner" ? 3 : 2;
  }

  function breakGiven(kind, catId, room) {
    if (!catId || !room || typeof state === "undefined") return false;
    var w = breakWindowFor(kind);
    try {
      return !!(state[w] && state[w][catId] && state[w][catId][room]);
    } catch (e) {
      return false;
    }
  }

  function collectBreakQueue(kind) {
    var items = [];
    var seen = {};
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = g.roomStaff[cat][room];
        if (!rec || rec.closed || !rec.name) return;
        if ((rec.kind || breakKind(rec.shift)) !== kind) return;
        var deleted = typeof catEditState !== "undefined" && catEditState[cat] && catEditState[cat].deletedRooms.has(room);
        if (deleted) return;
        if (!stillInHouse(rec.shift)) return;
        var k = nameKey(rec.name);
        if (seen[k]) return;
        seen[k] = 1;
        items.push({
          name: rec.name,
          shift: rec.shift,
          kind: kind,
          cat: cat,
          room: room,
          given: breakGiven(kind, cat, room),
          end: shiftEndHour(rec.shift),
          role: ""
        });
      });
    });
    (g.onDeck || []).forEach(function (p) {
      if (!p || !p.name) return;
      var knd = p.kind || breakKind(p.shift);
      if (knd !== kind) return;
      if (!stillInHouse(p.shift)) return;
      var k = nameKey(p.name);
      if (seen[k]) return;
      seen[k] = 1;
      items.push({
        name: p.name,
        shift: p.shift,
        kind: kind,
        cat: "",
        room: p.lastRoom || "",
        given: false,
        end: shiftEndHour(p.shift),
        role: p.role || "deck"
      });
    });
    items.sort(function (a, b) {
      if (a.given !== b.given) return a.given ? 1 : -1;
      if (a.end !== b.end) return a.end - b.end;
      return chipName(a.name).localeCompare(chipName(b.name));
    });
    return items;
  }

  function parseWeekday(cells) {
    var date = excelDate(cell(cells, "H", 3));
    var ntR = parseStaff(cell(cells, "E", 2));
    var stR = parseStaff(cell(cells, "E", 3));
    var rooms = [];
    var unmatched = [];
    var closed = [];
    var onDeck = [];
    var filled = {};
    var mriIcN = 0;
    function addDeck(st, lastRoom, role) {
      if (!st || !st.name || st.closed) return;
      if (isJunkStaff(st.name, st.shift)) return;
      if (/^\d{3,4}\s*-\s*\d/.test(st.name) || /^2200/.test(st.name)) return;
      onDeck.push({
        name: st.name, shift: st.shift || "", kind: st.kind || "none",
        lastRoom: lastRoom || "", role: role || "",
        intended: (role === "extra" || role === "offsite") ? (lastRoom || "") : "",
        student: !!st.student, firstCase: st.firstCase || ""
      });
    }
    function addPair(roomRaw, staffRaw, src) {
      roomRaw = (roomRaw || "").trim();
      staffRaw = (staffRaw || "").trim();
      if (!roomRaw) return;
      if (skipRoomLabel(roomRaw) && !/^STE\.?\s*10[1-9]$/i.test(roomRaw)) return;
      if (!staffRaw) return;
      var locRoom = roomRaw;
      if (/^MRI IC\b/i.test(roomRaw)) {
        mriIcN += 1;
        locRoom = mriIcN === 1 ? "MRI IC 1" : "MRI IC 2";
      }
      var loc = normalizeRoom(locRoom);
      var st = parseStaff(staffRaw);
      if (!loc) {
        if (st && st.name) addDeck(st, roomRaw, "unplaced");
        unmatched.push({ room: roomRaw, staff: staffRaw });
        return;
      }
      if (!st) return;
      if (isJunkStaff(st.name, st.shift)) return;
      var key = loc.cat + "|" + loc.room;
      var isSte = src === "ste" || /^STE\.?\s*10[1-9]$/i.test(roomRaw);
      if (filled[key] && !st.closed && !isSte) {
        addDeck(st, loc.room, src === "cd" ? "offsite" : "extra");
        return;
      }
      var rec = {
        cat: loc.cat, room: loc.room, shift: st.shift || "", name: st.name || "",
        closed: !!st.closed, kind: st.kind || "none", early: !!st.early, orient: !!st.orient,
        student: !!st.student, firstCase: st.firstCase || ""
      };
      if (filled[key]) {
        rooms = rooms.filter(function (x) { return !(x.cat === loc.cat && x.room === loc.room); });
      }
      filled[key] = rec;
      rooms.push(rec);
      if (rec.closed) closed.push({ cat: loc.cat, room: loc.room });
    }
    for (var r = 4; r <= 44; r++) {
      addPair(cell(cells, "A", r), cell(cells, "B", r), "ab");
      addPair(cell(cells, "E", r), cell(cells, "F", r), "ef");
    }
    var inWbf = false;
    for (r = 4; r <= 44; r++) {
      var cRaw = cell(cells, "C", r);
      var dRaw = cell(cells, "D", r);
      if (/^WBF$/i.test(cRaw)) { inWbf = true; continue; }
      if (inWbf) {
        var wst = parseStaff(dRaw);
        if (wst && wst.name && !wst.closed && !isJunkStaff(wst.name, wst.shift)) {
          addDeck(wst, "WBF", "wbf");
        }
        continue;
      }
      addPair(cRaw, dRaw, "cd");
    }
    for (r = 4; r <= 44; r++) {
      var steLab = cell(cells, "E", r);
      if (/^STE\.?\s*10[1-9]$/i.test(steLab)) addPair(steLab, cell(cells, "F", r), "ste");
    }
    function addLabelled(raw, lastRoom, role) {
      var st = parseStaff(raw);
      if (st && st.name) addDeck(st, lastRoom, role);
    }
    for (r = 5; r <= 43; r++) {
      var g = cell(cells, "G", r);
      var h = cell(cells, "H", r);
      if (/^MIDNIGHT/i.test(g)) {
        if (!isTimeLabel(h)) addLabelled(h, g, "midnight");
      }
      else if (/^#\d/.test(g)) { /* late stay handled below */ }
      else if (/^(CV|CCS|3N|2N|ENDO|ST 1|ST 2|STE 1|STE 2)$/i.test(g)) addLabelled(h, g, "breaker");
      else if (/^(MN CALL|HEART CALL)$/i.test(g)) {
        var who = h && !/^\d{3,4}/.test(h) ? h : cell(cells, "G", r + 1);
        if (who && !/^(MN CALL|HEART CALL|NT BREAKERS|ST BREAKERS|LATE STAY)/i.test(who)) {
          addLabelled(who, g, "call");
        }
      }
    }
    var lateStays = [];
    var lsWave = null;
    for (r = 5; r <= 43; r++) {
      g = cell(cells, "G", r);
      h = cell(cells, "H", r);
      var win = (g + " " + h).match(/(15|17|19|21)30\s*[-–]\s*(17|19|21|23)30/);
      if (win) lsWave = parseInt(win[1], 10);
      if (/late\s*stay/i.test(g) || /late\s*stay/i.test(h)) continue;
      var num = String(g || "").match(/^#\s*(\d+)/);
      if (num && lsWave != null) {
        var lst = parseStaff(h);
        if (lst && lst.name && !isJunkStaff(lst.name, lst.shift)) {
          lateStays.push({
            n: parseInt(num[1], 10),
            name: lst.name,
            shift: lst.shift || "",
            wave: lsWave,
            role: "latestay"
          });
        }
      }
    }
    var placed = {};
    rooms.forEach(function (x) {
      if (x.name && !x.closed) {
        placed[nameKey(x.name)] = 1;
      }
    });
    onDeck = onDeck.filter(function (p) {
      return p.name && !placed[nameKey(p.name)];
    });
    var seen = {};
    onDeck = onDeck.filter(function (p) {
      var k = nameKey(p.name);
      if (seen[k]) return false;
      seen[k] = 1;
      return true;
    });
    var pos = {};
    for (r = 36; r <= 41; r++) {
      var lab = cell(cells, "G", r);
      var val = cell(cells, "H", r);
      var blob = lab + " " + val;
      function take(key, re) {
        var mm = blob.match(re);
        if (mm) pos[key] = parseInt(mm[1], 10);
      }
      take("1530-1730", /1530-1730\s*=?\s*(\d+)/);
      take("1730-1930", /1730-1930\s*=?\s*(\d+)/);
      take("1930-2130", /1930-2130\s*=?\s*(\d+)/);
      take("2130-2300", /2130-2300\s*=?\s*(\d+)/);
      take("2300-0700", /2300-0700\s*=\s*(\d+)/);
      if (/1530/.test(lab) && !pos["1530-1730"]) {
        var mm = val.match(/(\d+)/);
        if (mm) pos["1530-1730"] = parseInt(mm[1], 10);
      }
    }
    return {
      kind: "weekday", date: date,
      runners: { nt: ntR && ntR.name, st: stR && stR.name },
      rooms: rooms, closed: closed, unmatched: unmatched, pos: pos, people: [], onDeck: onDeck,
      lateStays: lateStays
    };
  }

  function bumpDate(d) {
    var p = String(d || "").split(/[-/]/).map(function (x) { return parseInt(x, 10); });
    if (p.length < 3 || !p[0]) return d;
    var dt = new Date(p[0], p[1] - 1, p[2]);
    dt.setDate(dt.getDate() + 1);
    return dt.getFullYear() + "-" + (dt.getMonth() + 1) + "-" + dt.getDate();
  }

  function dayKey(d) {
    var p = String(d || "").split(/[-/]/);
    if (p.length < 3) return "";
    return parseInt(p[0], 10) + "-" + parseInt(p[1], 10) + "-" + parseInt(p[2], 10);
  }

  function compactShiftCell(a) {
    a = String(a || "").trim();
    var m = a.match(/^(\d{1,2})(?::(\d{2}))?\s*[-–]\s*(\d{1,2})(?::(\d{2}))?\s*([ap])m?$/i);
    if (m) return m[1] + m[4] + "-" + m[3] + m[4];
    return a;
  }

  function looksLikeShiftCell(a) {
    a = compactShiftCell(a);
    if (!a) return false;
    if (/^(?:o\/)?\*?[DdMmSsQqWwEeNnTt](?:\/[DdMmSsQqWwEeNnTt])?$/i.test(a)) return true;
    if (parseTimeRange(a)) return true;
    return false;
  }

  function parseWeekend(cells) {
    var people = [];
    var currentDate = null, currentTower = "";
    var lastShift = "";
    var sawNT = 0;
    var eveningTotals = [];
    var dayTotals = [];
    var posSection = "";
    var pendingCall = false;
    var extraF = [];
    var r;

    function already(date, name) {
      var k = dayKey(date) + "|" + nameKey(name);
      for (var i = 0; i < people.length; i++) {
        if (dayKey(people[i].date) + "|" + nameKey(people[i].name) === k) return true;
      }
      return false;
    }

    function addPerson(opts) {
      if (!opts || !opts.name || isJunkStaff(opts.name, opts.shift)) return;
      if (already(opts.date, opts.name)) return;
      var assign = String(opts.assign || "").trim();
      var at3p = /@\s*3\s*p/i.test(assign);
      var assignClean = assign.replace(/@\s*3\s*p.*/i, "").trim();
      var loc = assignClean && !/^BR$/i.test(assignClean) ? normalizeRoom(assignClean) : null;
      var role = opts.role || "";
      if (/^BR$/i.test(assignClean) || /^BR$/i.test(assign)) role = "breaker";
      if (at3p && !role) role = "float";
      people.push({
        date: opts.date, tower: opts.tower || currentTower,
        shift: opts.shift || "", name: cleanName(opts.name),
        assign: assign, cat: loc && loc.cat, room: loc && loc.room,
        closed: false, kind: opts.kind || breakKind(opts.shift),
        early: !!opts.early, orient: !!opts.orient, student: !!opts.student,
        role: role, at3p: at3p, intended: at3p ? "OB" : (opts.intended || "")
      });
    }

    for (r = 1; r <= 80; r++) {
      var a = cell(cells, "A", r), b = cell(cells, "B", r), c = cell(cells, "C", r);
      var e = cell(cells, "E", r), f = cell(cells, "F", r);

      if (/North Tower/i.test(b) || /North Tower/i.test(a)) {
        var dt = excelDate(f) || excelDate(e) || excelDate(c);
        sawNT += 1;
        if (dt) currentDate = dt;
        else if (sawNT > 1 && currentDate) currentDate = bumpDate(currentDate);
        currentTower = "NT";
        lastShift = "";
        pendingCall = false;
        continue;
      }
      if (/South Tower/i.test(b) || /South Tower/i.test(a)) {
        currentTower = "ST";
        lastShift = "";
        pendingCall = false;
        continue;
      }
      if (/^Days$/i.test(e)) posSection = "day";
      if (/^Evenings$/i.test(e)) posSection = "eve";
      if (/^total:/i.test(e)) {
        var tm = e.match(/(\d+)/);
        if (tm) {
          var tot = parseInt(tm[1], 10);
          if (posSection === "eve") eveningTotals.push({ date: currentDate, n: tot });
          else dayTotals.push({ date: currentDate, n: tot });
        }
      }

      if (/CRNA Call/i.test(e) || /CRNA Call/i.test(b)) { pendingCall = true; continue; }
      if (pendingCall) {
        var callNm = "";
        if (e && /[A-Za-z]{3,}/.test(e) && !/liver|resid|call|note/i.test(e)) callNm = e;
        else if (b && /[A-Za-z]{3,}/.test(b) && !/CRNA|tower|shift/i.test(b)) callNm = b;
        if (callNm) {
          addPerson({ date: currentDate, name: callNm, shift: "", assign: "call", role: "call" });
          pendingCall = false;
        }
      }
      var callM = String(e || "").match(/^(6\s*-?\s*3|2\s*-?\s*11|10p\s*-?\s*0?7:?30)\s*#?\s*\d*\s*:?\s*(.+)$/i);
      if (callM && /[A-Za-z]{3,}/.test(callM[2]) && !/^[A-Z]{2,}\s+\d/.test(callM[2])) {
        var cs = /10p/i.test(callM[1]) ? "N" : (/2/.test(callM[1]) ? "E" : "D");
        addPerson({ date: currentDate, name: callM[2], shift: cs, assign: callM[1].replace(/\s+/g, ""), role: "call" });
      }

      if (f && looksLikeShiftCell(f.split(/\s+/)[0]) && /[A-Za-z]{3,}/.test(f)) {
        extraF.push({ date: currentDate, raw: f });
      }

      if (/^shift$/i.test(a) || /^CRNA$/i.test(b)) continue;
      if (/^(check hemacue|trauma rm)/i.test(a)) continue;

      var shiftCell = compactShiftCell(a);
      var hasShift = looksLikeShiftCell(shiftCell);
      if (!hasShift && lastShift && b && /[A-Za-z]{3,}/.test(b) && !skipRoomLabel(b) && !/^CRNA$/i.test(b) && !/tower/i.test(b)) {
        shiftCell = lastShift;
        hasShift = true;
      }
      if (!hasShift || !b) continue;
      if (looksLikeShiftCell(compactShiftCell(a))) lastShift = compactShiftCell(a);

      var raw = looksLikeShiftCell(compactShiftCell(a)) ? (compactShiftCell(a) + " " + b) : (shiftCell + " " + b);
      var stA = parseStaff(raw);
      if (!stA || !stA.name) continue;
      addPerson({
        date: currentDate, tower: currentTower, name: stA.name, shift: stA.shift,
        assign: c, kind: stA.kind, early: stA.early, orient: stA.orient, student: stA.student
      });
    }
    extraF.forEach(function (x) {
      var stF = parseStaff(x.raw);
      if (!stF || !stF.name || !stF.shift) return;
      var fRole = (coreShift(stF.shift) === "N" || lastShiftPart(stF.shift).letter === "N") ? "midnight" : "float";
      addPerson({ date: x.date, name: stF.name, shift: stF.shift, assign: "", role: fRole, kind: stF.kind });
    });

    var today = "";
    try { if (typeof todayStr === "function") today = dayKey(todayStr()); } catch (err) {}
    var dayPeople = people.filter(function (p) {
      if (!today) return true;
      var dk = dayKey(p.date);
      return !dk || dk === today;
    });

    var rooms = [];
    var unmatched = [];
    var onDeck = [];
    var used = {};
    var obQueue = [];

    function take(p, extra) {
      var k = nameKey(p.name);
      if (used[k]) return;
      used[k] = 1;
      var rec = {
        name: p.name, shift: p.shift || "", kind: p.kind || breakKind(p.shift),
        lastRoom: (extra && extra.lastRoom) || p.assign || p.role || "",
        role: (extra && extra.role) || p.role || "float",
        intended: (extra && extra.intended) || p.intended || "",
        student: !!p.student
      };
      onDeck.push(rec);
    }

    dayPeople.forEach(function (p) {
      if (!p.name || used[nameKey(p.name)]) return;
      if (p.role === "breaker") { take(p, { role: "breaker", lastRoom: "BR" }); return; }
      if (p.role === "call") { take(p, { role: "call", lastRoom: p.assign || "call" }); return; }
      if (p.role === "midnight") { take(p, { role: "midnight", lastRoom: "night" }); return; }
      if (p.at3p) { take(p, { role: "float", lastRoom: "OB @3p", intended: "OB" }); return; }
      if (p.room && /^OB/i.test(p.room)) { obQueue.push(p); return; }
      if (p.cat && p.room) {
        rooms.push({ cat: p.cat, room: p.room, shift: p.shift, name: p.name, closed: false, kind: p.kind, student: !!p.student });
        used[nameKey(p.name)] = 1;
        return;
      }
      if (p.assign && !/^BR$/i.test(p.assign) && p.role !== "call") unmatched.push({ room: p.assign, staff: p.name });
      take(p, { role: p.role || "float" });
    });
    var obSlots = ["OB 1", "OB 2"];
    obQueue.forEach(function (p, i) {
      if (used[nameKey(p.name)]) return;
      if (i < obSlots.length) {
        rooms.push({ cat: "fbc", room: obSlots[i], shift: p.shift, name: p.name, closed: false, kind: p.kind, student: !!p.student });
        used[nameKey(p.name)] = 1;
      } else {
        take(p, { role: "float", lastRoom: "OB", intended: "OB" });
      }
    });

    var dates = [];
    people.forEach(function (p) {
      var dk = dayKey(p.date);
      if (dk && dates.indexOf(dk) < 0) dates.push(dk);
    });
    var todayPos = dayTotals.filter(function (x) { return !today || dayKey(x.date) === today; })[0];
    var todayEve = eveningTotals.filter(function (x) { return !today || dayKey(x.date) === today; })[0];
    return {
      kind: "weekend", date: today || dates[0] || null, dates: dates, runners: {},
      rooms: rooms, closed: [], unmatched: unmatched, people: dayPeople,
      pos: { day: todayPos ? todayPos.n : null, evening: todayEve ? todayEve.n : 4, "2300-0700": 4 },
      onDeck: onDeck
    };
  }

  async function parseAssignmentWorkbook(arrayBuffer, fileName) {
    buildRoomIndex();
    var zip = await unzip(arrayBuffer);
    var ss = loadSharedStrings(zip["xl/sharedStrings.xml"] ? xmlText(zip["xl/sharedStrings.xml"]) : "");
    var sheet = zip["xl/worksheets/sheet1.xml"];
    if (!sheet) throw new Error("No sheet1 in workbook");
    var cells = loadCells(xmlText(sheet), ss);
    var blob = Object.keys(cells).map(function (k) { return cells[k]; }).join(" ").toUpperCase();
    var out;
    if (blob.indexOf("CRNA SCHEDULE") >= 0 && blob.indexOf("NORTH TOWER") >= 0) out = parseWeekday(cells);
    else out = parseWeekend(cells);
    out.file = fileName || "";
    out.onDeck = out.onDeck || [];
    out.late = (out.rooms || []).filter(function (x) { return x.kind === "late" && !x.closed && x.name; });
    out.dinner = (out.rooms || []).filter(function (x) { return x.kind === "dinner" && !x.closed && x.name; });
    if (out.kind === "weekend" && out.people) {
      out.late = out.people.filter(function (x) { return x.kind === "late"; });
      out.dinner = out.people.filter(function (x) { return x.kind === "dinner"; });
    }
    return out;
  }

  function shiftPillHtml(shift) {
    var sh = String(shift || "");
    if (!sh) return "";
    var extra = " ";
    if (sh === "Dr") extra += "dr";
    else if (parseTimeRange(sh)) extra += "time";
    else if (sh.length > 2) extra += "wide";
    return '<span class="shift-pill' + extra + '">' + sh + "</span>";
  }

  function deckTag(p) {
    if (p.role === "wbf") return "WBF";
    if (p.role === "freed" && p.lastRoom) return "last " + p.lastRoom;
    if (p.role === "breaker") return "breaker";
    if (p.role === "call") return "call";
    if (p.role === "midnight") return "night";
    return "";
  }

  function studentMark(on) {
    return on ? '<span class="staff-stu" title="Student in room">^</span>' : "";
  }

  function hasStudent(rec) {
    if (!rec) return false;
    if (rec.student) return true;
    return /\^/.test(rec.name || "");
  }

  function staffChipHtml(catId, room) {
    var rec = occupantOf(catId, room) || (g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room]);
    if ((!rec || !rec.name) && typeof staffRecForRoom === "function") rec = staffRecForRoom(room, catId);
    if (!rec || rec.closed || !rec.name) return "";
    var nm = chipName(rec.name);
    var size = nm.length > 10 ? " tiny" : nm.length > 7 ? " long" : "";
    var gone = !stillInHouse(rec.shift);
    var start = rec.firstCase ? '<span class="staff-last">' + rec.firstCase + "</span>" : "";
    return '<span class="staff-line">' + shiftPillHtml(rec.shift) + '<span class="staff-name' + size + (gone ? " gone" : "") + '">' + nm + "</span>" + studentMark(hasStudent(rec)) + start + "</span>";
  }

  function roomBreakKind(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || rec.closed || !rec.name) return "none";
    return rec.kind || breakKind(rec.shift);
  }

  function tintRoomBtn(btn, catId, room) {
    if (!btn) return;
    btn.classList.remove("kind-late", "kind-dinner");
    var win = typeof currentWindow === "number" ? currentWindow : 0;
    if (win < 2) return;
    var k = roomBreakKind(catId, room);
    if (k === "late" || k === "dinner") btn.classList.add("kind-" + k);
  }

  g.tintRoomBtn = tintRoomBtn;
  g.roomBreakKind = roomBreakKind;

  function applyAssignmentResult(res) {
    try {
      if (!g.roomStaff) g.roomStaff = {};
      var listed = {};
      (res.rooms || []).forEach(function (r) {
        if (!r.cat || !r.room) return;
        listed[r.cat + "|" + r.room] = r;
      });
      var now = Date.now();
      var openN = 0;
      var closedN = 0;
      function roomSet(es) {
        if (!es) return null;
        if (!es.deletedRooms || typeof es.deletedRooms.add !== "function") {
          es.deletedRooms = new Set(Array.isArray(es.deletedRooms) ? es.deletedRooms : []);
        }
        if (!es.deletedEvents || typeof es.deletedEvents !== "object") es.deletedEvents = {};
        return es;
      }
      if (typeof CATEGORIES !== "undefined") {
        CATEGORIES.forEach(function (c) {
          if (!g.roomStaff[c.id]) g.roomStaff[c.id] = {};
          (c.rooms || []).forEach(function (room) {
            var r = listed[c.id + "|" + room];
            var named = r && r.name && !r.closed && !isJunkStaff(r.name, r.shift);
            var es = (typeof catEditState !== "undefined") ? roomSet(catEditState[c.id]) : null;
            if (named) {
              g.roomStaff[c.id][room] = {
                name: r.name, shift: r.shift, kind: r.kind, closed: false,
                early: !!r.early, orient: !!r.orient, student: !!r.student,
                firstCase: r.firstCase || ""
              };
              if (es) {
                es.deletedRooms.delete(room);
                es.deletedEvents[room] = { deleted: false, ts: now };
              }
              openN++;
            } else {
              g.roomStaff[c.id][room] = { name: "", shift: "", kind: "none", closed: true };
              if (es) {
                es.deletedRooms.add(room);
                es.deletedEvents[room] = { deleted: true, ts: now };
              }
              closedN++;
            }
          });
        });
      } else {
        (res.rooms || []).forEach(function (r) {
          if (!r.cat || !r.room) return;
          if (!g.roomStaff[r.cat]) g.roomStaff[r.cat] = {};
          g.roomStaff[r.cat][r.room] = { name: cleanName(r.name), shift: r.shift, kind: r.kind, closed: r.closed, student: !!r.student };
        });
      }
      g.onDeck = (res.onDeck || []).map(function (p) {
        return {
          name: cleanName(p.name), shift: p.shift || "", kind: p.kind || "none",
          lastRoom: p.lastRoom || "", role: p.role || "",
          intended: p.role === "wbf" ? "" : (p.intended || ((p.role === "extra" || p.role === "offsite") ? (p.lastRoom || "") : "")),
          student: !!p.student, firstCase: p.firstCase || ""
        };
      }).filter(function (p) { return p.name && !isJunkStaff(p.name, p.shift); });
      g.lateStays = (res.lateStays || []).map(function (p) {
        return { n: p.n, name: cleanName(p.name), shift: p.shift || "", wave: p.wave, role: "latestay" };
      });
      g.assignmentMeta = {
        date: res.date, kind: res.kind, file: res.fileName || "", pos: res.pos,
        lateN: (res.late || []).length, dinnerN: (res.dinner || []).length,
        roomsN: openN, closedN: closedN,
        unmatchedN: (res.unmatched || []).length, deckN: g.onDeck.length, appliedAt: now,
        lateStays: g.lateStays
      };
      g.reliefPlan = {};
      try { if (typeof dayShiftPOCCalculated !== "undefined") dayShiftPOCCalculated = true; } catch (e) {}
      try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
      try {
        if (typeof buildBoard === "function" && !document.getElementById("card-nt")) buildBoard();
        else if (typeof refreshBoard === "function") refreshBoard();
      } catch (e) {}
      try { if (typeof updateActivePOCCounter === "function") updateActivePOCCounter(); } catch (e) {}
      try { renderLateBoardBar(); } catch (e) { console.warn(e); }
      try { renderOnDeck(); } catch (e) { console.warn(e); }
      try { renderShiftChange(); } catch (e) { console.warn(e); }
      try { paintWkndStaff(); } catch (e) { console.warn(e); }
      try { if (typeof renderDeskPeople === "function") renderDeskPeople(); } catch (e) {}
      try {
        if (typeof showToast === "function") {
          showToast(openN + " rooms staffed · " + closedN + " closed");
        }
      } catch (e) {}
    } catch (e) {
      console.warn(e);
      try { if (typeof showToast === "function") showToast("Apply failed: " + ((e && e.message) || "error")); } catch (x) {}
    }
  }

  function focusQueueRoom(catId, room) {
    if (!catId || !room) return;
    var card = document.getElementById("card-" + catId);
    var btn = document.querySelector("#rooms-" + catId + " .room-btn[data-room=\"" + room + "\"]");
    var target = btn || card;
    if (!target) return;
    try { target.scrollIntoView({ behavior: "smooth", block: "center" }); } catch (e) {}
    if (btn) {
      btn.classList.add("queue-flash");
      setTimeout(function () { btn.classList.remove("queue-flash"); }, 1200);
    }
  }

  function queueRowHtml(p, n) {
    var tag = p.given ? "given" : (p.role === "breaker" ? "breaker" : (p.room || ""));
    var last = tag ? '<span class="staff-last">' + tag + "</span>" : "";
    var num = p.given ? "" : '<span class="breakq-n">' + n + "</span>";
    return '<button type="button" class="breakq-item' + (p.given ? " given" : "") + '"' +
      (p.cat ? ' data-qcat="' + p.cat + '"' : "") +
      (p.room ? ' data-qroom="' + p.room + '"' : "") + ">" +
      num + shiftPillHtml(p.shift) +
      '<span class="staff-name">' + chipName(p.name) + "</span>" + last +
      "</button>";
  }

  function queueBlockHtml(label, items) {
    var due = items.filter(function (p) { return !p.given; });
    var givenN = items.length - due.length;
    if (!items.length) {
      return '<div class="late-board-row"><strong>' + label + "</strong> none in house</div>";
    }
    var rows = due.map(function (p, i) { return queueRowHtml(p, i + 1); }).join("");
    var givenNote = givenN ? '<span class="breakq-given-n">' + givenN + " given</span>" : "";
    return '<div class="late-board-row"><strong>' + label + "</strong> " + due.length + " due " + givenNote + "</div>" +
      '<div class="breakq-list">' + rows + "</div>";
  }

  function runnerDrawerOpen() {
    var d = document.getElementById("runner-drawer");
    return !!(d && d.classList.contains("open"));
  }

  function setRunnerDrawer(open) {
    var d = document.getElementById("runner-drawer");
    if (!d) return;
    d.classList.toggle("open", !!open);
    d.classList.toggle("closed", !open);
    try { localStorage.setItem("anespresso_runner_drawer", open ? "open" : "closed"); } catch (e) {}
  }

  function updateRunnerDrawerPeek(text) {
    var el = document.getElementById("runner-drawer-peek");
    if (el) el.textContent = text || "Runner tools";
  }

  function bindRunnerDrawerSwipe(handle) {
    if (!handle || handle._swipeOn) return;
    handle._swipeOn = true;
    var startY = 0, lastY = 0, tracking = false;
    function yOf(e) {
      if (e.touches && e.touches[0]) return e.touches[0].clientY;
      if (e.changedTouches && e.changedTouches[0]) return e.changedTouches[0].clientY;
      return e.clientY;
    }
    handle.addEventListener("pointerdown", function (e) {
      tracking = true;
      startY = lastY = yOf(e);
      try { handle.setPointerCapture(e.pointerId); } catch (x) {}
    });
    handle.addEventListener("pointermove", function (e) {
      if (!tracking) return;
      lastY = yOf(e);
    });
    function end(e) {
      if (!tracking) return;
      tracking = false;
      var dy = (e ? yOf(e) : lastY) - startY;
      if (dy > 36) setRunnerDrawer(false);
      else if (dy < -36) setRunnerDrawer(true);
      else setRunnerDrawer(!runnerDrawerOpen());
    }
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", function () { tracking = false; });
  }

  function ensureRunnerDrawer() {
    var d = document.getElementById("runner-drawer");
    if (!d) {
      d = document.createElement("div");
      d.id = "runner-drawer";
      d.className = "runner-drawer desk-only";
      d.innerHTML =
        '<button type="button" class="board-desk-btn" id="board-desk-btn">' +
        '<span class="desk-title">Board desk</span>' +
        '<span class="desk-peek" id="runner-drawer-peek">Rooms · people · relief</span></button>';
      var tb = document.getElementById("time-poc-bar");
      if (tb && tb.parentNode) tb.parentNode.insertBefore(d, tb.nextSibling);
      else {
        var app = document.getElementById("app");
        if (app) app.insertBefore(d, app.firstChild);
      }
    } else if (!d.querySelector("#board-desk-btn")) {
      d.className = "runner-drawer desk-only";
      d.innerHTML =
        '<button type="button" class="board-desk-btn" id="board-desk-btn">' +
        '<span class="desk-title">Board desk</span>' +
        '<span class="desk-peek" id="runner-drawer-peek">Rooms · people · relief</span></button>';
    }
    var btn = document.getElementById("board-desk-btn");
    if (btn && !btn._deskBound) {
      btn._deskBound = true;
      btn.onclick = function () { openBoardDesk(); };
    }
    var show = typeof currentRole === "undefined" || currentRole === "runner";
    d.style.display = show ? "block" : "none";
    ["evening-runner-bar", "late-board-bar", "shift-change-bar", "edit-active-bar"].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.style.display = "none";
    });
    return d;
  }

  function renderLateBoardBar() {
    ensureRunnerDrawer();
    var bar = document.getElementById("late-board-bar");
    if (bar) bar.style.display = "none";
    var hasAssign = !!(g.assignmentMeta || (g.roomStaff && Object.keys(g.roomStaff).length));
    var hour = hospitalHour();
    var afternoon = (typeof currentWindow === "number" && currentWindow >= 2) || hour >= 15;
    var evening = typeof currentWindow === "number" && currentWindow === 3;
    var weekend = typeof isWeekend === "function" && isWeekend();
    var show = hasAssign && (typeof currentRole === "undefined" || currentRole === "runner");
    if (bar) bar.style.display = "none";
    if (!show) {
      updateRunnerDrawerPeek("Rooms · people · relief");
      return;
    }

    var posOpen = 0;
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = g.roomStaff[cat][room];
        if (!rec || rec.closed) return;
        var dr = typeof catEditState !== "undefined" && catEditState[cat] && catEditState[cat].deletedRooms;
        if (dr && typeof dr.has === "function" && dr.has(room)) return;
        posOpen++;
      });
    });
    var posNote = "";
    if (g.assignmentMeta && g.assignmentMeta.pos) {
      if (weekend && g.assignmentMeta.pos.evening != null) posNote = "Sheet evening POS " + g.assignmentMeta.pos.evening;
      else if (g.assignmentMeta.pos["1530-1730"] != null) posNote = "Sheet 15:30 POS " + g.assignmentMeta.pos["1530-1730"];
    }

    var late = collectBreakQueue("late");
    var dinner = collectBreakQueue("dinner");
    var title = afternoon || weekend ? "Breaks due" : "Late board";
    var body;
    if (afternoon || evening || weekend) {
      body = queueBlockHtml("Late (M+S)", late) + queueBlockHtml("Dinner (Q+W+E)", dinner);
    } else {
      body =
        '<div class="late-board-row"><strong>Late (M+S)</strong> ' + late.length +
        (late.length ? " — " + late.slice(0, 10).map(function (p) { return chipName(p.name); }).join(", ") : "") + "</div>" +
        '<div class="late-board-row"><strong>Dinner (Q+W+E)</strong> ' + dinner.length +
        (dinner.length ? " — " + dinner.slice(0, 10).map(function (p) { return chipName(p.name); }).join(", ") : "") + "</div>";
    }
    var lateDue = late.filter(function (p) { return !p.given; }).length;
    var dinnerDue = dinner.filter(function (p) { return !p.given; }).length;
    var waves = collectAllLeavingWaves();
    var peekOut = waves.slice(0, 2).map(function (b) {
      return b.rooms.length + " at " + waveClock(b.wave);
    }).join(" · ");
    updateRunnerDrawerPeek(lateDue + " late · " + dinnerDue + " dinner" + (peekOut ? " · " + peekOut : ""));
  }

  function isDeckOut(p) {
    if (!p) return false;
    var end = shiftEndHour(p.shift);
    if (end >= 99) return false;
    return !stillInHouse(p.shift);
  }

  function sortByShiftLen(arr) {
    arr.sort(function (a, b) {
      var da = shiftEndHour(a && a.shift);
      var db = shiftEndHour(b && b.shift);
      if (da !== db) return da - db;
      return String(chipName(a && a.name)).localeCompare(String(chipName(b && b.name)));
    });
    return arr;
  }

  function renderOnDeck() {
    var board = document.getElementById("board");
    if (!board) return;
    var card = document.getElementById("card-ondeck");
    var list = (g.onDeck || []).filter(function (p) {
      return p && p.name && !isJunkStaff(p.name, p.shift);
    });
    if (!list.length) {
      if (card) card.remove();
      return;
    }
    if (!card) {
      card = document.createElement("div");
      card.className = "cat-card ondeck-card";
      card.id = "card-ondeck";
      board.insertBefore(card, board.firstChild);
    }
    var here = [];
    var gone = [];
    list.forEach(function (p) {
      if (isDeckOut(p)) gone.push(p);
      else here.push(p);
    });
    sortByShiftLen(here);
    sortByShiftLen(gone);
    var sel = g.selectedDeck || null;
    if (sel && gone.some(function (p) { return nameKey(p.name) === sel; })) g.deckOutOpen = true;
    function chipHtml(p, out) {
      var key = nameKey(p.name);
      var tag = deckTag(p);
      var last = tag ? '<span class="staff-last">' + tag + "</span>" : "";
      var on = sel === key ? " selected" : "";
      return '<button type="button" class="ondeck-chip' + (out ? " out" : "") + on + '" data-deck="' + key + '">' +
        shiftPillHtml(p.shift) + '<span class="staff-name">' + chipName(p.name) + "</span>" + studentMark(hasStudent(p)) + last +
        (out ? '<span class="staff-last">out</span>' : "") +
        "</button>";
    }
    var hereChips = here.map(function (p) { return chipHtml(p, false); }).join("");
    var goneChips = gone.map(function (p) { return chipHtml(p, true); }).join("");
    var open = !!g.deckOutOpen;
    var outBtn = gone.length
      ? '<button type="button" class="ondeck-out-toggle" id="ondeck-out-toggle">' +
        (open ? "Hide out · " : "Out · ") + gone.length + (open ? " ▴" : " ▾") + "</button>" +
        (open ? '<div class="ondeck-grid ondeck-out-grid">' + goneChips + "</div>" : "")
      : "";
    var hint = g.fillTarget
      ? '<div class="ondeck-hint">Tap who goes in ' + g.fillTarget.room + "</div>"
      : (sel
        ? '<div class="ondeck-hint">Tap a room to place them · tap the chip again to cancel</div>'
        : '<div class="ondeck-hint">Tap someone, then tap a room</div>');
    card.innerHTML =
      '<div class="cat-header-row"><div class="cup-indicator">☕</div><div class="cat-info">' +
      '<div class="cat-name">On deck</div>' +
      '<div class="cat-full-name">Not in a room · tap to place</div></div>' +
      '<div class="cat-progress-label"><div class="cat-pct">' + here.length + '</div>' +
      '<div class="cat-count">available</div></div></div>' +
      hint +
      '<div class="ondeck-grid">' + (hereChips || '<span class="roster-empty">Nobody still on shift</span>') + "</div>" +
      outBtn;
    card.querySelectorAll("[data-deck]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        var k = el.getAttribute("data-deck");
        if (g.fillTarget && g.fillTarget.cat && g.fillTarget.room) {
          g.selectedDeck = k;
          var tgt = g.fillTarget;
          g.fillTarget = null;
          assignSelectedTo(tgt.cat, tgt.room);
          return;
        }
        g.selectedDeck = g.selectedDeck === k ? null : k;
        document.body.classList.toggle("assigning", !!g.selectedDeck);
        renderOnDeck();
        renderShiftChange();
      };
    });
    var tog = document.getElementById("ondeck-out-toggle");
    if (tog) tog.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      g.deckOutOpen = !g.deckOutOpen;
      renderOnDeck();
    };
    document.body.classList.toggle("assigning", !!g.selectedDeck);
  }

  function pushOnDeckFromRoom(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || !rec.name) return;
    g.onDeck = g.onDeck || [];
    var k = nameKey(rec.name);
    g.onDeck = g.onDeck.filter(function (p) { return nameKey(p.name) !== k; });
    g.onDeck.unshift({ name: rec.name, shift: rec.shift || "", kind: rec.kind || "none", lastRoom: room, role: "freed", student: !!rec.student });
    delete g.roomStaff[catId][room];
    renderOnDeck();
  }

  function assignSelectedTo(catId, room) {
    var key = g.selectedDeck;
    if (!key || !catId || !room) return false;
    var list = g.onDeck || [];
    var idx = -1;
    for (var i = 0; i < list.length; i++) {
      if (nameKey(list[i].name) === key) { idx = i; break; }
    }
    if (idx < 0) { g.selectedDeck = null; return false; }
    var person = list[idx];
    g.roomStaff = g.roomStaff || {};
    g.roomStaff[catId] = g.roomStaff[catId] || {};
    var occ = g.roomStaff[catId][room];
    list.splice(idx, 1);
    if (occ && occ.name) {
      list.unshift({ name: occ.name, shift: occ.shift || "", kind: occ.kind || "none", lastRoom: room, role: "freed", student: !!occ.student });
    }
    g.roomStaff[catId][room] = { name: person.name, shift: person.shift || "", kind: person.kind || "none", closed: false, student: !!person.student };
    if (typeof catEditState !== "undefined" && catEditState[catId]) {
      catEditState[catId].deletedRooms.delete(room);
    }
    g.selectedDeck = null;
    document.body.classList.remove("assigning");
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    renderOnDeck();
    try { if (typeof showToast === "function") showToast(chipName(person.name) + " → " + room); } catch (e) {}
    return true;
  }

  function roomIsActive(catId, room) {
    if (typeof catEditState === "undefined" || !catEditState[catId]) return true;
    return !catEditState[catId].deletedRooms.has(room);
  }

  function findIncomingFor(room, occupant, used) {
    var occEnd = occupant && occupant.shift ? shiftEndHour(occupant.shift) : hospitalHour();
    var hour = hospitalHour();
    var best = null;
    var bestScore = 99;
    (g.onDeck || []).forEach(function (p, idx) {
      if (!p || !p.name) return;
      var k = nameKey(p.name);
      if (used[k]) return;
      if (!stillInHouse(p.shift) && shiftEndHour(p.shift) < hour) return;
      var intended = intendedRoomOf(p);
      if (intended !== room) return;
      var start = shiftStartHour(p.shift);
      var score = Math.abs(start - occEnd);
      if (start < occEnd - 2) score += 8;
      if (score < bestScore) {
        bestScore = score;
        best = { p: p, idx: idx };
      }
    });
    return best;
  }

  function occupantEndHours() {
    var seen = {};
    var hours = [];
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        if (!roomIsActive(cat, room)) return;
        var rec = g.roomStaff[cat][room];
        if (!rec || rec.closed || !rec.name) return;
        if (coreShift(rec.shift) === "Dr") return;
        var e = shiftEndHour(rec.shift);
        if (e == null || e >= 99) return;
        if (seen[e]) return;
        seen[e] = 1;
        hours.push(e);
      });
    });
    hours.sort(function (a, b) { return a - b; });
    return hours;
  }

  function nextReliefWave() {
    var hour = hospitalHour();
    var hours = occupantEndHours();
    var i;
    for (i = 0; i < hours.length; i++) {
      if (hours[i] >= hour) return hours[i];
    }
    for (i = 0; i < hours.length; i++) {
      if (hours[i] >= hour - 1) return hours[i];
    }
    return hours.length ? hours[hours.length - 1] : 15;
  }

  function waveClock(h) {
    var n = Number(h) || 0;
    var next = n >= 24;
    var hr = ((n % 24) + 24) % 24;
    var label = (((hr + 11) % 12) + 1) + ":00";
    if (next) label += "a";
    return label;
  }

  function canPlaceRelief(wave) {
    return hospitalHour() >= (wave || 15) - 1;
  }

  function planKey(cat, room) { return cat + "|" + room; }

  function plannedInn(cat, room) {
    var p = (g.reliefPlan || {})[planKey(cat, room)];
    return (p && p.name) ? p : null;
  }

  function collectReliefJobs() {
    return collectLeavingRooms().rooms.map(function (r) {
      return {
        type: r.inn ? "plan" : "empty",
        cat: r.cat, room: r.room, out: r.out, inn: r.inn,
        canPlace: !!(r.inn && canPlaceRelief(r.wave))
      };
    });
  }

  var CAT_ORDER = { nt: 0, st: 1, s100: 2, endo: 3, nora: 4, ep: 5, ccs: 6, fbc: 7 };

  function catOrder(id) {
    return CAT_ORDER[id] != null ? CAT_ORDER[id] : 99;
  }

  function sortReliefRooms(rooms) {
    rooms.sort(function (a, b) {
      if (!!a.latestay !== !!b.latestay) return a.latestay ? 1 : -1;
      if (a.latestay && b.latestay) return (a.lsN || 0) - (b.lsN || 0);
      var da = catOrder(a.cat), db = catOrder(b.cat);
      if (da !== db) return da - db;
      return String(a.room).localeCompare(String(b.room), undefined, { numeric: true });
    });
    return rooms;
  }

  function lateStayRows(wave) {
    return (g.lateStays || (g.assignmentMeta && g.assignmentMeta.lateStays) || []).filter(function (p) {
      return p && p.wave === wave && p.name;
    }).sort(function (a, b) { return (a.n || 0) - (b.n || 0); }).map(function (p) {
      var room = "LS #" + p.n;
      return {
        cat: "ls", catName: "Late stay", room: room,
        out: { name: p.name, shift: p.shift || "", kind: "none" },
        inn: plannedInn("ls", room),
        wave: wave, latestay: true, lsN: p.n
      };
    });
  }

  function collectLeavingAt(wave, used) {
    used = used || {};
    var rooms = [];
    var cats = (typeof CATEGORIES !== "undefined") ? CATEGORIES : [];
    cats.forEach(function (c) {
      var staff = (g.roomStaff || {})[c.id] || {};
      var list = (c.rooms || []).slice();
      Object.keys(staff).forEach(function (r) { if (list.indexOf(r) < 0) list.push(r); });
      list.forEach(function (room) {
        if (!roomIsActive(c.id, room)) return;
        var rec = staff[room];
        if (!rec || rec.closed || !rec.name) return;
        if (shiftEndHour(rec.shift) !== wave) return;
        if (coreShift(rec.shift) === "Dr") return;
        var inn = plannedInn(c.id, room);
        if (!inn) {
          var found = findIncomingFor(room, rec, used);
          if (found && found.p) {
            inn = found.p;
            used[nameKey(found.p.name)] = 1;
          }
        } else {
          used[nameKey(inn.name)] = 1;
        }
        rooms.push({ cat: c.id, catName: c.name, room: room, out: rec, inn: inn, wave: wave });
      });
    });
    return sortReliefRooms(rooms.concat(lateStayRows(wave)));
  }

  function collectDoctorRooms() {
    var rooms = [];
    var cats = (typeof CATEGORIES !== "undefined") ? CATEGORIES : [];
    cats.forEach(function (c) {
      var staff = (g.roomStaff || {})[c.id] || {};
      var list = (c.rooms || []).slice();
      Object.keys(staff).forEach(function (r) { if (list.indexOf(r) < 0) list.push(r); });
      list.forEach(function (room) {
        if (!roomIsActive(c.id, room)) return;
        var rec = staff[room];
        if (!rec || rec.closed || !rec.name) return;
        if (coreShift(rec.shift) !== "Dr") return;
        rooms.push({
          cat: c.id, catName: c.name, room: room, out: rec,
          inn: plannedInn(c.id, room), wave: "md"
        });
      });
    });
    rooms.sort(function (a, b) {
      var da = catOrder(a.cat), db = catOrder(b.cat);
      if (da !== db) return da - db;
      return String(a.room).localeCompare(String(b.room), undefined, { numeric: true });
    });
    return rooms;
  }

  function collectAllLeavingWaves() {
    var hour = hospitalHour();
    var hours = occupantEndHours().filter(function (h) { return h >= hour - 1 && h < 30; });
    (g.lateStays || (g.assignmentMeta && g.assignmentMeta.lateStays) || []).forEach(function (p) {
      var w = p && p.wave;
      if (w >= hour - 1 && w < 30 && hours.indexOf(w) < 0) hours.push(w);
    });
    hours.sort(function (a, b) { return a - b; });
    if (!hours.length) {
      var nxt = nextReliefWave();
      if (nxt < 30) hours = [nxt];
    }
    var used = {};
    return hours.map(function (w) {
      return { wave: w, rooms: collectLeavingAt(w, used) };
    }).filter(function (block) { return block.rooms.length; });
  }

  function collectLeavingRooms() {
    var all = collectAllLeavingWaves();
    if (!all.length) return { wave: nextReliefWave(), rooms: [] };
    return { wave: all[0].wave, rooms: all[0].rooms };
  }

  function collectArrivals() {
    return (g.onDeck || []).filter(function (p) {
      if (!p || !p.name || isJunkStaff(p.name, p.shift)) return false;
      if (p.role === "breaker" || p.role === "call" || p.role === "freed") return false;
      if (isTimeLabel(lastName(p.name))) return false;
      if (p.role === "midnight") return true;
      return shiftStartHour(p.shift) >= 15;
    });
  }

  function countLeavingAt(endH) {
    var n = 0;
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        if (!roomIsActive(cat, room)) return;
        var rec = g.roomStaff[cat][room];
        if (!rec || !rec.name || rec.closed) return;
        if (shiftEndHour(rec.shift) === endH) n++;
      });
    });
    return n;
  }

  function reliefWaveLabel(jobs) {
    var leave = collectLeavingRooms();
    if (leave && leave.rooms && leave.rooms.length) return "out at " + waveClock(leave.wave);
    return "today";
  }

  function closeReliefRoster() {
    var ov = document.getElementById("relief-roster-overlay");
    if (ov) ov.remove();
  }

  function moveStaffToRoom(from, toCat, toRoom, closeFrom) {
    var person = null;
    if (from.kind === "deck") {
      var list = g.onDeck || [];
      var idx = -1;
      for (var i = 0; i < list.length; i++) {
        if (nameKey(list[i].name) === from.key) { idx = i; break; }
      }
      if (idx < 0) return false;
      person = list.splice(idx, 1)[0];
    } else {
      var rec = g.roomStaff[from.cat] && g.roomStaff[from.cat][from.room];
      if (!rec || !rec.name) return false;
      person = { name: rec.name, shift: rec.shift, kind: rec.kind, student: rec.student };
      g.roomStaff[from.cat][from.room] = { name: "", shift: "", kind: "none", closed: !!closeFrom };
      if (closeFrom && typeof catEditState !== "undefined" && catEditState[from.cat]) {
        catEditState[from.cat].deletedRooms.add(from.room);
      }
    }
    g.roomStaff = g.roomStaff || {};
    g.roomStaff[toCat] = g.roomStaff[toCat] || {};
    var occ = g.roomStaff[toCat][toRoom];
    if (occ && occ.name) {
      g.onDeck = g.onDeck || [];
      g.onDeck.unshift({
        name: occ.name, shift: occ.shift || "", kind: occ.kind || "none",
        lastRoom: toRoom, role: "freed", student: !!occ.student
      });
    }
    g.roomStaff[toCat][toRoom] = {
      name: person.name, shift: person.shift || "", kind: person.kind || "none",
      closed: false, student: !!person.student
    };
    if (typeof catEditState !== "undefined" && catEditState[toCat]) {
      catEditState[toCat].deletedRooms.delete(toRoom);
    }
    return true;
  }

  function pickRelief(from, fromCat, fromRoom, fromKey, toCat, toRoom) {
    var person = null;
    if (from === "deck") {
      (g.onDeck || []).forEach(function (p) {
        if (nameKey(p.name) === fromKey) person = p;
      });
    } else {
      var rec = g.roomStaff[fromCat] && g.roomStaff[fromCat][fromRoom];
      if (rec && rec.name) person = rec;
    }
    if (!person || !person.name) return;
    g.reliefPlan = g.reliefPlan || {};
    var k = planKey(toCat, toRoom);
    Object.keys(g.reliefPlan).forEach(function (pk) {
      if (pk !== k && nameKey((g.reliefPlan[pk] || {}).name) === nameKey(person.name)) delete g.reliefPlan[pk];
    });
    g.reliefPlan[k] = {
      name: person.name, shift: person.shift || "", kind: person.kind || "none",
      student: !!person.student, from: from, fromCat: fromCat || "", fromRoom: fromRoom || "", fromKey: fromKey || nameKey(person.name)
    };
    var placed = false;
    if (hospitalHour() >= 14) {
      placed = moveStaffToRoom(
        from === "deck" ? { kind: "deck", key: nameKey(person.name) } : { kind: "room", cat: fromCat, room: fromRoom },
        toCat, toRoom, true
      );
      if (placed) delete g.reliefPlan[k];
    }
    closeReliefRoster();
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    renderOnDeck();
    renderShiftChange();
    try {
      if (typeof showToast === "function") {
        showToast(chipName(person.name) + (placed ? " → " : " planned for ") + toRoom);
      }
    } catch (e) {}
  }

  function rosterPill(p, loc, from, fromCat, fromRoom) {
    var locHtml = loc ? '<span class="roster-loc">' + loc + "</span>" : "";
    return '<button type="button" class="roster-pill" data-rfrom="' + from +
      '" data-rcat="' + (fromCat || "") + '" data-rroom="' + (fromRoom || "") +
      '" data-rkey="' + nameKey(p.name) + '">' + locHtml +
      shiftPillHtml(p.shift) + '<span class="staff-name">' + chipName(p.name) + "</span>" +
      studentMark(hasStudent(p)) + "</button>";
  }

  function openReliefRoster(cat, room) {
    openPersonRoster({
      purpose: "relief",
      cat: cat,
      room: room,
      search: false
    });
  }

  function openStaffRoster(cat, room) {
    openPersonRoster({
      purpose: "staff",
      cat: cat,
      room: room,
      search: true
    });
  }

  function rosterPersonList(opts) {
    var cat = opts.cat, room = opts.room, purpose = opts.purpose;
    var rec = g.roomStaff && g.roomStaff[cat] && g.roomStaff[cat][room];
    var exclude = rec && rec.name ? nameKey(rec.name) : "";
    var wave = nextReliefWave();
    var deck = [];
    (g.onDeck || []).forEach(function (p) {
      if (!p || !p.name || isJunkStaff(p.name, p.shift)) return;
      if (exclude && nameKey(p.name) === exclude) return;
      if (coreShift(p.shift) === "Dr") return;
      if (isTimeLabel(lastName(p.name))) return;
      if (purpose === "relief") {
        var end = shiftEndHour(p.shift);
        if (end <= wave && !isArrivingShift(p.shift) && p.role !== "midnight") return;
      }
      deck.push(p);
    });
    var byCat = [];
    var cats = (typeof CATEGORIES !== "undefined") ? CATEGORIES : [];
    cats.forEach(function (c) {
      var staff = (g.roomStaff || {})[c.id] || {};
      var pills = [];
      (c.rooms || []).forEach(function (r) {
        if (c.id === cat && r === room) return;
        if (purpose !== "staff" && !roomIsActive(c.id, r)) return;
        var s = staff[r];
        if (!s || s.closed || !s.name) return;
        if (exclude && nameKey(s.name) === exclude) return;
        if (coreShift(s.shift) === "Dr") return;
        if (purpose === "relief" && shiftEndHour(s.shift) <= wave) return;
        pills.push({ p: s, room: r, cat: c.id });
      });
      if (pills.length) byCat.push({ name: c.name, pills: pills });
    });
    return { rec: rec, exclude: exclude, deck: deck, byCat: byCat, wave: wave };
  }

  function openPersonRoster(opts) {
    closeReliefRoster();
    var cat = opts.cat, room = opts.room, purpose = opts.purpose || "relief";
    var list = rosterPersonList(opts);
    var rec = list.rec;
    if (purpose === "relief" && (!rec || !rec.name)) return;
    var planned = purpose === "relief" ? plannedInn(cat, room) : null;
    var plannedKey = planned ? nameKey(planned.name) : "";

    var deckPills = list.deck.map(function (p) {
      return rosterPill(p, "", "deck", "", "");
    }).join("");
    var catSections = list.byCat.map(function (sec) {
      var pills = sec.pills.map(function (x) {
        return rosterPill(x.p, x.room, "room", x.cat, x.room);
      }).join("");
      return '<div class="assign-cat-label">' + sec.name + "</div>" +
        '<div class="roster-grid">' + pills + "</div>";
    }).join("");

    var title, sub;
    if (purpose === "staff") {
      title = "Staff " + room;
      sub = rec && rec.name
        ? shiftPillHtml(rec.shift) + '<span class="staff-name">' + chipName(rec.name) + "</span> is here"
        : "Empty · tap who goes in";
    } else {
      title = "Relief for " + room;
      sub = shiftPillHtml(rec.shift) + '<span class="staff-name">' + chipName(rec.name) + "</span> out at " +
        waveClock(list.wave) + (planned ? " · now " + chipName(planned.name) : "");
    }

    var ov = document.createElement("div");
    ov.id = "relief-roster-overlay";
    ov.className = "relief-roster-overlay";
    ov.innerHTML =
      '<div class="relief-roster-card">' +
      '<div class="relief-roster-title">' + title + "</div>" +
      '<div class="relief-roster-sub">' + sub + "</div>" +
      (opts.search ? '<input type="search" class="roster-search" placeholder="Search names" autocomplete="off">' : "") +
      '<div class="relief-roster-body">' +
      '<div class="assign-cat-label">On deck</div>' +
      '<div class="roster-grid">' + (deckPills || '<span class="roster-empty">Nobody free</span>') + "</div>" +
      catSections +
      "</div>" +
      '<button type="button" class="relief-roster-cancel">Cancel</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeReliefRoster(); });
    document.body.appendChild(ov);
    addSheetHandle(ov.querySelector(".relief-roster-card"));
    bindPullDismiss(ov, ".relief-roster-card", closeReliefRoster);
    ov.querySelector(".relief-roster-cancel").onclick = closeReliefRoster;
    ov.querySelectorAll(".roster-pill").forEach(function (el) {
      if (plannedKey && el.getAttribute("data-rkey") === plannedKey) el.classList.add("selected");
      el.onclick = function (ev) {
        ev.stopPropagation();
        var from = el.getAttribute("data-rfrom");
        var fromCat = el.getAttribute("data-rcat");
        var fromRoom = el.getAttribute("data-rroom");
        var fromKey = el.getAttribute("data-rkey");
        if (purpose === "staff") {
          pickStaff(from, fromCat, fromRoom, fromKey, cat, room);
        } else {
          pickRelief(from, fromCat, fromRoom, fromKey, cat, room);
        }
      };
    });
    var search = ov.querySelector(".roster-search");
    if (search) {
      search.focus();
      search.addEventListener("input", function () {
        var q = nameKey(search.value);
        ov.querySelectorAll(".roster-pill").forEach(function (el) {
          var hit = !q || (el.getAttribute("data-rkey") || "").indexOf(q) >= 0;
          el.style.display = hit ? "" : "none";
        });
        ov.querySelectorAll(".assign-cat-label").forEach(function (lab) {
          var grid = lab.nextElementSibling;
          if (!grid || !grid.classList.contains("roster-grid")) return;
          var any = false;
          grid.querySelectorAll(".roster-pill").forEach(function (p) {
            if (p.style.display !== "none") any = true;
          });
          lab.style.display = any ? "" : "none";
          grid.style.display = any ? "" : "none";
        });
      });
    }
  }

  function pickStaff(from, fromCat, fromRoom, fromKey, toCat, toRoom) {
    var moved = moveStaffToRoom(
      from === "deck" ? { kind: "deck", key: fromKey } : { kind: "room", cat: fromCat, room: fromRoom },
      toCat, toRoom, false
    );
    closeReliefRoster();
    if (!moved) return;
    if (typeof weekendActive !== "undefined" && weekendActive[toCat]) weekendActive[toCat][toRoom] = true;
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    renderOnDeck();
    renderShiftChange();
    paintEditDayButtons();
    try { if (typeof showToast === "function") showToast("Staffed " + toRoom); } catch (e) {}
  }

  g.editSitesMode = g.editSitesMode || "rooms";
  g.pendingPour = g.pendingPour || null;

  function setEditSitesHint(text) {
    var el = document.getElementById("edit-sites-hint");
    if (!el) return;
    if (text) {
      el.style.display = "block";
      el.textContent = text;
    } else {
      el.style.display = "none";
      el.textContent = "";
    }
  }

  function setEditSitesMode(mode) {
    g.editSitesMode = mode === "staff" ? "staff" : "rooms";
    g.pendingPour = null;
    var rooms = document.getElementById("edit-mode-rooms");
    var staff = document.getElementById("edit-mode-staff");
    if (rooms) rooms.classList.toggle("on", g.editSitesMode === "rooms");
    if (staff) staff.classList.toggle("on", g.editSitesMode === "staff");
    var overlay = document.getElementById("weekend-overlay");
    if (overlay) overlay.classList.toggle("edit-staffing", g.editSitesMode === "staff");
    if (g.editSitesMode === "staff") setEditSitesHint("Tap a room, then pick who is covering it.");
    else setEditSitesHint("");
  }

  function paintEditDayBtn(btn, catId, room) {
    if (!btn) return;
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    var html = '<span class="wknd-room-id">' + room + "</span>";
    if (rec && rec.name && !rec.closed) html += staffChipHtml(catId, room);
    btn.innerHTML = html;
    if (g.pendingPour && g.pendingPour.cat === catId && g.pendingPour.room === room) {
      btn.classList.add("pour-source");
    } else {
      btn.classList.remove("pour-source");
    }
  }

  function paintEditDayButtons() {
    if (typeof CATEGORIES === "undefined") return;
    CATEGORIES.forEach(function (c) {
      (c.rooms || []).forEach(function (room) {
        var btn = document.getElementById("wknd-" + c.id + "-" + room.replace(/\s/g, "_"));
        if (btn) paintEditDayBtn(btn, c.id, room);
      });
    });
  }

  function applyRoomActive(catId, room, wantOn) {
    if (typeof weekendActive !== "undefined") {
      weekendActive[catId] = weekendActive[catId] || {};
      weekendActive[catId][room] = !!wantOn;
    }
    if (typeof catEditState !== "undefined" && catEditState[catId]) {
      var now = Date.now();
      if (wantOn) {
        catEditState[catId].deletedRooms.delete(room);
        catEditState[catId].deletedEvents[room] = { deleted: false, ts: now };
      } else {
        catEditState[catId].deletedRooms.add(room);
        catEditState[catId].deletedEvents[room] = { deleted: true, ts: now };
      }
    }
    var btn = document.getElementById("wknd-" + catId + "-" + room.replace(/\s/g, "_"));
    if (btn) {
      btn.classList.toggle("active", !!wantOn);
      btn.classList.toggle("room-inactive-marker", !wantOn);
    }
  }

  function closeDeactSheet() {
    var ov = document.getElementById("deact-sheet-overlay");
    if (ov) ov.remove();
  }

  function namedIn(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || rec.closed || !rec.name) return null;
    return rec;
  }

  function deactToDeck(catId, room) {
    pushOnDeckFromRoom(catId, room);
    applyRoomActive(catId, room, false);
    closeDeactSheet();
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    renderOnDeck();
    renderShiftChange();
    paintEditDayButtons();
    try { if (typeof showToast === "function") showToast(room + " closed · on deck"); } catch (e) {}
  }

  function startKeepPouring(catId, room) {
    var rec = namedIn(catId, room);
    closeDeactSheet();
    if (!rec) {
      applyRoomActive(catId, room, false);
      return;
    }
    g.pendingPour = { cat: catId, room: room, name: rec.name };
    setEditSitesMode("rooms");
    setEditSitesHint("Tap the room " + chipName(rec.name) + " moves to. Tap " + room + " to cancel.");
    paintEditDayButtons();
  }

  function finishKeepPouring(toCat, toRoom) {
    var src = g.pendingPour;
    if (!src) return;
    if (src.cat === toCat && src.room === toRoom) {
      g.pendingPour = null;
      setEditSitesHint("");
      paintEditDayButtons();
      return;
    }
    var moved = moveStaffToRoom({ kind: "room", cat: src.cat, room: src.room }, toCat, toRoom, true);
    g.pendingPour = null;
    setEditSitesHint("");
    if (moved) {
      applyRoomActive(src.cat, src.room, false);
      applyRoomActive(toCat, toRoom, true);
    }
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    renderOnDeck();
    renderShiftChange();
    paintEditDayButtons();
    try {
      if (typeof showToast === "function") showToast(chipName(src.name) + " · " + src.room + " → " + toRoom);
    } catch (e) {}
  }

  function openDeactSheet(catId, room) {
    closeDeactSheet();
    var rec = namedIn(catId, room);
    if (!rec) {
      applyRoomActive(catId, room, false);
      try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
      try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
      return;
    }
    var ov = document.createElement("div");
    ov.id = "deact-sheet-overlay";
    ov.className = "deact-sheet-overlay";
    ov.innerHTML =
      '<div class="deact-sheet">' +
      '<div class="deact-kicker">' + room + " closing</div>" +
      '<div class="deact-who">' + shiftPillHtml(rec.shift) + '<span class="staff-name">' + chipName(rec.name) + "</span></div>" +
      '<div class="deact-ask">Where do they go?</div>' +
      '<button type="button" class="deact-choice" data-act="deck"><span class="deact-choice-title">On deck</span><span class="deact-choice-sub">Free — ready for breaks or relief</span></button>' +
      '<button type="button" class="deact-choice" data-act="pour"><span class="deact-choice-title">Move to another room</span><span class="deact-choice-sub">They stay assigned — pick the next OR</span></button>' +
      '<button type="button" class="deact-cancel">Cancel</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeDeactSheet(); });
    ov.querySelector(".deact-cancel").onclick = closeDeactSheet;
    ov.querySelector('[data-act="deck"]').onclick = function () { deactToDeck(catId, room); };
    ov.querySelector('[data-act="pour"]').onclick = function () { startKeepPouring(catId, room); };
    document.body.appendChild(ov);
  }

  function editDayClickRoom(catId, room, btn) {
    if (g.editSitesMode === "staff") {
      openStaffRoster(catId, room);
      return;
    }
    if (g.pendingPour) {
      finishKeepPouring(catId, room);
      return;
    }
    var isOn = typeof weekendActive !== "undefined" && weekendActive[catId] && weekendActive[catId][room];
    if (isOn && namedIn(catId, room)) {
      openDeactSheet(catId, room);
      return;
    }
    applyRoomActive(catId, room, !isOn);
    paintEditDayBtn(btn, catId, room);
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    try { if (typeof updateActivePOCCounter === "function") updateActivePOCCounter(); } catch (e) {}
  }

  g.setEditSitesMode = setEditSitesMode;
  g.paintEditDayButtons = paintEditDayButtons;
  g.editDayClickRoom = editDayClickRoom;
  g.openStaffRoster = openStaffRoster;

  function isArrivingShift(shift) {
    var s = coreShift(shift);
    return s === "E" || s === "e" || s === "N" || s === "t";
  }

  function placeRelief(catId, room, deckKey) {
    if (deckKey) {
      pickRelief("deck", "", "", deckKey, catId, room);
      return;
    }
    openReliefRoster(catId, room);
  }

  function renderShiftChange() {
    ensureRunnerDrawer();
    var boardCard = document.getElementById("card-relief");
    if (boardCard) boardCard.remove();
    var bar = document.getElementById("shift-change-bar");
    if (!bar) {
      bar = document.createElement("div");
      bar.id = "shift-change-bar";
      bar.style.display = "none";
    }
    ensureRunnerDrawer();
    var runner = typeof currentRole === "undefined" || currentRole === "runner";
    var hasAssign = !!(g.assignmentMeta || (g.roomStaff && Object.keys(g.roomStaff).length));
    var leave = collectLeavingRooms();
    var arrivals = collectArrivals();
    var show = runner && hasAssign && (leave.rooms.length || arrivals.length);
    bar.style.display = "none";
    if (document.getElementById("desk-relief-pane") && g.sitesFn === "relief") renderDeskRelief();
    if (!show) return;

    var hour = hospitalHour();
    var chips = leave.rooms.map(function (j) {
      var outNm = chipName(j.out.name);
      var planned = !!plannedInn(j.cat, j.room) || !!j.inn;
      var inNm = j.inn ? chipName(j.inn.name) : "pick";
      var place = "";
      if (canPlaceRelief(j.wave) && j.inn) {
        place = '<span class="relief-need-place">Place</span>';
      }
      return '<button type="button" class="relief-need' + (planned ? " planned" : "") +
        '" data-relcat="' + j.cat + '" data-relroom="' + j.room + '"' +
        (j.inn ? ' data-relwho="' + nameKey(j.inn.name) + '"' : "") + ">" +
        '<span class="relief-room">' + j.room + "</span>" +
        shiftPillHtml(j.out.shift) +
        '<span class="staff-name">' + outNm + "</span>" +
        '<span class="relief-arrow">→</span>' +
        (j.inn ? shiftPillHtml(j.inn.shift) : "") +
        '<span class="staff-name' + (j.inn ? "" : " missing") + '">' + inNm + "</span>" +
        place + "</button>";
    }).join("");

    var arrLine = arrivals.length
      ? '<div class="late-board-row">Coming on ' + arrivals.map(function (p) { return chipName(p.name); }).join(", ") + "</div>"
      : "";

    bar.innerHTML =
      '<div class="shift-change-title">Shift change</div>' +
      '<div class="late-board-row"><strong>' + leave.rooms.length + " out at " + waveClock(leave.wave) +
      "</strong> · tap a room to pick relief</div>" +
      arrLine +
      '<div class="relief-need-grid">' + chips + "</div>";

    bar.querySelectorAll(".relief-need").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        var cat = el.getAttribute("data-relcat");
        var room = el.getAttribute("data-relroom");
        var who = el.getAttribute("data-relwho") || "";
        if (canPlaceRelief(leave.wave) && who && ev.target && ev.target.classList && ev.target.classList.contains("relief-need-place")) {
          var plan = plannedInn(cat, room);
          if (plan && plan.from === "room") pickRelief("room", plan.fromCat, plan.fromRoom, who, cat, room);
          else pickRelief("deck", "", "", who, cat, room);
          return;
        }
        openReliefRoster(cat, room);
      };
    });
    pane.querySelectorAll(".breakq-item[data-qcat]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        focusQueueRoom(el.getAttribute("data-qcat"), el.getAttribute("data-qroom"));
      };
    });
  }

  function occupantOf(cat, room) {
    var rec = g.roomStaff && g.roomStaff[cat] && g.roomStaff[cat][room];
    if (rec && !rec.closed && rec.name) return rec;
    if (cat === "ls") {
      var n = parseInt(String(room).replace(/\D/g, ""), 10);
      var list = g.lateStays || (g.assignmentMeta && g.assignmentMeta.lateStays) || [];
      var p = list.filter(function (x) { return x.n === n; })[0];
      if (p) return { name: p.name, shift: p.shift || "", kind: "none", closed: false };
    }
    return null;
  }

  function isEditDayModal() {
    var b = document.getElementById("weekend-day-badge");
    var ov = document.getElementById("weekend-overlay");
    var label = b ? b.textContent : "";
    return !!(ov && ov.classList.contains("show") && (label === "Active Sites" || label === "Runner"));
  }

  function paintWkndStaff() {
    if (typeof CATEGORIES === "undefined") return;
    CATEGORIES.forEach(function (c) {
      (c.rooms || []).forEach(function (room) {
        var btn = document.getElementById("wknd-" + c.id + "-" + room.replace(/\s/g, "_"));
        if (!btn) return;
        var badge = btn.querySelector(".kind-badge");
        var label = btn.querySelector(".wknd-room");
        if (!label) {
          label = document.createElement("span");
          label.className = "wknd-room";
          var raw = "";
          Array.prototype.slice.call(btn.childNodes).forEach(function (n) {
            if (n.nodeType === 3) raw += n.textContent;
          });
          label.textContent = raw.trim() || room;
          Array.prototype.slice.call(btn.childNodes).forEach(function (n) {
            if (n.nodeType === 3) btn.removeChild(n);
          });
          var old = btn.querySelector(".staff-line");
          if (old) old.remove();
          btn.insertBefore(label, btn.firstChild);
        }
        var oldLine = btn.querySelector(".staff-line");
        if (oldLine) oldLine.remove();
        var html = staffChipHtml(c.id, room);
        if (html) {
          var wrap = document.createElement("span");
          wrap.innerHTML = html;
          if (wrap.firstChild) btn.appendChild(wrap.firstChild);
        }
        if (badge && badge.parentNode === btn) btn.appendChild(badge);
        var on = true;
        var es = typeof catEditState !== "undefined" && catEditState[c.id];
        var dr = es && es.deletedRooms;
        if (dr && typeof dr.has === "function") on = !dr.has(room);
        else if (Array.isArray(dr)) on = dr.indexOf(room) < 0;
        if (typeof weekendActive !== "undefined") {
          if (!weekendActive[c.id]) weekendActive[c.id] = {};
          weekendActive[c.id][room] = on;
        }
        btn.classList.toggle("active", on);
      });
    });
  }

  function sitesTabsHtml() {
    var fn = g.sitesFn || "rooms";
    function tab(id, label) {
      return '<button type="button" class="sites-fn-tab' + (fn === id ? " on" : "") + '" data-sitesfn="' + id + '">' + label + "</button>";
    }
    var hint =
      fn === "people" ? "Add someone who came in, tap a person to change shift or take them off, or tap a room to move them."
      : fn === "relief" ? "Who is leaving next, doctors (usually 5:30), and late or dinner breaks."
      : "Tap a room to open or close it. Closing a staffed room asks where they go.";
    return '<div class="sites-fn-tabs">' +
      tab("rooms", "Rooms") +
      tab("people", "People") +
      tab("relief", "Relief") +
      "</div>" +
      '<div class="sites-fn-hint">' + hint + "</div>" +
      '<button type="button" class="desk-upload-btn" id="desk-upload-btn">Upload assignment sheet</button>' +
      '<button type="button" class="desk-clear-btn" id="desk-clear-btn">Clear board</button>';
  }

  function openBoardDesk(tab) {
    if (tab) g.sitesFn = tab;
    g.sitesFn = g.sitesFn || "rooms";
    if (typeof openEditActiveSites === "function") openEditActiveSites();
  }

  function renderDeskRelief() {
    var pane = document.getElementById("desk-relief-pane");
    if (!pane) return;
    var waves = collectAllLeavingWaves();
    var docs = collectDoctorRooms();
    var late = collectBreakQueue("late");
    var dinner = collectBreakQueue("dinner");
    function chipHtml(j) {
      var planned = !!plannedInn(j.cat, j.room) || !!j.inn;
      var inNm = j.inn ? chipName(j.inn.name) : "pick";
      var canPlace = (j.wave === "md" || canPlaceRelief(j.wave)) && j.inn;
      var place = canPlace ? '<span class="relief-need-place">Place</span>' : "";
      return '<button type="button" class="relief-need' + (planned ? " planned" : "") +
        '" data-relcat="' + j.cat + '" data-relroom="' + j.room + '" data-relwave="' + j.wave + '"' +
        (j.inn ? ' data-relwho="' + nameKey(j.inn.name) + '"' : "") + ">" +
        '<span class="relief-room">' + j.room + "</span>" +
        shiftPillHtml(j.out.shift) +
        '<span class="staff-name">' + chipName(j.out.name) + "</span>" +
        '<span class="relief-arrow">→</span>' +
        (j.inn ? shiftPillHtml(j.inn.shift) : "") +
        '<span class="staff-name' + (j.inn ? "" : " missing") + '">' + inNm + "</span>" +
        place + "</button>";
    }
    var blocks = waves.map(function (block) {
      var chips = block.rooms.map(chipHtml).join("");
      return '<div class="shift-change-title">Out at ' + waveClock(block.wave) +
        ' · ' + block.rooms.length + '</div>' +
        (chips ? '<div class="relief-need-grid">' + chips + "</div>" : "");
    }).join("");
    if (docs.length) {
      blocks += '<div class="shift-change-title">Doctors · usually 5:30 · ' + docs.length + '</div>' +
        '<div class="relief-need-grid">' + docs.map(chipHtml).join("") + "</div>";
    }
    pane.innerHTML =
      (blocks || '<div class="roster-empty">Nobody leaving yet</div>') +
      queueBlockHtml("Late (M+S)", late) +
      queueBlockHtml("Dinner (Q+W+E)", dinner);
    pane.querySelectorAll(".relief-need").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        var cat = el.getAttribute("data-relcat");
        var room = el.getAttribute("data-relroom");
        var who = el.getAttribute("data-relwho") || "";
        var waveRaw = el.getAttribute("data-relwave");
        var canPlace = waveRaw === "md" || canPlaceRelief(parseInt(waveRaw, 10));
        if (canPlace && who && ev.target && ev.target.classList && ev.target.classList.contains("relief-need-place")) {
          var plan = plannedInn(cat, room);
          if (plan && plan.from === "room") pickRelief("room", plan.fromCat, plan.fromRoom, who, cat, room);
          else pickRelief("deck", "", "", who, cat, room);
          renderDeskRelief();
          return;
        }
        openReliefRoster(cat, room);
      };
    });
    pane.querySelectorAll(".breakq-item[data-qcat]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        focusQueueRoom(el.getAttribute("data-qcat"), el.getAttribute("data-qroom"));
      };
    });
  }

  function applyDeskPeopleFilter() {
    var q = nameKey(g.deskPeopleQ || "");
    var pane = document.getElementById("desk-people-pane");
    if (pane) {
      var shown = 0;
      pane.querySelectorAll("[data-deckkey]").forEach(function (el) {
        var hit = !q || (el.getAttribute("data-deckkey") || "").indexOf(q) !== -1;
        el.style.display = hit ? "" : "none";
        if (hit) shown++;
      });
    }
    var rooms = document.getElementById("weekend-rooms-container");
    if (rooms && g.sitesFn === "people") {
      rooms.querySelectorAll(".weekend-room-btn").forEach(function (btn) {
        if (!q) { btn.style.display = ""; return; }
        var id = btn.id || "";
        var m = id.match(/^wknd-([^-]+)-(.*)$/);
        if (!m) { btn.style.display = "none"; return; }
        var rec = occupantOf(m[1], m[2].replace(/_/g, " "));
        btn.style.display = (rec && nameKey(rec.name).indexOf(q) !== -1) ? "" : "none";
      });
    }
  }

  function renderDeskPeople() {
    var pane = document.getElementById("desk-people-pane");
    if (!pane) return;
    var list = (window.onDeck || g.onDeck || []).filter(function (p) {
      return p && p.name && !isJunkStaff(p.name, p.shift);
    });
    var here = [], gone = [];
    list.forEach(function (p) { (isDeckOut(p) ? gone : here).push(p); });
    sortByShiftLen(here);
    sortByShiftLen(gone);
    var q = nameKey(g.deskPeopleQ || "");
    if (q && gone.some(function (p) { return nameKey(p.name).indexOf(q) !== -1; })) g.deskOutOpen = true;
    function pill(p, out) {
      var tag = deckTag(p);
      return '<button type="button" class="roster-pill' + (out ? " out" : "") + '" data-deckkey="' + nameKey(p.name) + '">' +
        shiftPillHtml(p.shift) + '<span class="staff-name">' + chipName(p.name) + "</span>" +
        studentMark(hasStudent(p)) + (tag ? '<span class="staff-last">' + tag + "</span>" : "") +
        (out ? '<span class="staff-last">out</span>' : "") + "</button>";
    }
    var hereChips = here.map(function (p) { return pill(p, false); }).join("");
    var goneChips = gone.map(function (p) { return pill(p, true); }).join("");
    var open = !!g.deskOutOpen;
    var outBlock = gone.length
      ? '<button type="button" class="ondeck-out-toggle" id="desk-out-toggle">' +
        (open ? "Hide out · " : "Out · ") + gone.length + (open ? " ▴" : " ▾") + "</button>" +
        (open ? '<div class="roster-grid">' + goneChips + "</div>" : "")
      : "";
    pane.innerHTML =
      '<input id="desk-people-q" class="desk-people-search" type="search" placeholder="Search a name" autocomplete="off" autocorrect="off" spellcheck="false">' +
      '<button type="button" class="desk-add-btn" id="desk-add-staff">Add someone who came in</button>' +
      '<div class="assign-cat-label">On deck — tap to change shift or remove</div>' +
      '<div class="roster-grid">' + (hereChips || '<span class="roster-empty">Nobody still on shift</span>') + "</div>" +
      outBlock +
      '<div class="sites-fn-hint" style="margin-top:10px">Tap a room below to move someone in, take them off, or change the person in that OR.</div>';
    var add = document.getElementById("desk-add-staff");
    if (add) add.onclick = function () { openAddStaff("", ""); };
    var tog = document.getElementById("desk-out-toggle");
    if (tog) tog.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      g.deskOutOpen = !g.deskOutOpen;
      renderDeskPeople();
    };
    pane.querySelectorAll("[data-deckkey]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        openPersonSheet("deck", el.getAttribute("data-deckkey"), "", "");
      };
    });
    var inp = document.getElementById("desk-people-q");
    if (inp) {
      inp.value = g.deskPeopleQ || "";
      inp.oninput = function () {
        g.deskPeopleQ = inp.value;
        var qk = nameKey(inp.value);
        if (qk && !g.deskOutOpen && gone.some(function (p) { return nameKey(p.name).indexOf(qk) !== -1; })) {
          g.deskOutOpen = true;
          renderDeskPeople();
          var n = document.getElementById("desk-people-q");
          if (n) {
            n.focus();
            try { n.setSelectionRange(n.value.length, n.value.length); } catch (e) {}
          }
          return;
        }
        applyDeskPeopleFilter();
      };
    }
    applyDeskPeopleFilter();
  }

  function scrollDeskToTop() {
    var ov = document.getElementById("weekend-overlay");
    if (ov) {
      ov.scrollTop = 0;
      ov.scrollTo && ov.scrollTo(0, 0);
    }
    var modal = ov && ov.querySelector(".weekend-modal");
    if (modal) modal.scrollTop = 0;
  }

  function enhanceEditDayModal(keepScroll) {
    g.sitesFn = g.sitesFn || "rooms";
    if (g.sitesFn === "staffing") g.sitesFn = "people";
    var title = document.querySelector("#weekend-overlay .weekend-modal-title");
    if (title) title.textContent = "Board desk";
    var badge = document.getElementById("weekend-day-badge");
    if (badge) badge.textContent = "Runner";
    var sub = document.querySelector("#weekend-overlay .weekend-modal-sub");
    if (sub) sub.innerHTML = sitesTabsHtml();
    var tabs = document.querySelector(".sites-fn-tabs");
    if (tabs) {
      tabs.querySelectorAll("[data-sitesfn]").forEach(function (el) {
        el.onclick = function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          g.sitesFn = el.getAttribute("data-sitesfn");
          enhanceEditDayModal(true);
        };
      });
    }
    var up = document.getElementById("desk-upload-btn");
    if (up) {
      up.onclick = function () {
        var el = document.getElementById("assign-file-input");
        if (!el) {
          try { ensureUploadUi(); } catch (e) {}
          el = document.getElementById("assign-file-input");
        }
        if (el) el.click();
      };
    }
    var clr = document.getElementById("desk-clear-btn");
    if (clr) clr.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      openClearBoardSheet();
    };
    var rooms = document.getElementById("weekend-rooms-container");
    var pane = document.getElementById("desk-relief-pane");
    if (!pane && rooms && rooms.parentNode) {
      pane = document.createElement("div");
      pane.id = "desk-relief-pane";
      pane.className = "desk-relief-pane";
      rooms.parentNode.insertBefore(pane, rooms.nextSibling);
    }
    var people = document.getElementById("desk-people-pane");
    if (!people && rooms && rooms.parentNode) {
      people = document.createElement("div");
      people.id = "desk-people-pane";
      people.className = "desk-people-pane";
      rooms.parentNode.insertBefore(people, rooms);
    }
    var onRelief = g.sitesFn === "relief";
    var onPeople = g.sitesFn === "people";
    if (rooms) rooms.style.display = onRelief ? "none" : "";
    if (pane) {
      pane.style.display = onRelief ? "block" : "none";
      if (onRelief) renderDeskRelief();
    }
    if (people) {
      people.style.display = onPeople ? "block" : "none";
      if (onPeople) renderDeskPeople();
    }
    var go = document.querySelector("#weekend-overlay .weekend-go-btn");
    if (go) go.textContent = "Done";
    var back = document.getElementById("weekend-back-btn");
    if (back) {
      back.style.display = "block";
      back.textContent = "Close";
      back.setAttribute("onclick", "document.getElementById('weekend-overlay').classList.remove('show');if(typeof _resetModalToDefaults==='function')_resetModalToDefaults();");
    }
    paintWkndStaff();
    try { installSheetDismiss(); } catch (e) {}
    if (!keepScroll) scrollDeskToTop();
    setTimeout(function () {
      try {
        paintWkndStaff();
        if (g.sitesFn === "people") renderDeskPeople();
        if (g.sitesFn === "relief") renderDeskRelief();
        if (!keepScroll) scrollDeskToTop();
      } catch (e) {}
    }, 30);
  }

  function openClearBoardSheet() {
    closeSitesSheet();
    var ov = document.createElement("div");
    ov.id = "sites-sheet-overlay";
    ov.className = "sites-sheet-overlay";
    ov.innerHTML =
      '<div class="sites-sheet-card">' +
      '<div class="sheet-handle"><span></span></div>' +
      '<div class="relief-roster-title">Clear the board?</div>' +
      '<div class="relief-roster-sub">Removes names, on-deck people, and relief picks. Rooms and break marks stay. This is for testing.</div>' +
      '<button type="button" class="sites-choice" id="desk-clear-yes"><strong>Clear assignment</strong><span>Start over — upload the sheet again</span></button>' +
      '<button type="button" class="relief-roster-cancel" id="desk-clear-no">Cancel</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeSitesSheet(); });
    document.body.appendChild(ov);
    bindPullDismiss(ov, ".sites-sheet-card", closeSitesSheet);
    ov.querySelector("#desk-clear-no").onclick = closeSitesSheet;
    ov.querySelector("#desk-clear-yes").onclick = function () {
      clearAssignmentBoard();
      closeSitesSheet();
    };
  }

  function clearAssignmentBoard() {
    g.roomStaff = {};
    g.onDeck = [];
    g.reliefPlan = {};
    g.lateStays = [];
    g.assignmentMeta = null;
    try { window.roomStaff = g.roomStaff; window.onDeck = g.onDeck; window.reliefPlan = g.reliefPlan; window.assignmentMeta = null; } catch (e) {}
    if (typeof CATEGORIES !== "undefined") {
      CATEGORIES.forEach(function (c) { g.roomStaff[c.id] = {}; });
    }
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    try { renderOnDeck(); } catch (e) {}
    try { paintWkndStaff(); } catch (e) {}
    try { renderDeskPeople(); } catch (e) {}
    try { renderDeskRelief(); } catch (e) {}
    try { renderShiftChange(); } catch (e) {}
    try { renderLateBoardBar(); } catch (e) {}
    try { if (typeof showToast === "function") showToast("Board cleared"); } catch (e) {}
  }

  function addSheetHandle(card) {
    if (!card || card.querySelector(".sheet-handle")) return;
    var h = document.createElement("div");
    h.className = "sheet-handle";
    h.innerHTML = "<span></span>";
    card.insertBefore(h, card.firstChild);
  }

  function bindPullDismiss(ov, cardSel, onClose) {
    if (!ov || ov._pullBound) return;
    ov._pullBound = true;
    ov.addEventListener("click", function (e) {
      if (e.target === ov) onClose();
    });
    var startY = 0, lastY = 0, tracking = false;
    function yOf(e) {
      if (e.touches && e.touches[0]) return e.touches[0].clientY;
      if (e.changedTouches && e.changedTouches[0]) return e.changedTouches[0].clientY;
      return e.clientY;
    }
    function card() { return ov.querySelector(cardSel); }
    ov.addEventListener("pointerdown", function (e) {
      addSheetHandle(card());
      var c = card();
      if (!c) return;
      var t = e.target;
      var y = yOf(e);
      var rect = c.getBoundingClientRect();
      var chrome = !!(t.closest && t.closest(".sheet-handle, .weekend-modal-title, .weekend-modal-sub, .weekend-day-badge, .relief-roster-title, .relief-roster-sub"));
      var fromTop = y >= rect.top - 6 && y <= rect.top + 64;
      if (t !== ov && !chrome && !fromTop) return;
      tracking = true;
      startY = lastY = y;
    });
    ov.addEventListener("pointermove", function (e) {
      if (!tracking) return;
      lastY = yOf(e);
      var dy = Math.max(0, lastY - startY);
      var c = card();
      if (c && dy > 4) {
        c.style.transition = "none";
        c.style.transform = "translateY(" + dy + "px)";
      }
    });
    function end(e) {
      if (!tracking) return;
      tracking = false;
      var dy = (e ? yOf(e) : lastY) - startY;
      var c = card();
      if (c) {
        c.style.transition = "transform .22s ease";
        c.style.transform = "";
      }
      if (dy > 80) onClose();
    }
    ov.addEventListener("pointerup", end);
    ov.addEventListener("pointercancel", end);
  }

  function dismissWeekendOverlay() {
    if (document.getElementById("sites-sheet-overlay")) { closeSitesSheet(); return; }
    if (document.getElementById("relief-roster-overlay")) { closeReliefRoster(); return; }
    if (document.getElementById("assign-upload-overlay")) return;
    var go = document.querySelector("#weekend-overlay .weekend-go-btn");
    if (go && /done/i.test(go.textContent || "")) {
      try {
        if (typeof confirmEditActiveSites === "function") confirmEditActiveSites();
        else document.getElementById("weekend-overlay").classList.remove("show");
      } catch (e) {}
      return;
    }
    var ov = document.getElementById("weekend-overlay");
    if (ov) ov.classList.remove("show");
    try { if (typeof _resetModalToDefaults === "function") _resetModalToDefaults(); } catch (e) {}
  }

  function installSheetDismiss() {
    var ov = document.getElementById("weekend-overlay");
    if (!ov) return;
    addSheetHandle(ov.querySelector(".weekend-modal"));
    bindPullDismiss(ov, ".weekend-modal", dismissWeekendOverlay);
  }

  function closeSitesSheet() {
    var ov = document.getElementById("sites-sheet-overlay");
    if (ov) ov.remove();
  }

  function deactivateRoomNow(cat, room) {
    if (typeof weekendActive !== "undefined" && weekendActive[cat]) weekendActive[cat][room] = false;
    if (typeof catEditState !== "undefined" && catEditState[cat]) {
      catEditState[cat].deletedRooms.add(room);
      catEditState[cat].deletedEvents[room] = { deleted: true, ts: Date.now() };
    }
    var btn = document.getElementById("wknd-" + cat + "-" + room.replace(/\s/g, "_"));
    if (btn) btn.classList.remove("active");
  }

  function pourOff(cat, room) {
    pushOnDeckFromRoom(cat, room);
    deactivateRoomNow(cat, room);
    closeSitesSheet();
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    paintWkndStaff();
    renderOnDeck();
    renderShiftChange();
    try { if (typeof updateActivePOCCounter === "function") updateActivePOCCounter(); } catch (e) {}
  }

  function openPourInto(fromCat, fromRoom) {
    var rec = occupantOf(fromCat, fromRoom);
    if (!rec) return;
    var pills = "";
    (typeof CATEGORIES !== "undefined" ? CATEGORIES : []).forEach(function (c) {
      var group = [];
      (c.rooms || []).forEach(function (r) {
        if (c.id === fromCat && r === fromRoom) return;
        var on = typeof weekendActive !== "undefined" && weekendActive[c.id] && weekendActive[c.id][r];
        if (!on) return;
        var who = occupantOf(c.id, r);
        group.push('<button type="button" class="roster-pill" data-icat="' + c.id + '" data-iroom="' + r + '">' +
          '<span class="roster-loc">' + r + "</span>" +
          (who ? shiftPillHtml(who.shift) + '<span class="staff-name">' + chipName(who.name) + "</span>" : '<span class="staff-name missing">open</span>') +
          "</button>");
      });
      if (group.length) {
        pills += '<div class="assign-cat-label">' + c.name + "</div><div class='roster-grid'>" + group.join("") + "</div>";
      }
    });
    var body = document.querySelector("#sites-sheet-overlay .sites-sheet-body");
    if (!body) return;
    body.innerHTML = '<div class="sites-sheet-lead">Tap the room ' + chipName(rec.name) + " moves to</div>" +
      (pills || '<div class="roster-empty">No other open rooms</div>');
    body.querySelectorAll("[data-iroom]").forEach(function (el) {
      el.onclick = function () {
        var toCat = el.getAttribute("data-icat");
        var toRoom = el.getAttribute("data-iroom");
        moveStaffToRoom({ kind: "room", cat: fromCat, room: fromRoom }, toCat, toRoom, true);
        deactivateRoomNow(fromCat, fromRoom);
        closeSitesSheet();
        try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
        try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
        paintWkndStaff();
        renderOnDeck();
        renderShiftChange();
      };
    });
  }

  function openPourSheet(cat, room) {
    var rec = occupantOf(cat, room);
    if (!rec) return false;
    closeSitesSheet();
    var ov = document.createElement("div");
    ov.id = "sites-sheet-overlay";
    ov.className = "sites-sheet-overlay";
    ov.innerHTML =
      '<div class="sites-sheet-card">' +
      '<div class="relief-roster-title">Closing ' + room + "</div>" +
      '<div class="relief-roster-sub">' + shiftPillHtml(rec.shift) +
      '<span class="staff-name">' + chipName(rec.name) + "</span></div>" +
      '<div class="sites-sheet-body">' +
      '<button type="button" class="sites-choice" data-pour="off"><strong>On deck</strong><span>Free — ready for breaks or relief</span></button>' +
      '<button type="button" class="sites-choice" data-pour="into"><strong>Move to another room</strong><span>They stay assigned — pick the next OR</span></button>' +
      "</div>" +
      '<button type="button" class="relief-roster-cancel" data-pour="keep">Keep open</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeSitesSheet(); });
    document.body.appendChild(ov);
    addSheetHandle(ov.querySelector(".sites-sheet-card"));
    bindPullDismiss(ov, ".sites-sheet-card", closeSitesSheet);
    ov.querySelector('[data-pour="off"]').onclick = function () { pourOff(cat, room); };
    ov.querySelector('[data-pour="into"]').onclick = function () { openPourInto(cat, room); };
    ov.querySelector('[data-pour="keep"]').onclick = closeSitesSheet;
    return true;
  }

  function staffMatches(p, q) {
    if (!q) return true;
    var hay = (lastName(p.name) + " " + (p.name || "") + " " + (p.shift || "")).toLowerCase();
    return hay.indexOf(q) >= 0;
  }

  function openStaffingSheet(cat, room) {
    closeSitesSheet();
    var rec = occupantOf(cat, room);
    var ov = document.createElement("div");
    ov.id = "sites-sheet-overlay";
    ov.className = "sites-sheet-overlay";
    ov.innerHTML =
      '<div class="sites-sheet-card">' +
      '<div class="relief-roster-title">' + room + "</div>" +
      '<div class="relief-roster-sub" id="staffing-who">' +
      (rec ? shiftPillHtml(rec.shift) + '<span class="staff-name">' + chipName(rec.name) + "</span>" + studentMark(hasStudent(rec))
        : '<span class="staff-name missing">Open</span>') + "</div>" +
      '<input class="staffing-search" id="staffing-search" type="search" placeholder="Search a name…" autocomplete="off">' +
      '<div class="sites-sheet-body" id="staffing-list"></div>' +
      '<button type="button" class="relief-roster-cancel">Cancel</button></div>';
    document.body.appendChild(ov);
    addSheetHandle(ov.querySelector(".sites-sheet-card"));
    bindPullDismiss(ov, ".sites-sheet-card", closeSitesSheet);
    ov.addEventListener("click", function (e) { if (e.target === ov) closeSitesSheet(); });
    ov.querySelector(".relief-roster-cancel").onclick = closeSitesSheet;
    var input = document.getElementById("staffing-search");
    function draw() {
      var q = (input.value || "").trim().toLowerCase();
      var exclude = rec ? nameKey(rec.name) : "";
      var deckHtml = (g.onDeck || []).filter(function (p) {
        if (!p || !p.name || isJunkStaff(p.name, p.shift)) return false;
        if (coreShift(p.shift) === "Dr") return false;
        if (exclude && nameKey(p.name) === exclude) return false;
        return staffMatches(p, q);
      }).map(function (p) { return rosterPill(p, "", "deck", "", ""); }).join("");
      var catHtml = "";
      (typeof CATEGORIES !== "undefined" ? CATEGORIES : []).forEach(function (c) {
        var pills = [];
        (c.rooms || []).forEach(function (r) {
          if (c.id === cat && r === room) return;
          var s = occupantOf(c.id, r);
          if (!s) return;
          if (coreShift(s.shift) === "Dr") return;
          if (!staffMatches(s, q) && String(r).toLowerCase().indexOf(q) < 0) return;
          pills.push(rosterPill(s, r, "room", c.id, r));
        });
        if (pills.length) {
          catHtml += '<div class="assign-cat-label">' + c.name + "</div><div class='roster-grid'>" + pills.join("") + "</div>";
        }
      });
      var list = document.getElementById("staffing-list");
      list.innerHTML =
        '<div class="assign-cat-label">☕ On deck</div><div class="roster-grid">' +
        (deckHtml || '<span class="roster-empty">Nobody free</span>') + "</div>" + catHtml;
      list.querySelectorAll(".roster-pill").forEach(function (el) {
        el.onclick = function (ev) {
          ev.stopPropagation();
          var from = el.getAttribute("data-rfrom");
          var ok = moveStaffToRoom(
            from === "deck"
              ? { kind: "deck", key: el.getAttribute("data-rkey") }
              : { kind: "room", cat: el.getAttribute("data-rcat"), room: el.getAttribute("data-rroom") },
            cat, room, false
          );
          if (!ok) return;
          closeSitesSheet();
          try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
          try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
          paintWkndStaff();
          renderOnDeck();
          renderShiftChange();
        };
      });
    }
    input.oninput = draw;
    draw();
    var recNow = occupantOf(cat, room);
    var edit = document.createElement("div");
    edit.className = "staff-edit-row";
    var html = "";
    if (recNow) {
      html += '<button type="button" class="staff-edit-btn" data-se="noshow">Didn\'t come in</button>';
      html += '<button type="button" class="staff-edit-btn" data-se="shift">Change shift</button>';
    }
    html += '<button type="button" class="staff-edit-btn" data-se="add">Add someone</button>';
    edit.innerHTML = html;
    var who = document.getElementById("staffing-who");
    if (who && who.parentNode) who.parentNode.insertBefore(edit, who.nextSibling);
    edit.querySelectorAll("[data-se]").forEach(function (b) {
      b.onclick = function (ev) {
        ev.stopPropagation();
        var act = b.getAttribute("data-se");
        if (act === "noshow") removeOccupant(cat, room);
        else if (act === "shift") openShiftPicker(cat, room);
        else if (act === "add") openAddStaff(cat, room);
      };
    });
    setTimeout(function () { try { input.focus(); } catch (e) {} }, 50);
  }

  var SHIFT_LETTERS = ["D", "S", "M", "Q", "W", "E", "N", "t", "d", "e"];

  function persistStaff() {
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    paintWkndStaff();
    renderOnDeck();
    renderShiftChange();
    renderLateBoardBar();
    try { renderDeskPeople(); } catch (e) {}
  }

  function removeOccupant(cat, room) {
    var rec = occupantOf(cat, room);
    if (!rec) return;
    confirmRemoveStaff(rec.name, function () {
      var nm = chipName(rec.name);
      g.roomStaff[cat][room] = { name: "", shift: "", kind: "none", closed: false };
      closeSitesSheet();
      persistStaff();
      try { if (typeof showToast === "function") showToast(nm + " removed · didn't come in"); } catch (e) {}
    });
  }

  function dropFromDeck(key) {
    g.onDeck = (g.onDeck || []).filter(function (p) { return nameKey(p.name) !== key; });
    persistStaff();
  }

  function closeRemoveConfirm() {
    var ov = document.getElementById("remove-staff-overlay");
    if (ov) ov.remove();
  }

  function confirmRemoveStaff(name, onYes) {
    closeRemoveConfirm();
    var nm = chipName(name) || "this person";
    var ov = document.createElement("div");
    ov.id = "remove-staff-overlay";
    ov.className = "deact-sheet-overlay";
    ov.innerHTML =
      '<div class="deact-sheet">' +
      '<div class="deact-kicker">Removing a staff member</div>' +
      '<div class="deact-who"><span class="staff-name">' + nm + "</span></div>" +
      '<div class="deact-ask">Are you sure? They will be off the board today.</div>' +
      '<button type="button" class="deact-choice" data-act="yes"><span class="deact-choice-title">Yes, remove</span><span class="deact-choice-sub">Off the board — add them back in People if this is a mistake</span></button>' +
      '<button type="button" class="deact-cancel" data-act="no">Keep them</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeRemoveConfirm(); });
    ov.querySelector('[data-act="no"]').onclick = closeRemoveConfirm;
    ov.querySelector('[data-act="yes"]').onclick = function (ev) {
      if (typeof isPlantedTap === "function" && !isPlantedTap(ev, true)) return;
      closeRemoveConfirm();
      if (typeof onYes === "function") onYes();
    };
    document.body.appendChild(ov);
  }

  function installAppTapGuard() {
    if (g._appTapGuard) return;
    g._appTapGuard = true;
    document.addEventListener("click", function (e) {
      var t = e.target;
      if (!t || !t.closest) return;
      if (t.closest("input, textarea, select, label")) return;
      if (t.closest("#assign-upload-overlay, .assign-upload-overlay, .assign-apply, .assign-cancel, #remove-staff-overlay, #sites-sheet-overlay, #force-update-overlay, .sheet-handle, #ondeck-out-toggle")) return;
      var ctrl = typeof closestControl === "function" ? closestControl(t) : t.closest("button, .room-btn, .roster-pill, .relief-need, .sites-fn-tab, .weekend-room-btn");
      if (!ctrl) return;
      var strict = !!t.closest(".ondeck-drop, [data-act='yes'], [data-se='noshow']");
      if (typeof isPlantedTap === "function" && !isPlantedTap(e, strict)) {
        e.preventDefault();
        e.stopPropagation();
        if (e.stopImmediatePropagation) e.stopImmediatePropagation();
      }
    }, true);
  }

  function applyShift(cat, room, letter) {
    applyShiftAnywhere("room", "", cat, room, letter);
  }

  function applyShiftAnywhere(kind, key, cat, room, letter) {
    if (kind === "deck") {
      var p = (g.onDeck || []).filter(function (x) { return nameKey(x.name) === key; })[0];
      if (!p) return;
      p.shift = letter;
      p.kind = breakKind(letter);
      closeSitesSheet();
      persistStaff();
      try { if (typeof showToast === "function") showToast(chipName(p.name) + " · " + letter); } catch (e) {}
      renderDeskPeople();
      return;
    }
    g.roomStaff = g.roomStaff || {};
    g.roomStaff[cat] = g.roomStaff[cat] || {};
    var rec = g.roomStaff[cat][room] || { name: "", closed: false };
    rec.shift = letter;
    rec.kind = breakKind(letter);
    g.roomStaff[cat][room] = rec;
    closeSitesSheet();
    persistStaff();
    try { if (typeof showToast === "function") showToast(chipName(rec.name) + " · " + letter); } catch (e) {}
    renderDeskPeople();
  }

  function openPersonSheet(kind, key, cat, room) {
    closeSitesSheet();
    var rec = kind === "deck"
      ? (g.onDeck || []).filter(function (x) { return nameKey(x.name) === key; })[0]
      : occupantOf(cat, room);
    if (!rec) return;
    var ov = document.createElement("div");
    ov.id = "sites-sheet-overlay";
    ov.className = "sites-sheet-overlay";
    ov.innerHTML =
      '<div class="sites-sheet-card">' +
      '<div class="relief-roster-title">' + chipName(rec.name) + "</div>" +
      '<div class="relief-roster-sub">' + shiftPillHtml(rec.shift) +
      (kind === "deck" ? '<span class="staff-last">on deck</span>' : '<span class="roster-loc">' + room + "</span>") +
      "</div>" +
      '<div class="sites-sheet-lead">Change shift — stay later or go home earlier</div>' +
      '<div class="shift-pick">' + SHIFT_LETTERS.map(function (l) {
        var on = coreShift(rec.shift) === l ? " on" : "";
        return '<button type="button" class="shift-pick-btn' + on + '" data-sh="' + l + '">' + l + "</button>";
      }).join("") + "</div>" +
      '<button type="button" class="sites-choice" data-se="remove"><strong>Remove from the board</strong><span>They didn\'t come in, or they\'re done for the day</span></button>' +
      '<button type="button" class="relief-roster-cancel">Cancel</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeSitesSheet(); });
    document.body.appendChild(ov);
    ov.querySelector(".relief-roster-cancel").onclick = closeSitesSheet;
    ov.querySelectorAll("[data-sh]").forEach(function (b) {
      b.onclick = function () { applyShiftAnywhere(kind, key, cat, room, b.getAttribute("data-sh")); };
    });
    ov.querySelector('[data-se="remove"]').onclick = function () {
      if (kind === "deck") confirmRemoveStaff(rec.name, function () { dropFromDeck(key); closeSitesSheet(); renderDeskPeople(); });
      else removeOccupant(cat, room);
    };
  }

  function openShiftPicker(cat, room) {
    openPersonSheet("room", "", cat, room);
  }

  function openAddStaff(cat, room) {
    closeSitesSheet();
    var toRoom = !!(cat && room);
    var ov = document.createElement("div");
    ov.id = "sites-sheet-overlay";
    ov.className = "sites-sheet-overlay";
    var picked = { shift: "S" };
    ov.innerHTML =
      '<div class="sites-sheet-card">' +
      '<div class="relief-roster-title">Add someone</div>' +
      '<div class="relief-roster-sub">' + (toRoom ? "Into " + room : "They'll sit On deck until you place them") + "</div>" +
      '<input class="add-staff-name" id="add-staff-name" type="text" placeholder="Last name" autocomplete="off" autocapitalize="words">' +
      '<div class="sites-sheet-lead">Shift</div>' +
      '<div class="shift-pick" id="add-staff-shifts">' + SHIFT_LETTERS.map(function (l) {
        var on = l === "S" ? " on" : "";
        return '<button type="button" class="shift-pick-btn' + on + '" data-sh="' + l + '">' + l + "</button>";
      }).join("") + "</div>" +
      '<button type="button" class="sites-choice" id="add-staff-go"><strong>' +
      (toRoom ? "Put them in " + room : "Add to On deck") +
      "</strong><span>" + (toRoom ? "If someone is already here they go On deck" : "Place them into a room from People when you're ready") +
      "</span></button>" +
      '<button type="button" class="relief-roster-cancel">Cancel</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closeSitesSheet(); });
    document.body.appendChild(ov);
    ov.querySelector(".relief-roster-cancel").onclick = closeSitesSheet;
    ov.querySelectorAll("[data-sh]").forEach(function (b) {
      b.onclick = function () {
        picked.shift = b.getAttribute("data-sh");
        ov.querySelectorAll("[data-sh]").forEach(function (x) { x.classList.toggle("on", x === b); });
      };
    });
    document.getElementById("add-staff-go").onclick = function () {
      var name = ((document.getElementById("add-staff-name") || {}).value || "").trim();
      if (!name) {
        try { if (typeof showToast === "function") showToast("Type a name"); } catch (e) {}
        return;
      }
      if (toRoom) {
        g.roomStaff = g.roomStaff || {};
        g.roomStaff[cat] = g.roomStaff[cat] || {};
        var occ = occupantOf(cat, room);
        if (occ) {
          g.onDeck = g.onDeck || [];
          g.onDeck.unshift({ name: occ.name, shift: occ.shift || "", kind: occ.kind || "none", lastRoom: room, role: "freed", student: !!occ.student });
        }
        g.roomStaff[cat][room] = { name: name, shift: picked.shift, kind: breakKind(picked.shift), closed: false };
        if (typeof catEditState !== "undefined" && catEditState[cat]) {
          catEditState[cat].deletedRooms.delete(room);
          catEditState[cat].deletedEvents[room] = { deleted: false, ts: Date.now() };
        }
        if (typeof weekendActive !== "undefined") {
          if (!weekendActive[cat]) weekendActive[cat] = {};
          weekendActive[cat][room] = true;
        }
      } else {
        g.onDeck = g.onDeck || [];
        g.onDeck.unshift({ name: name, shift: picked.shift, kind: breakKind(picked.shift), lastRoom: "", role: "extra" });
      }
      closeSitesSheet();
      persistStaff();
      renderDeskPeople();
      try {
        if (typeof showToast === "function") {
          showToast(chipName(name) + (toRoom ? " → " + room : " · on deck"));
        }
      } catch (e) {}
    };
    setTimeout(function () {
      var el = document.getElementById("add-staff-name");
      if (el) el.focus();
    }, 40);
  }

  function wrapSitesModal() {
    if (g._sitesWrapped) return;
    g._sitesWrapped = true;
    if (typeof buildEveningModal === "function") {
      var origBuild = buildEveningModal;
      window.buildEveningModal = function (mode) {
        if (mode !== "edit-day" && mode !== "poc") {
          try { if (typeof openEditActiveSites === "function") openEditActiveSites(); } catch (e) {}
          return;
        }
        if (mode === "edit-day") g.sitesFn = g.sitesFn || "rooms";
        origBuild.apply(this, arguments);
        if (mode === "edit-day") enhanceEditDayModal();
      };
    }
    if (typeof reopenEveningModal === "function") {
      window.reopenEveningModal = function () {
        if (typeof openEditActiveSites === "function") openEditActiveSites();
      };
    }
    if (typeof buildWeekendModal === "function") {
      window.buildWeekendModal = function () {
        return;
      };
    }
    if (typeof reopenWeekendModal === "function") {
      window.reopenWeekendModal = function () {
        if (typeof openEditActiveSites === "function") openEditActiveSites();
      };
    }
    if (typeof weekendToggleRoom === "function") {
      var origToggle = weekendToggleRoom;
      window.weekendToggleRoom = function (catId, room, btn) {
        if (isEditDayModal() || (document.getElementById("weekend-day-badge") || {}).textContent === "Active Sites") {
          if (g.sitesFn === "staffing" || g.sitesFn === "people") {
            openStaffingSheet(catId, room);
            return;
          }
          if (g.sitesFn === "relief") return;
          var on = typeof weekendActive !== "undefined" && weekendActive[catId] && weekendActive[catId][room];
          if (on && occupantOf(catId, room)) {
            openPourSheet(catId, room);
            return;
          }
        }
        origToggle.apply(this, arguments);
        paintWkndStaff();
      };
    }
    if (typeof weekendToggleCat === "function") {
      var origCat = weekendToggleCat;
      window.weekendToggleCat = function () {
        origCat.apply(this, arguments);
        paintWkndStaff();
      };
    }
    if (typeof confirmEditActiveSites === "function") {
      var origSave = confirmEditActiveSites;
      window.confirmEditActiveSites = function () {
        if (typeof CATEGORIES !== "undefined" && typeof weekendActive !== "undefined") {
          CATEGORIES.forEach(function (c) {
            (c.rooms || []).forEach(function (room) {
              var want = weekendActive[c.id] && weekendActive[c.id][room];
              if (want) return;
              if (occupantOf(c.id, room)) pushOnDeckFromRoom(c.id, room);
            });
          });
        }
        origSave.apply(this, arguments);
        renderOnDeck();
        renderShiftChange();
        renderLateBoardBar();
      };
    }
    if (typeof applyPayload === "function" && !g._payloadWrapped) {
      g._payloadWrapped = true;
      var origPay = applyPayload;
      window.applyPayload = function () {
        var out = origPay.apply(this, arguments);
        try {
          if (window.roomStaff) g.roomStaff = window.roomStaff;
          if (Array.isArray(window.onDeck)) g.onDeck = window.onDeck;
          if (g.assignmentMeta && g.assignmentMeta.lateStays) g.lateStays = g.assignmentMeta.lateStays;
          renderOnDeck();
          if (document.getElementById("weekend-overlay") && document.getElementById("weekend-overlay").classList.contains("show")) {
            paintWkndStaff();
            if (g.sitesFn === "people") renderDeskPeople();
            if (g.sitesFn === "relief") renderDeskRelief();
          }
        } catch (e) {}
        return out;
      };
    }
    if (typeof openEditActiveSites === "function" && !g._openDeskWrapped) {
      g._openDeskWrapped = true;
      var origOpen = openEditActiveSites;
      window.openEditActiveSites = function () {
        origOpen.apply(this, arguments);
        scrollDeskToTop();
        setTimeout(function () { try { enhanceEditDayModal(); scrollDeskToTop(); } catch (e) {} }, 0);
      };
    }
  }

  function wrapDeactivate() {
    if (typeof toggleDeleteRoom === "function" && !g._deckWrapped) {
      g._deckWrapped = true;
      var orig = toggleDeleteRoom;
      window.toggleDeleteRoom = function (catId, room) {
        if (g.selectedDeck && assignSelectedTo(catId, room)) return;
        var was = typeof catEditState !== "undefined" && catEditState[catId] && catEditState[catId].deletedRooms.has(room);
        orig.apply(this, arguments);
        var now = typeof catEditState !== "undefined" && catEditState[catId] && catEditState[catId].deletedRooms.has(room);
        if (!was && now) pushOnDeckFromRoom(catId, room);
        else renderOnDeck();
      };
    }
    if (typeof toggleRoom === "function" && !g._toggleRoomWrapped) {
      g._toggleRoomWrapped = true;
      var origToggle = toggleRoom;
      window.toggleRoom = function (catId, room, ev) {
        if (g.selectedDeck && assignSelectedTo(catId, room)) return;
        return origToggle.apply(this, arguments);
      };
    }
    if (typeof setWindow === "function" && !g._setWindowWrapped) {
      g._setWindowWrapped = true;
      var origSet = setWindow;
      window.setWindow = function () {
        origSet.apply(this, arguments);
        renderOnDeck();
        renderLateBoardBar();
        renderShiftChange();
      };
    }
    if (typeof refreshBoard === "function" && !g._refreshWrapped) {
      g._refreshWrapped = true;
      var origRefresh = refreshBoard;
      window.refreshBoard = function () {
        origRefresh.apply(this, arguments);
        renderOnDeck();
        renderLateBoardBar();
        renderShiftChange();
        hideTowerEditChrome();
      };
    }
    wrapSitesModal();
    wrapTowerEdit();
  }

  function hideTowerEditChrome() {
    if (typeof CATEGORIES === "undefined") return;
    CATEGORIES.forEach(function (c) {
      var card = document.getElementById("card-" + c.id);
      if (!card) return;
      var tabs = card.querySelector(".cat-mode-tabs");
      if (tabs) tabs.style.display = "none";
    });
  }

  function wrapTowerEdit() {
    if (g._towerEditWrapped) return;
    g._towerEditWrapped = true;
    if (typeof renderCatCard === "function") {
      var orig = renderCatCard;
      window.renderCatCard = function () {
        orig.apply(this, arguments);
        hideTowerEditChrome();
      };
    }
    if (typeof updateEditActiveBar === "function") {
      var origBar = updateEditActiveBar;
      window.updateEditActiveBar = function () {
        origBar.apply(this, arguments);
        var bar = document.getElementById("edit-active-bar");
        if (bar) bar.style.display = "none";
        ensureRunnerDrawer();
      };
    }
  }

  function showAssignSummary(res, file) {
    var existing = document.getElementById("assign-upload-overlay");
    if (existing) existing.remove();
    var ov = document.createElement("div");
    ov.id = "assign-upload-overlay";
    ov.className = "assign-upload-overlay";
    var today = typeof todayStr === "function" ? todayStr() : "";
    var dateWarn = res.date && today && res.date !== today
      ? '<div class="assign-warn">Sheet date is ' + res.date + " (today is " + today + "). Apply anyway?</div>"
      : "";
    var um = (res.unmatched || []).slice(0, 6).map(function (u) {
      return "<li>" + (u.room || "") + " — " + (u.staff || "") + "</li>";
    }).join("");
    var hideN = 0;
    if (typeof CATEGORIES !== "undefined") {
      var named = {};
      (res.rooms || []).forEach(function (r) {
        if (r && r.cat && r.room && r.name && !r.closed) named[r.cat + "|" + r.room] = 1;
      });
      CATEGORIES.forEach(function (c) {
        (c.rooms || []).forEach(function (room) {
          if (!named[c.id + "|" + room]) hideN++;
        });
      });
    }
    ov.innerHTML =
      '<div class="assign-upload-card">' +
      "<h3>Sheet applied</h3>" +
      "<p class='assign-meta'>" + (res.kind || "") + (res.date ? " · " + res.date : "") + (file ? " · " + file.name : "") + "</p>" +
      dateWarn +
      "<ul class='assign-stats'>" +
      "<li>" + (res.rooms || []).length + " rooms with a name</li>" +
      "<li>" + (res.closed || []).length + " marked CLOSED</li>" +
      (hideN ? "<li>" + hideN + " not on sheet (hidden)</li>" : "") +
      "<li>" + (res.late || []).length + " late (M+S)</li>" +
      "<li>" + (res.dinner || []).length + " dinner (Q+W+E)</li>" +
      "<li>" + ((res.onDeck || []).length) + " on deck (not in a room)</li>" +
      (res.unmatched && res.unmatched.length ? "<li>" + res.unmatched.length + " unmatched labels</li>" : "") +
      "</ul>" +
      (um ? "<ul class='assign-unmatched'>" + um + "</ul>" : "") +
      '<div class="assign-actions">' +
      '<button type="button" class="assign-apply">Done</button>' +
      "</div></div>";
    document.body.appendChild(ov);
    function closeOv(ev) {
      if (ev) { try { ev.preventDefault(); ev.stopPropagation(); } catch (e) {} }
      ov.remove();
    }
    ov.addEventListener("click", function (e) {
      if (e.target === ov || (e.target.closest && e.target.closest(".assign-apply"))) closeOv(e);
    }, true);
    var done = ov.querySelector(".assign-apply");
    done.onclick = closeOv;
    done.addEventListener("pointerup", closeOv);
    done.addEventListener("touchend", closeOv, { passive: false });
  }

  async function handleAssignmentFile(file) {
    if (!file) return;
    try {
      var buf = await file.arrayBuffer();
      var res = await parseAssignmentWorkbook(buf, file.name);
      applyAssignmentResult(res);
      showAssignSummary(res, file);
    } catch (e) {
      try { if (typeof showToast === "function") showToast("Could not read that sheet"); } catch (x) {}
      console.warn(e);
    }
  }

  function ensureUploadUi() {
    if (document.getElementById("assign-file-input")) return;
    var input = document.createElement("input");
    input.type = "file";
    input.id = "assign-file-input";
    input.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    input.style.display = "none";
    input.addEventListener("change", function () {
      var f = input.files && input.files[0];
      input.value = "";
      handleAssignmentFile(f);
    });
    document.body.appendChild(input);
    if (!document.getElementById("assign-upload-css")) {
      var st = document.createElement("style");
      st.id = "assign-upload-css";
      st.textContent =
        ".staff-chip{display:none;}" +
        ".staff-line{display:flex;align-items:center;justify-content:center;gap:2px;margin-top:2px;white-space:nowrap;overflow:hidden;max-width:100%;}" +
        ".shift-pill{flex:0 0 auto;font-size:8px;font-weight:800;letter-spacing:.02em;padding:1px 4px;border-radius:5px;background:rgba(122,78,45,.14);color:#7A4E2D;line-height:1.2;}" +
        ".shift-pill.dr{background:transparent;border:1px solid rgba(122,78,45,.3);font-weight:700;font-size:7px;padding:1px 3px;}" +
        ".shift-pill.wide{font-size:7px;padding:1px 3px;letter-spacing:0;}" +
        ".shift-pill.time{font-size:6.5px;padding:1px 3px;letter-spacing:0;white-space:nowrap;}" +
        ".ondeck-chip .shift-pill{position:static;}" +
        ".staff-name{font-size:10px;font-weight:600;color:#1E0E04;letter-spacing:-0.03em;}" +
        ".staff-name.long{font-size:8px;}" +
        ".staff-name.tiny{font-size:7px;letter-spacing:-0.05em;}" +
        ".staff-stu{font-size:9px;font-weight:800;color:#C8781A;margin-left:1px;line-height:1;flex-shrink:0;}" +
        ".staff-last{font-size:10px;font-weight:500;color:#9A6A38;}" +
        ".staff-last::before{content:'·';margin:0 3px;color:rgba(122,78,45,.45);}" +
        ".ondeck-hint{padding:0 12px 6px;font-size:11px;color:#9A6A38;}" +
        "body.assigning .room-btn{box-shadow:inset 0 0 0 1.5px rgba(122,78,45,.35);}" +
        "#late-board-bar{display:none;padding:8px 14px 10px;background:transparent;border-bottom:none;font-size:12px;color:#1E0E04;}" +
        "#late-board-bar .late-board-title{font-weight:800;font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:#7A4E2D;margin-bottom:4px;}" +
        "#late-board-bar .late-board-row{margin:2px 0;line-height:1.35;}" +
        "#edit-active-bar,#evening-runner-bar{background:transparent!important;border-bottom:none!important;}" +
        ".runner-drawer{flex-shrink:0;background:#F7F0E6;border-bottom:1px solid rgba(160,98,42,.15);}" +
        ".runner-drawer-handle{padding:8px 14px 8px;cursor:pointer;touch-action:none;-webkit-user-select:none;user-select:none;}" +
        ".runner-drawer-knob{width:36px;height:4px;border-radius:2px;background:rgba(122,78,45,.28);margin:0 auto 6px;}" +
        ".runner-drawer-peek{font-size:11px;font-weight:700;color:#7A4E2D;text-align:center;letter-spacing:.01em;}" +
        ".runner-drawer.open .runner-drawer-peek{opacity:.55;font-weight:600;}" +
        ".runner-drawer-body{display:none;overflow-y:auto;-webkit-overflow-scrolling:touch;max-height:58vh;padding-bottom:8px;}" +
        ".runner-drawer.open .runner-drawer-body{display:block;}" +
        ".breakq-list{display:flex;flex-wrap:wrap;gap:5px;margin:4px 0 8px;}" +
        ".breakq-item{display:inline-flex;align-items:center;gap:3px;background:#FDF6EC;border:1px solid rgba(160,98,42,.2);border-radius:8px;padding:4px 7px;font:inherit;color:#1E0E04;cursor:pointer;-webkit-appearance:none;appearance:none;}" +
        ".breakq-item.given{opacity:.38;}" +
        ".breakq-n{font-size:9px;font-weight:800;color:#9A6A38;min-width:10px;}" +
        ".breakq-given-n{font-size:11px;color:#9A6A38;margin-left:6px;font-weight:500;}" +
        ".room-btn.queue-flash{box-shadow:inset 0 0 0 1.5px #7A4E2D;background:#F5E6D0;}" +
        ".assign-upload-overlay{position:fixed;inset:0;background:rgba(30,14,4,.55);z-index:9000;display:flex;align-items:center;justify-content:center;padding:18px;}" +
        ".assign-upload-card{background:#fff;border-radius:16px;padding:18px 16px;max-width:360px;width:100%;color:#1E0E04;}" +
        ".assign-upload-card h3{margin:0 0 6px;font-size:18px;}" +
        ".assign-meta{font-size:12px;color:#7A4E2D;margin:0 0 8px;}" +
        ".assign-warn{background:#FDEBD0;padding:8px;border-radius:8px;font-size:12px;margin-bottom:8px;}" +
        ".assign-stats{margin:0 0 8px;padding-left:18px;font-size:13px;}" +
        ".assign-unmatched{font-size:11px;color:#6b5344;max-height:90px;overflow:auto;}" +
        ".assign-actions{display:flex;gap:8px;margin-top:12px;}" +
        ".assign-actions button{flex:1;padding:12px;border:none;border-radius:10px;font-weight:700;}" +
        ".assign-cancel{background:#eee;color:#1E0E04;-webkit-appearance:none;appearance:none;}" +
        ".assign-apply{background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;}" +
        "#edit-active-bar .assign-upload-btn{margin-top:6px;width:100%;font-size:11px;font-weight:700;padding:7px;border-radius:7px;border:1px solid rgba(122,78,45,.35);background:#fff;color:#7A4E2D;cursor:pointer;}" +
        ".ondeck-card{border:1px dashed rgba(122,78,45,.28);}" +
        ".ondeck-grid{display:flex;flex-wrap:wrap;gap:6px;padding:4px 12px 12px;}" +
        ".ondeck-chip{display:inline-flex;align-items:center;gap:3px;white-space:nowrap;max-width:100%;background:#FDF6EC;border:1px solid rgba(160,98,42,.22);border-radius:8px;padding:5px 8px;cursor:pointer;-webkit-appearance:none;appearance:none;font:inherit;color:inherit;}" +
        ".ondeck-chip.selected{border-color:#7A4E2D;background:#F5E6D0;box-shadow:inset 0 0 0 1px #7A4E2D;}" +
        ".ondeck-out-toggle{display:flex;align-items:center;justify-content:center;gap:6px;margin:2px 12px 10px;padding:8px 10px;width:calc(100% - 24px);box-sizing:border-box;border-radius:8px;border:1px dashed rgba(122,78,45,.35);background:#FBF6F0;color:#7A4E2D;font:inherit;font-size:12px;font-weight:700;cursor:pointer;-webkit-appearance:none;appearance:none;}" +
        ".ondeck-out-grid{opacity:.95;padding-top:0;}" +
        ".staff-name.gone{opacity:.42;}" +
        ".relief-card{border:1px dashed rgba(122,78,45,.28);}" +
        ".relief-list{padding:0 12px 12px;display:flex;flex-direction:column;gap:6px;}" +
        ".relief-row{display:flex;align-items:center;gap:8px;background:#FDF6EC;border:1px solid rgba(160,98,42,.16);border-radius:10px;padding:7px 8px;}" +
        ".relief-room{flex:0 0 auto;font-size:11px;font-weight:800;color:#7A4E2D;min-width:52px;}" +
        ".relief-who{flex:1;min-width:0;display:flex;align-items:center;gap:5px;flex-wrap:wrap;font-size:11px;color:#1E0E04;}" +
        ".relief-out,.relief-in{display:inline-flex;align-items:center;gap:3px;}" +
        ".relief-in.missing .staff-name{color:#9A6A38;font-weight:500;}" +
        ".relief-arrow{color:#C4A882;font-weight:700;}" +
        ".relief-place{flex-shrink:0;font-size:11px;font-weight:800;padding:6px 10px;border-radius:8px;border:none;background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;cursor:pointer;-webkit-appearance:none;appearance:none;}" +
        ".relief-place.fill{background:#fff;color:#7A4E2D;border:1px solid rgba(122,78,45,.35);}" +
        ".relief-wait{flex-shrink:0;font-size:10px;font-weight:600;color:#9A6A38;}" +
        ".relief-arrivals{display:flex;flex-wrap:wrap;gap:6px;padding:0 12px 12px;}" +
        "#shift-change-bar{display:none;padding:4px 14px 10px;font-size:12px;color:#1E0E04;}" +
        ".shift-change-title{font-weight:800;font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:#7A4E2D;margin:4px 0;}" +
        ".relief-need-grid{display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-top:6px;}" +
        ".relief-need{display:flex;align-items:center;gap:3px;min-width:0;width:100%;overflow:hidden;background:#FDF6EC;border:1px solid rgba(160,98,42,.2);border-radius:8px;padding:5px 6px;font:inherit;color:#1E0E04;cursor:pointer;-webkit-appearance:none;appearance:none;}" +
        ".relief-need.planned{border-color:#7A4E2D;background:#F5E6D0;}" +
        ".relief-need .relief-room{flex:0 0 auto;font-size:10px;font-weight:800;white-space:nowrap;}" +
        ".relief-need .staff-name{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:10px;font-weight:700;}" +
        ".relief-need .shift-pill{flex-shrink:0;}" +
        ".relief-need .relief-arrow{flex-shrink:0;font-size:10px;}" +
        ".relief-need .staff-name.missing{color:#9A6A38;font-weight:500;}" +
        ".relief-need-place{flex-shrink:0;font-size:9px;font-weight:800;padding:2px 5px;border-radius:5px;background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;margin-left:1px;}" +
        ".relief-roster-overlay{position:fixed;inset:0;background:rgba(30,14,4,.55);z-index:820;display:flex;align-items:center;justify-content:center;padding:16px;}" +
        ".relief-roster-card{background:#fff;border-radius:16px;padding:16px 14px 14px;width:100%;max-width:380px;max-height:86vh;overflow-y:auto;-webkit-overflow-scrolling:touch;color:#1E0E04;box-shadow:0 16px 48px rgba(30,14,4,.22);}" +
        ".relief-roster-title{font-family:Georgia,'Times New Roman',serif;font-size:18px;font-weight:800;margin:0 0 4px;}" +
        ".relief-roster-sub{font-size:12px;color:#7A4E2D;margin:0 0 12px;display:flex;align-items:center;gap:4px;flex-wrap:wrap;}" +
        ".relief-roster-body .assign-cat-label{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#9A6A38;margin:12px 0 6px;}" +
        ".roster-grid{display:flex;flex-wrap:wrap;gap:6px;}" +
        ".roster-pill{display:inline-flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;min-width:72px;padding:7px 6px 6px;border-radius:8px;border:1px solid rgba(160,98,42,.22);background:#FDF6EC;color:#1E0E04;cursor:pointer;font:inherit;-webkit-appearance:none;appearance:none;}" +
        ".roster-pill.selected{border-color:#7A4E2D;background:#F5E6D0;box-shadow:inset 0 0 0 1px #7A4E2D;}" +
        ".roster-loc{font-size:10px;font-weight:800;color:#7A4E2D;}" +
        ".roster-empty{font-size:11px;color:#9A6A38;padding:4px 2px;}" +
        ".relief-roster-cancel{width:100%;margin-top:14px;padding:11px;border:none;border-radius:10px;font-size:13px;font-weight:700;color:#7A4E2D;background:#F3EDE6;cursor:pointer;}" +
        ".weekend-room-btn{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;min-height:46px;padding:7px 4px 6px;text-decoration:none;}" +
        ".weekend-room-btn .wknd-room{font-size:11px;font-weight:700;line-height:1.1;}" +
        ".weekend-room-btn:not(.active) .wknd-room{text-decoration:line-through;}" +
        ".weekend-room-btn .staff-line{margin-top:0;display:flex;align-items:center;justify-content:center;gap:2px;}" +
        ".weekend-room-btn .staff-name{font-size:10px;font-weight:700;color:#4A2E14;}" +
        ".weekend-room-btn.active .staff-name{color:#5A2E0A;}" +
        ".weekend-room-btn:not(.active) .staff-line{opacity:.55;}" +
        ".sites-fn-tabs{display:flex;gap:6px;justify-content:center;margin:10px 0 8px;}" +
        ".sites-fn-tab{flex:1;max-width:140px;padding:8px 10px;border-radius:20px;border:1.5px solid rgba(160,98,42,.25);background:#FEF6EC;color:#7A4E2D;font-size:12px;font-weight:700;cursor:pointer;}" +
        ".sites-fn-tab.on{background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;border-color:#7A4E2D;}" +
        ".sites-fn-hint{font-size:12px;color:#7A4E2D;line-height:1.45;margin:0 0 8px;}" +
        ".sites-sheet-overlay{position:fixed;inset:0;background:rgba(30,14,4,.55);z-index:860;display:flex;align-items:center;justify-content:center;padding:16px;}" +
        ".sites-sheet-card{background:#fff;border-radius:16px;padding:16px 14px 14px;width:100%;max-width:380px;max-height:86vh;overflow-y:auto;-webkit-overflow-scrolling:touch;color:#1E0E04;box-shadow:0 16px 48px rgba(30,14,4,.22);}" +
        ".sites-choice{width:100%;display:flex;flex-direction:column;align-items:flex-start;gap:2px;text-align:left;padding:12px 14px;margin:0 0 8px;border-radius:12px;border:1.5px solid rgba(160,98,42,.22);background:#FDF6EC;color:#1E0E04;cursor:pointer;font:inherit;}" +
        ".sites-choice strong{font-size:15px;color:#1E0E04;}" +
        ".sites-choice span{font-size:12px;color:#7A4E2D;}" +
        ".sites-sheet-lead{font-size:13px;font-weight:700;color:#7A4E2D;margin:0 0 8px;}" +
        ".staffing-search{width:100%;box-sizing:border-box;margin:4px 0 10px;padding:11px 12px;border-radius:10px;border:1.5px solid rgba(160,98,42,.3);font-size:15px;color:#1E0E04;outline:none;background:#fff;}" +
        ".staff-edit-row{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 10px;}" +
        ".staff-edit-btn{padding:8px 10px;border-radius:8px;border:1.5px solid rgba(160,98,42,.22);background:#FDF6EC;color:#1E0E04;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit;}" +
        ".shift-pick{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 12px;}" +
        ".shift-pick-btn{min-width:42px;padding:10px 8px;border-radius:8px;border:1.5px solid rgba(160,98,42,.22);background:#FDF6EC;color:#1E0E04;font-size:14px;font-weight:800;cursor:pointer;font-family:inherit;}" +
        ".shift-pick-btn.on{background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;border-color:#7A4E2D;}" +
        ".add-staff-name{width:100%;box-sizing:border-box;padding:11px 12px;border-radius:10px;border:1.5px solid rgba(160,98,42,.3);font-size:15px;color:#1E0E04;outline:none;background:#fff;margin:0 0 8px;}" +
        ".ondeck-chip{position:relative;}" +
        ".ondeck-drop{position:absolute;top:-6px;right:-6px;width:22px;height:22px;border-radius:50%;background:#7A4E2D;color:#fff;font-size:13px;line-height:22px;text-align:center;font-weight:800;border:none;padding:0;cursor:pointer;font-family:inherit;}" +
        ".room-btn.kind-late{background:#FFF0D8;border-color:#C8781A;}" +
        ".room-btn.kind-dinner{background:#7A4E2D;border-color:#5A3418;color:#FDF6EC;}" +
        ".room-btn.kind-dinner .staff-name{color:#FDF6EC;}" +
        ".room-btn.kind-dinner .shift-pill{background:#FDF6EC;color:#7A4E2D;}" +
        ".room-btn.kind-late.done{box-shadow:inset 0 0 0 2px rgba(122,78,45,.35);}" +
        ".room-btn.kind-dinner.done{opacity:.88;box-shadow:inset 0 0 0 2px rgba(253,246,236,.45);}" +
        ".runner-drawer.desk-only{padding:8px 14px 10px;background:#fff;border-bottom:1px solid rgba(160,98,42,.12);}" +
        ".board-desk-btn{width:100%;display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 14px;border:none;border-radius:12px;background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;cursor:pointer;font:inherit;text-align:left;box-shadow:0 1px 3px rgba(122,78,45,.2);}" +
        ".board-desk-btn .desk-title{font-size:15px;font-weight:800;}" +
        ".board-desk-btn .desk-peek{font-size:11px;font-weight:600;opacity:.9;}" +
        ".desk-upload-btn{width:100%;margin:0 0 8px;padding:8px;border-radius:8px;border:1.5px solid rgba(160,98,42,.25);background:#fff;color:#7A4E2D;font-size:12px;font-weight:700;cursor:pointer;}" +
        ".desk-clear-btn{width:100%;margin:0 0 10px;padding:8px;border-radius:8px;border:1.5px dashed rgba(160,98,42,.35);background:transparent;color:#9A6A38;font-size:12px;font-weight:700;cursor:pointer;}" +
        ".desk-people-pane{padding:4px 0 8px;}" +
        ".desk-people-search{width:100%;box-sizing:border-box;margin:0 0 10px;padding:11px 12px;border-radius:10px;border:1px solid rgba(160,98,42,.28);background:#fff;color:#1E0E04;font:inherit;font-size:15px;}" +
        ".desk-people-search:focus{outline:none;border-color:#7A4E2D;box-shadow:0 0 0 3px rgba(122,78,45,.12);}" +
        ".desk-add-btn{width:100%;margin:0 0 10px;padding:11px 12px;border-radius:10px;border:1.5px dashed rgba(160,98,42,.4);background:#FDF6EC;color:#7A4E2D;font-size:13px;font-weight:800;cursor:pointer;font-family:inherit;}" +
        ".sheet-handle{display:flex;justify-content:center;padding:4px 0 10px;cursor:grab;touch-action:none;}" +
        ".sheet-handle span{width:40px;height:4px;border-radius:2px;background:rgba(122,78,45,.3);}" +
        ".sites-fn-tab{max-width:none;}";
      document.head.appendChild(st);
    }
  }

  function mountUploadButton() {
    ensureUploadUi();
    var bar = document.getElementById("edit-active-bar");
    if (!bar || bar.querySelector(".assign-upload-btn")) return;
    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "assign-upload-btn";
    btn.textContent = "Upload CRNA assignment sheet";
    btn.onclick = function () {
      var el = document.getElementById("assign-file-input");
      if (el) el.click();
    };
    bar.appendChild(btn);
  }

  g.roomStaff = g.roomStaff || {};
  g.onDeck = g.onDeck || [];
  g.reliefPlan = g.reliefPlan || {};
  g.lateStays = g.lateStays || [];
  g.assignmentMeta = g.assignmentMeta || null;
  g.parseAssignmentWorkbook = parseAssignmentWorkbook;
  g.applyAssignmentResult = applyAssignmentResult;
  g.handleAssignmentFile = handleAssignmentFile;
  g.staffChipHtml = staffChipHtml;
  g.renderLateBoardBar = renderLateBoardBar;
  g.renderOnDeck = renderOnDeck;
  g.renderShiftChange = renderShiftChange;
  g.ensureRunnerDrawer = ensureRunnerDrawer;
  g.mountUploadButton = mountUploadButton;
  g.ensureUploadUi = ensureUploadUi;
  g.collectBreakQueue = collectBreakQueue;
  g.paintWkndStaff = paintWkndStaff;
  g.enhanceEditDayModal = enhanceEditDayModal;
  g.openBoardDesk = openBoardDesk;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { ensureUploadUi(); wrapDeactivate(); wrapSitesModal(); installAppTapGuard(); installSheetDismiss(); });
  } else { ensureUploadUi(); wrapDeactivate(); wrapSitesModal(); installAppTapGuard(); installSheetDismiss(); }
  if (!g._reliefTimer) {
    g._reliefTimer = setInterval(function () {
      try { renderShiftChange(); renderOnDeck(); renderLateBoardBar(); } catch (e) {}
    }, 60000);
  }
  setTimeout(wrapDeactivate, 500);
})(window);
