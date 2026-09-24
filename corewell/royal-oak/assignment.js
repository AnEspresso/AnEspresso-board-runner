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
    var iso = val.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) return parseInt(iso[1], 10) + "-" + parseInt(iso[2], 10) + "-" + parseInt(iso[3], 10);
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

  function extractRoomTime(s) {
    var fromAt = extractFirstCase(s);
    if (fromAt) return fromAt;
    var m = String(s || "").trim().match(/\s+(?:@\s*)?(?:0?(\d{1,2})[:.](\d{2})|0?([6-9]\d{2}|1[0-2]\d{2}))(?:\s*[ap]m?)?\s*$/i);
    if (!m) return "";
    if (m[1]) return parseInt(m[1], 10) + ":" + m[2];
    var d = String(m[3] || "");
    if (d.length === 3) return d.charAt(0) + ":" + d.slice(1);
    if (d.length === 4) return parseInt(d.slice(0, 2), 10) + ":" + d.slice(2);
    return "";
  }

  function stripRoomTime(s) {
    var t = String(s || "").replace(/\s+/g, " ").trim();
    if (!extractRoomTime(t)) return t;
    return t.replace(/\s+(?:@\s*)?(?:0?\d{1,2}[:.]\d{2}|0?[6-9]\d{2}|1[0-2]\d{2})(?:\s*[ap]m?)?\s*$/i, "").trim();
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

  function escHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "\x26amp;")
      .replace(/</g, "\x26lt;")
      .replace(/>/g, "\x26gt;")
      .replace(/"/g, "\x26quot;")
      .replace(/'/g, "\x26#39;");
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
    if (/^(rotate|available|closed)$/i.test(n)) return true;
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
    var endH = end.h;
    if (endH <= start.h) endH += 24;
    return { startH: start.h, endH: endH, label: compactRangeLabel(m[1], m[2]) };
  }

  function findTimeRange(s) {
    var m = String(s || "").match(/(\d{1,2}(?::\d{2})?\s*[ap]m?)\s*[-–to]+\s*(\d{1,2}(?::\d{2})?\s*[ap]?m?)/i);
    if (!m) return null;
    return parseTimeRange(m[1] + "-" + m[2]);
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
    m = rest.match(/^(Dr\.)\s*(.+)$/i);
    if (m && /[A-Za-z]/.test(m[2])) return rec("Dr", m[2], "none");
    m = rest.match(/^(Dr)\s+(.+)$/i);
    if (m) return rec("Dr", m[2], "none");
    m = rest.match(/^([DdMmSsQqWwEeNnTt](?:\/[DdMmSsQqWwEeNnTt])?)\*?\s+(.+)$/);
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

  function deckRoleForUnknown(roomRaw) {
    var t = String(roomRaw || "").replace(/\s+/g, " ").trim();
    if (!t) return null;
    if (/SRNA/i.test(t) || (/^OB\b/i.test(t) && findTimeRange(t))) {
      var tr = findTimeRange(t);
      var ob = /OB/i.test(t);
      return { role: ob ? "resident" : "srna", last: ob ? "OB" : "SRNA", shift: tr ? tr.label : "", student: !ob };
    }
    if (/^OB\s*(RESIDENT|RES)\b/i.test(t)) {
      return { role: "resident", last: "OB" };
    }
    if (/^EVES?$/i.test(t)) return { role: "evening", last: "Eves" };
    if (isTimeLabel(t)) return { role: "night", last: t };
    return null;
  }

  function parseShiftLabel(s) {
    var t = String(s || "").replace(/\s+/g, "").trim();
    var m = t.match(/^(o\/)?(\*?)([DdMmSsQqWwEeNnTt](?:\/[DdMmSsQqWwEeNnTt])?)(\*?)$/);
    if (!m) return null;
    var orient = !!m[1];
    var early = !!(m[2] || m[4]);
    var sh = normalizeShift(m[3]);
    return { shift: fmtShift(sh, early, orient), kind: breakKind(sh), early: early, orient: orient, letter: sh };
  }

  function normalizeRoom(raw) {
    if (!ROOM_INDEX) buildRoomIndex();
    var s = stripRoomTime(raw);
    if (!s) return null;
    var ste = s.match(/^STE\.?\s*(10[1-9])$/i);
    if (ste) {
      var steRoom = "OR " + ste[1];
      if (ROOM_INDEX[steRoom]) return ROOM_INDEX[steRoom];
      return { cat: "s100", room: steRoom };
    }
    if ((/^OB(@|\s|$)/i.test(s) || /^ENDO$/i.test(s) || /^MRI$/i.test(s)) && !/POC/i.test(s)) {
      /* OB, Endo, and MRI are real weekend assignments */
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
    if (su === "IR") su = "IR";
    su = su.replace(/^CCL\b.*/, "CATH LAB");
    su = su.replace(/^BMBX?\d*.*/, "BMB");
    su = su.replace(/^CATH(?:ETER)?\s*LAB(?:ORATORY)?(?:\s*\d+)?$/, "CATH LAB");
    su = su.replace(/^CT\b.*/, "CT");
    su = su.replace(/^TEE\b.*/, "TEE");
    su = su.replace(/^VCU\b.*/, "VCU");
    su = su.replace(/OR\s*39\s*\/\s*40(?:x\d+)?/i, "OR 38-39");
    su = su.replace("OR39/40", "OR 38-39").replace("OR 39/40", "OR 38-39");
    if (/^MRI IC$/i.test(su)) su = "MRI IC 1";
    else if (/^MRI IC\s*1\b/i.test(su)) su = "MRI IC 1";
    else if (/^MRI IC\s*2\b/i.test(su)) su = "MRI IC 2";
    if (su === "OB") su = "OB 1";
    if (su === "ENDO") su = "ENDO 1";
    if (su === "MRI") su = "MRI IH";
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

  function guessCatForLabel(raw) {
    var loc = normalizeRoom(raw);
    if (loc) return loc;
    var s = stripRoomTime(raw);
    var su = s.toUpperCase();
    var n = parseInt((su.match(/\d+/) || [])[0], 10);
    if (n) {
      if (n >= 1 && n <= 36) return { cat: "nt", room: /^OR\b/i.test(s) ? s : "OR " + n };
      if (n >= 51 && n <= 66) return { cat: "st", room: /^OR\b/i.test(s) ? s : "OR " + n };
      if (n >= 71 && n <= 76) return { cat: "ccs", room: /^OR\b/i.test(s) ? s : "OR " + n };
      if (n >= 100 && n <= 109) return { cat: "s100", room: /^OR\b/i.test(s) ? s : "OR " + n };
    }
    if (/^ENDO/i.test(su)) return { cat: "endo", room: s };
    if (/^(EP|TEE|2231)/i.test(su)) return { cat: "ep", room: s };
    if (/^OB/i.test(su)) return { cat: "fbc", room: s };
    if (/^STE/i.test(su)) return { cat: "s100", room: s };
    return { cat: "nora", room: s || String(raw || "").trim() };
  }

  function expandComboRooms(raw) {
    var s = String(raw || "").trim();
    if (!s || s.indexOf("+") < 0) return [s];
    var parts = s.split(/\s*\+\s*/).map(function (x) { return x.trim(); }).filter(Boolean);
    if (parts.length < 2) return [s];
    var first = parts[0];
    return parts.map(function (p, i) {
      if (i === 0) return p;
      if (/^\d{3,4}$/.test(p)) return p;
      if (/^\d{1,2}$/.test(p)) {
        if (/^EP/i.test(first)) return "EP " + p;
        if (/^ENDO/i.test(first)) return "Endo " + p;
        return "OR " + p;
      }
      return p;
    });
  }

  function suggestFor(roomRaw) {
    var s = String(roomRaw || "").trim();
    var out = [];
    if (typeof CATEGORIES === "undefined") return out;
    function take(pred) {
      CATEGORIES.forEach(function (c) {
        (c.rooms || []).forEach(function (r) {
          if (pred(c.id, r)) out.push({ cat: c.id, room: r });
        });
      });
    }
    if (/^IR\b/i.test(s)) take(function (id, r) { return /^IR\s/i.test(r); });
    else if (/^E$/i.test(s)) take(function (id, r) { return id === "endo"; });
    else if (/SRNA\s*OB|OB\s*RESIDENT/i.test(s)) take(function (id, r) { return id === "fbc"; });
    return out;
  }

  function bareName(n) {
    return String(n || "").replace(/^Dr\.?\s*/i, "").trim();
  }

  function flipCommaName(s) {
    var m = String(s || "").trim().match(/^([^,]+),\s*(.+)$/);
    if (!m) return cleanName(s);
    return cleanName(m[2] + " " + m[1]);
  }

  function editDistance(a, b) {
    a = String(a || "");
    b = String(b || "");
    if (a === b) return 0;
    if (!a || !b) return 99;
    if (Math.abs(a.length - b.length) > 2) return 99;
    var prev = [];
    var i, j;
    for (j = 0; j <= b.length; j++) prev[j] = j;
    for (i = 1; i <= a.length; i++) {
      var cur = [i];
      var rowMin = i;
      for (j = 1; j <= b.length; j++) {
        var cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
        if (cur[j] < rowMin) rowMin = cur[j];
      }
      if (rowMin > 2) return 99;
      prev = cur;
    }
    return prev[b.length];
  }

  function namesMatch(a, b) {
    if (!a || !b) return false;
    if (nameKey(bareName(a)) === nameKey(bareName(b))) return true;
    var la = lastName(a).toLowerCase(), lb = lastName(b).toLowerCase();
    var fa = firstName(a).toLowerCase(), fb = firstName(b).toLowerCase();
    if (!la || !lb) return false;
    var lastOk = la === lb || (la.length >= 6 && lb.length >= 6 && editDistance(la, lb) <= 2);
    if (!lastOk) return false;
    if (!fa || !fb) return true;
    if (fa === fb) return true;
    if (fa.charAt(0) === fb.charAt(0) && (fa.length === 1 || fb.length === 1)) return true;
    if (la === lb && fa.charAt(0) === fb.charAt(0)) return true;
    return false;
  }

  function nameKey(n) {
    return String(n || "").toLowerCase().replace(/[^a-z]/g, "");
  }

  function deckRoleRank(role) {
    return ({
      breaker: 50, wbf: 40, call: 30, midnight: 25,
      evening: 22, resident: 22, night: 20, latestay: 15,
      extra: 10, offsite: 10, unplaced: 1, shift: 1
    }[role] || 0);
  }

  function mergeDeckPeople(list) {
    var by = {};
    var order = [];
    (list || []).forEach(function (p) {
      if (!p || !p.name) return;
      var k = nameKey(p.name);
      if (!by[k]) {
        by[k] = {
          name: p.name, shift: p.shift || "", kind: p.kind || "none",
          lastRoom: p.lastRoom || "", role: p.role || "",
          intended: p.intended || "", student: !!p.student, firstCase: p.firstCase || ""
        };
        order.push(k);
        return;
      }
      var cur = by[k];
      if (deckRoleRank(p.role) > deckRoleRank(cur.role)) {
        if (!p.shift && cur.shift) p = Object.assign({}, p, { shift: cur.shift, kind: p.kind || cur.kind });
        if (!p.intended && cur.intended) p = Object.assign({}, p, { intended: cur.intended });
        cur.role = p.role || cur.role;
        cur.lastRoom = p.lastRoom || cur.lastRoom;
        cur.shift = p.shift || cur.shift;
        cur.kind = p.kind || cur.kind || "none";
        cur.intended = p.intended || cur.intended;
        cur.student = cur.student || !!p.student;
        cur.firstCase = cur.firstCase || p.firstCase || "";
      } else {
        if (!cur.shift && p.shift) { cur.shift = p.shift; cur.kind = p.kind || cur.kind; }
        if (!cur.intended && p.intended) cur.intended = p.intended;
        cur.student = cur.student || !!p.student;
      }
    });
    return order.map(function (k) { return by[k]; });
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

  function isLeavingSoon(shift, p) {
    var t = personHours(p || { shift: shift });
    if (!t || t.unknown || t.end == null || t.end === 99) return false;
    var nowM = hospitalMins();
    var endM;
    if (t.end > 24) {
      if (nowM >= (t.start * 60)) endM = t.end * 60;
      else endM = (t.end - 24) * 60;
    } else endM = t.end * 60;
    var until = endM - nowM;
    return until >= 0 && until <= 70;
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

  function hospitalMins() {
    if (typeof g._minsOverride === "number") return g._minsOverride;
    try {
      var d = typeof hospitalNow === "function" ? hospitalNow() : new Date(new Date().toLocaleString("en-US", { timeZone: "America/Detroit" }));
      return d.getHours() * 60 + d.getMinutes();
    } catch (e) {
      var n = new Date();
      return n.getHours() * 60 + n.getMinutes();
    }
  }

  function personHours(p) {
    var shift = (p && p.shift) || "";
    var tr = parseTimeRange(shift) || findTimeRange(shift) || findTimeRange((p && p.lastRoom) || "");
    if (tr) return { start: tr.startH, end: tr.endH };
    var start = shiftStartHour(shift);
    var end = shiftEndHour(shift);
    if (end === 99) return { start: start, end: 99, unknown: true };
    if (end <= start) end += 24;
    return { start: start, end: end };
  }

  function canMoveStaff() {
    try { return typeof currentRole !== "undefined" && currentRole === "runner"; } catch (e) { return false; }
  }

  function stillInHouse(shift) {
    return hospitalHour() < shiftEndHour(shift);
  }

  function breakWindowFor(kind) {
    return kind === "dinner" ? 3 : 2;
  }

  function breakGiven(kind, catId, room) {
    if (!catId || !room || typeof state === "undefined") return false;
    function marked(w) {
      try {
        return !!(state[w] && state[w][catId] && state[w][catId][room]);
      } catch (e) {
        return false;
      }
    }
    if (marked(breakWindowFor(kind))) return true;
    if (typeof currentWindow === "number" && marked(currentWindow)) return true;
    return false;
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
    var calls = [];
    var filled = {};
    var mriIcN = 0;
    var unsure = [];
    var preview = [];
    var sheetPeople = [];
    function pushPreview(section, room, staff) {
      room = (room || "").trim();
      staff = (staff || "").trim();
      if (!room && !staff) return;
      preview.push({ section: section, room: room, staff: staff });
    }
    function recordPerson(roomRaw, staffRaw, st, locs, role, refs) {
      if (!st || !st.name || st.closed) return;
      if (isJunkStaff(st.name, st.shift)) return;
      sheetPeople.push({
        roomRaw: roomRaw || "",
        staffRaw: staffRaw || "",
        name: st.name,
        shift: st.shift || "",
        student: !!st.student,
        kind: st.kind || "none",
        locs: locs || [],
        role: role || "",
        cellRefs: refs || []
      });
    }
    function addUnsure(roomRaw, staffRaw, st, refs) {
      if (!st || !st.name || st.closed) return;
      unsure.push({
        roomRaw: roomRaw,
        staffRaw: staffRaw,
        name: st.name,
        shift: st.shift || "",
        student: !!st.student,
        kind: st.kind || "none",
        suggestions: suggestFor(roomRaw),
        cellRefs: refs || []
      });
    }
    function addDeck(st, lastRoom, role) {
      if (!st || !st.name || st.closed) return;
      if (isJunkStaff(st.name, st.shift)) return;
      if (/^\d{3,4}\s*-\s*\d/.test(st.name) || /^2200/.test(st.name)) return;
      var intended = "";
      if (role === "extra" || role === "offsite") intended = lastRoom || "";
      if (role === "wbf" && lastRoom && !/^WBF$/i.test(lastRoom)) intended = lastRoom;
      onDeck.push({
        name: st.name, shift: st.shift || "", kind: st.kind || "none",
        lastRoom: lastRoom || "", role: role || "",
        intended: intended,
        student: !!st.student, firstCase: st.firstCase || ""
      });
    }
    function addPair(roomRaw, staffRaw, src, roomRef, staffRef) {
      roomRaw = (roomRaw || "").trim();
      staffRaw = (staffRaw || "").trim();
      if (!roomRaw) return;
      if (skipRoomLabel(roomRaw) && !/^STE\.?\s*10[1-9]$/i.test(roomRaw)) return;
      if (!staffRaw) return;
      var refs = [roomRef, staffRef].filter(Boolean);
      var locRoom = roomRaw;
      if (/^MRI IC\b/i.test(roomRaw) && roomRaw.indexOf("+") < 0) {
        mriIcN += 1;
        locRoom = mriIcN === 1 ? "MRI IC 1" : "MRI IC 2";
      }
      var st = parseStaff(staffRaw);
      var labels = expandComboRooms(locRoom);
      var locs = [];
      labels.forEach(function (lab) {
        var loc = normalizeRoom(lab);
        if (loc) locs.push(loc);
      });
      if (!locs.length) {
        if (st && st.closed) return;
        if (st && st.name && !st.closed && !isJunkStaff(st.name, st.shift)) {
          var shiftLab = parseShiftLabel(roomRaw);
          var deckGuess = deckRoleForUnknown(roomRaw);
          if (shiftLab) {
            if (!st.shift) {
              st.shift = shiftLab.shift;
              st.kind = shiftLab.kind;
              st.early = shiftLab.early;
              st.orient = shiftLab.orient;
            }
            addDeck(st, "", "shift");
            unsure.push({
              reason: "shift",
              roomRaw: roomRaw,
              staffRaw: staffRaw,
              name: st.name,
              shift: st.shift,
              student: !!st.student,
              kind: st.kind || "none",
              suggestions: [],
              guessedShift: shiftLab.shift,
              cellRefs: refs
            });
            recordPerson(roomRaw, staffRaw, st, [], "shift", refs);
          } else if (deckGuess) {
            if (deckGuess.shift && !st.shift) {
              st.shift = deckGuess.shift;
              st.kind = breakKind(st.shift) || st.kind;
            }
            if (deckGuess.student) st.student = true;
            addDeck(st, deckGuess.last, deckGuess.role);
            recordPerson(roomRaw, staffRaw, st, [], deckGuess.role, refs);
          } else {
            addDeck(st, roomRaw, "unplaced");
            addUnsure(roomRaw, staffRaw, st, refs);
            recordPerson(roomRaw, staffRaw, st, [], "unplaced", refs);
            unmatched.push({ room: roomRaw, staff: staffRaw });
          }
        } else if (st && st.closed) {
          return;
        } else if (/rotate/i.test(staffRaw || "")) {
          return;
        } else if (!st) {
          unmatched.push({ room: roomRaw, staff: staffRaw });
        }
        return;
      }
      if (!st) return;
      if (st.closed) {
        locs.forEach(function (loc) {
          var key = loc.cat + "|" + loc.room;
          if (filled[key] && filled[key].closed) return;
          var rec = {
            cat: loc.cat, room: loc.room, shift: "", name: "",
            closed: true, kind: "none", early: false, orient: false,
            student: false, firstCase: ""
          };
          if (filled[key]) {
            rooms = rooms.filter(function (x) { return !(x.cat === loc.cat && x.room === loc.room); });
          }
          filled[key] = rec;
          rooms.push(rec);
          closed.push({ cat: loc.cat, room: loc.room });
        });
        return;
      }
      if (isJunkStaff(st.name, st.shift)) return;
      var isCombo = labels.length > 1;
      var placedLocs = [];
      var bumped = false;
      locs.forEach(function (loc) {
        var key = loc.cat + "|" + loc.room;
        var isSte = src === "ste" || /^STE\.?\s*10[1-9]$/i.test(roomRaw);
        if (filled[key] && !st.closed && !isSte && !isCombo) {
          addDeck(st, loc.room, src === "cd" ? "offsite" : "extra");
          bumped = true;
          return;
        }
        if (filled[key] && isCombo) return;
        var rec = {
          cat: loc.cat, room: loc.room, shift: st.shift || "", name: st.name || "",
          closed: !!st.closed, kind: st.kind || "none", early: !!st.early, orient: !!st.orient,
          student: !!st.student, firstCase: st.firstCase || extractRoomTime(roomRaw) || ""
        };
        if (filled[key]) {
          rooms = rooms.filter(function (x) { return !(x.cat === loc.cat && x.room === loc.room); });
        }
        filled[key] = rec;
        rooms.push(rec);
        placedLocs.push(loc);
        if (rec.closed) closed.push({ cat: loc.cat, room: loc.room });
      });
      recordPerson(roomRaw, staffRaw, st, isCombo ? locs : placedLocs, bumped && !placedLocs.length ? "extra" : (isCombo ? "combo" : ""), refs);
    }
    var ntImplied = [];
    try {
      var ntCat = (typeof CATEGORIES !== "undefined") && CATEGORIES.filter(function (c) { return c.id === "nt"; })[0];
      (ntCat && ntCat.rooms || []).forEach(function (room) {
        var n = parseInt(String(room).replace(/\D/g, ""), 10);
        if (n && n <= 21) ntImplied.push(room);
      });
    } catch (e) {
      ntImplied = ["OR 1","OR 2","OR 3","OR 4","OR 7","OR 8","OR 9","OR 10","OR 11","OR 12","OR 14","OR 15","OR 16","OR 17","OR 20","OR 21"];
    }
    var ntSeq = -1;
    for (var r = 4; r <= 44; r++) {
      var aRaw = cell(cells, "A", r);
      var bRaw = cell(cells, "B", r);
      var aUse = aRaw;
      if (/NORTH TOWER/i.test(aRaw)) {
        ntSeq = 0;
      } else if (/^OR\s*22\b/i.test(aRaw) || /^CCS$/i.test(aRaw)) {
        ntSeq = 99;
      } else if (!aRaw && bRaw && ntSeq >= 0 && ntSeq < ntImplied.length) {
        aUse = ntImplied[ntSeq];
        ntSeq += 1;
      }
      pushPreview("NT", aUse, bRaw);
      addPair(aUse, bRaw, "ab", "A" + r, "B" + r);
      var eRaw = cell(cells, "E", r);
      var fRaw = cell(cells, "F", r);
      var dRawEf = cell(cells, "D", r);
      pushPreview("ST", eRaw || dRawEf, fRaw);
      if (!eRaw && /^OR\s*\d/i.test(dRawEf) && fRaw) addPair(dRawEf, fRaw, "df", "D" + r, "F" + r);
      else addPair(eRaw, fRaw, "ef", "E" + r, "F" + r);
    }
    var inWbf = false;
    for (r = 4; r <= 44; r++) {
      var cRaw = cell(cells, "C", r);
      var dRaw = cell(cells, "D", r);
      if (/^WBF$/i.test(cRaw)) { inWbf = true; continue; }
      pushPreview(inWbf ? "WBF" : "Mid", cRaw, dRaw);
      if (inWbf) {
        var wst = parseStaff(dRaw);
        if (wst && wst.name && !wst.closed && !isJunkStaff(wst.name, wst.shift)) {
          addDeck(wst, cRaw, "wbf");
          recordPerson(cRaw || "WBF", dRaw, wst, [], "wbf", ["C" + r, "D" + r]);
        }
        continue;
      }
      addPair(cRaw, dRaw, "cd", "C" + r, "D" + r);
    }
    for (r = 4; r <= 44; r++) {
      var steLab = cell(cells, "E", r);
      if (/^STE\.?\s*10[1-9]$/i.test(steLab)) addPair(steLab, cell(cells, "F", r), "ste", "E" + r, "F" + r);
    }
    function addLabelled(raw, lastRoom, role, refs) {
      var st = parseStaff(raw);
      if (st && st.name) {
        if (role === "call") {
          var callShift = st.shift || "";
          if (!callShift) {
            rooms.forEach(function (x) {
              if (x.name && nameKey(x.name) === nameKey(st.name) && x.shift) callShift = x.shift;
            });
          }
          calls.push({ name: st.name, shift: callShift, lastRoom: lastRoom || "", role: "call" });
        }
        addDeck(st, lastRoom, role);
        recordPerson(lastRoom, raw, st, [], role, refs);
      }
    }
    for (r = 5; r <= 43; r++) {
      var g = cell(cells, "G", r);
      var h = cell(cells, "H", r);
      if (/^MIDNIGHT/i.test(g)) {
        if (!isTimeLabel(h)) addLabelled(h, g, "midnight", ["G" + r, "H" + r]);
      }
      else if (/^#\d/.test(g)) { /* late stay handled below */ }
      else if (/^(CV|CCS|3N|2N|ENDO|ST 1|ST 2|STE 1|STE 2)$/i.test(g)) addLabelled(h, g, "breaker", ["G" + r, "H" + r]);
      else if (/^(MN CALL|HEART CALL)$/i.test(g)) {
        var who = h && !/^\d{3,4}/.test(h) ? h : cell(cells, "G", r + 1);
        if (who && !/^(MN CALL|HEART CALL|NT BREAKERS|ST BREAKERS|LATE STAY)/i.test(who)) {
          addLabelled(who, g, "call", ["G" + r, "H" + r]);
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
          var lsShift = lst.shift || "";
          rooms.forEach(function (x) {
            if (!x.name || x.closed || !namesMatch(x.name, lst.name)) return;
            if (nameKey(x.name) !== nameKey(lst.name)) lst.name = x.name;
            if (!lsShift && x.shift) lsShift = x.shift;
          });
          if (!lsShift) {
            onDeck.forEach(function (p) {
              if (p.name && namesMatch(p.name, lst.name) && p.shift) lsShift = p.shift;
            });
          }
          lateStays.push({
            n: parseInt(num[1], 10),
            name: lst.name,
            shift: lsShift,
            wave: lsWave,
            role: "latestay"
          });
          recordPerson("Late stay #" + num[1], h, lst, [], "latestay", ["G" + r, "H" + r]);
        }
      }
    }
    var placed = {};
    var placedDr = {};
    rooms.forEach(function (x) {
      if (x.name && !x.closed) {
        placed[nameKey(x.name)] = x;
        if (coreShift(x.shift) === "Dr") placedDr[nameKey(x.name)] = 1;
      }
    });
    onDeck = onDeck.filter(function (p) {
      if (!p.name) return false;
      var k = nameKey(p.name);
      if (!placed[k]) return true;
      if ((p.role === "midnight" || p.role === "call") && placedDr[k] && coreShift(p.shift) !== "Dr") return true;
      return false;
    });
    onDeck = mergeDeckPeople(onDeck);
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
      lateStays: lateStays, unsure: unsure, preview: preview, sheetPeople: sheetPeople,
      calls: calls
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
    var inSrna = false;
    var extraF = [];
    var r;

    function findPerson(date, name) {
      var k = dayKey(date);
      for (var i = 0; i < people.length; i++) {
        if (dayKey(people[i].date) !== k) continue;
        if (namesMatch(people[i].name, name)) return people[i];
      }
      return null;
    }

    function readAssign(assign) {
      var raw = String(assign || "").trim();
      var parts = raw.split(/\s*\/\s*/).map(function (x) { return x.trim(); }).filter(Boolean);
      var roomPart = "";
      var obTime = "";
      var br = false;
      var bareOb = false;
      parts.forEach(function (p) {
        if (/^BR$/i.test(p)) { br = true; return; }
        var ob = p.match(/^OB\s*(?:@?\s*)?(\d{1,2})\s*([ap])m?$/i);
        if (/^OB\b/i.test(p) && ob) { obTime = ob[1] + ob[2].toLowerCase(); return; }
        if (/^OB$/i.test(p)) { bareOb = true; return; }
        if (!roomPart) roomPart = p;
      });
      if (!parts.length && /^OB$/i.test(raw)) bareOb = true;
      return { raw: raw, roomPart: roomPart, obTime: obTime, br: br, bareOb: bareOb && !obTime };
    }

    function callBits(raw) {
      var s = String(raw || "").replace(/\s+/g, " ").trim();
      if (!s || !/(\d\s*[ap]?m?\s*[-–]\s*\d|#\s*\d|10p)/i.test(s)) return [];
      if (/please note|required|dial |password|hemo|trauma|total:|days$|evenings|general crna|crna call|resid|weekend srna|\(none\)/i.test(s)) return [];
      var chunks = /[A-Za-z]{3,}.*\/.*[A-Za-z]{3,}/.test(s) ? s.split(/\s*\/\s*/) : [s];
      var out = [];
      chunks.forEach(function (chunk) {
        if (!/(\d\s*[-–]\s*\d|#\s*\d|10p)/i.test(chunk)) return;
        var name = chunk;
        for (var pass = 0; pass < 3; pass++) {
          var next = name.replace(/^(?:10p\s*[-–]\s*\d{3,4}|\d{1,2}\s*[ap]?m?\s*[-–]\s*\d{1,2}\s*[ap]?m?)\s*/i, "");
          next = next.replace(/^#\s*\d+\s*:?\s*/, "").replace(/^[:\s]+/, "");
          if (next === name) break;
          name = next;
        }
        name = name.replace(/[:]+/g, " ").replace(/\s+/g, " ").trim();
        if (!/[A-Za-z]{3,}/.test(name)) return;
        var win = chunk.match(/(10p\s*[-–]\s*\d{3,4}|\d{1,2}\s*[ap]?m?\s*[-–]\s*\d{1,2}\s*[ap]?m?)/i);
        var shift = /10p|11p/i.test(chunk) ? "N" : (/2\s*[-–]\s*11|2\s*[-–]\s*7|3p\s*[-–]\s*6a|7-11p/i.test(chunk) ? "E" : "D");
        out.push({ name: name, shift: shift, label: win ? win[1].replace(/\s+/g, "") : "call" });
      });
      return out;
    }

    function addSrna(raw) {
      var s = String(raw || "").replace(/\s+/g, " ").trim();
      if (!s || /^\(none\)$/i.test(s)) return;
      if (/resid|crna call|general crna/i.test(s)) return;
      s = s.replace(/^(?:EVES|MN)\s*:\s*/i, "");
      var shift = "";
      var name = s;
      var m = s.match(/^(\d{1,2}\s*[ap]?m?\s*[-–]\s*\d{1,2}\s*[ap]?m?)\s+(.+)$/i);
      if (m) { shift = m[1].replace(/\s+/g, ""); name = m[2]; }
      else {
        m = s.match(/^([DdEeSsNnWwTt])\s+(.+)$/);
        if (m) { shift = m[1]; name = m[2]; }
      }
      var tr = name.match(/^(.*?)(\d{1,2}\s*[ap]?m?\s*[-–]\s*\d{1,2}\s*[ap]?m?)\s*$/i);
      if (tr && /[A-Za-z]{3,}/.test(tr[1])) {
        name = tr[1].trim();
        if (!shift) shift = tr[2].replace(/\s+/g, "");
      }
      name = name.replace(/[:]+/g, " ").trim();
      if (!/[A-Za-z]{3,}/.test(name) || isJunkStaff(name, shift)) return;
      addPerson({ date: currentDate, name: name, shift: shift, assign: "", role: "srna", student: true });
    }

    function addPerson(opts) {
      if (!opts || !opts.name || isJunkStaff(opts.name, opts.shift)) return;
      var parsed = readAssign(opts.assign);
      if (parsed.bareOb && !parsed.obTime) {
        var lateM = String(opts.shift || "").match(/(\d{1,2})\s*([ap])?/i);
        if (lateM) {
          var lh = parseInt(lateM[1], 10);
          var lap = (lateM[2] || "").toLowerCase();
          if (lap === "p" && lh < 12) lh += 12;
          if (!lap && lh > 0 && lh <= 7) lh += 12;
          if (lh >= 15) {
            parsed.obTime = lateM[1] + (lateM[2] || "p");
            parsed.bareOb = false;
          }
        }
      }
      var noted = /^(MRI\s*ST|ST\s*MRI)$/i.test(parsed.roomPart || parsed.raw);
      var loc = (!parsed.bareOb && parsed.roomPart && !parsed.br) ? normalizeRoom(parsed.roomPart) : null;
      if (noted) loc = null;
      var role = opts.role || "";
      if (parsed.br) role = "breaker";
      if (parsed.obTime && !role) role = "float";
      var prev = findPerson(opts.date, opts.name);
      if (prev) {
        var incomingRoom = !!(loc && loc.room) || parsed.bareOb || parsed.br;
        var prevRoom = !!(prev.room || prev.ob || prev.role === "breaker");
        if (incomingRoom && !prevRoom && role !== "call" && role !== "midnight" && role !== "srna") {
          if (nameKey(prev.name) === nameKey(opts.name)) prev.name = cleanName(opts.name);
          prev.shift = opts.shift || prev.shift;
          prev.assign = parsed.raw || prev.assign;
          prev.cat = loc && loc.cat;
          prev.room = loc && loc.room;
          prev.kind = opts.kind || prev.kind;
          prev.role = role;
          prev.at3p = !!parsed.obTime;
          prev.ob = parsed.bareOb && !parsed.obTime;
          prev.arrival = parsed.obTime || "";
          prev.noted = noted;
          prev.intended = parsed.obTime ? "OB" : (prev.intended || "");
        } else if (role === "call" && !prev.room && !prev.ob && prev.role !== "breaker" && prev.role !== "srna") {
          prev.role = "call";
          prev.shift = opts.shift || prev.shift;
          prev.assign = parsed.raw || prev.assign;
        }
        return;
      }
      people.push({
        date: opts.date, tower: opts.tower || currentTower,
        shift: opts.shift || "", name: cleanName(opts.name),
        assign: parsed.raw, cat: loc && loc.cat, room: loc && loc.room,
        closed: false, kind: opts.kind || breakKind(opts.shift),
        early: !!opts.early, orient: !!opts.orient, student: !!opts.student,
        role: role, at3p: !!parsed.obTime, ob: parsed.bareOb && !parsed.obTime,
        arrival: parsed.obTime || "", noted: noted,
        intended: parsed.obTime ? "OB" : (opts.intended || "")
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
        inSrna = false;
        continue;
      }
      if (/South Tower/i.test(b) || /South Tower/i.test(a)) {
        currentTower = "ST";
        lastShift = "";
        pendingCall = false;
        inSrna = false;
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

      if (/Weekend SRNA/i.test(e)) inSrna = true;
      if (inSrna && e && !/Weekend SRNA/i.test(e)) {
        if (/resid|crna call|general crna/i.test(e)) inSrna = false;
        else addSrna(e);
      }

      if (/CRNA Call/i.test(e) || /CRNA Call/i.test(b)) pendingCall = true;
      if (pendingCall && e && !/CRNA Call/i.test(e)) {
        var bare = String(e).replace(/\s+/g, " ").trim();
        if (/^[A-Za-z][A-Za-z.'’-]+(?:\s+[A-Za-z][A-Za-z.'’-]+)+$/.test(bare) && !/liver|resid|call|note|days|evenings|weekend|srna|please|total|tower|hemo|trauma/i.test(bare)) {
          addPerson({ date: currentDate, name: bare, shift: "", assign: "call", role: "call" });
          pendingCall = false;
        } else if (callBits(e).length) pendingCall = false;
      }
      if (!inSrna) callBits(e).forEach(function (bit) {
        addPerson({ date: currentDate, name: bit.name, shift: bit.shift, assign: bit.label, role: "call" });
      });

      if (f && looksLikeShiftCell(f.split(/\s+/)[0]) && /[A-Za-z]{3,}/.test(f)) {
        extraF.push({ date: currentDate, raw: f });
      }

      if (/^shift$/i.test(a) || /^CRNA$/i.test(b)) continue;
      if (/^(check hemacue|trauma rm)/i.test(a)) continue;
      if (a && (!b || /^CRNA$/i.test(b))) {
        var glued = String(a).trim().match(/^(\S+)\s+([A-Za-z].+)$/);
        if (glued && (looksLikeShiftCell(glued[1]) || parseShiftLabel(glued[1].replace(/\s+/g, "")))) {
          a = glued[1];
          b = glued[2];
        }
      }

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
    var datesSeen = [];
    people.forEach(function (p) {
      var dk = dayKey(p.date);
      if (dk && datesSeen.indexOf(dk) < 0) datesSeen.push(dk);
    });
    function dateNum(k) {
      var p = String(k || "").split("-").map(function (x) { return parseInt(x, 10); });
      if (p.length < 3 || !p[0]) return 0;
      return p[0] * 10000 + p[1] * 100 + p[2];
    }
    var pickDay = "";
    if (today && datesSeen.indexOf(today) >= 0) pickDay = today;
    else {
      var nowN = today ? dateNum(today) : 0;
      var future = datesSeen.filter(function (d) { return dateNum(d) >= nowN; }).sort(function (a, b) { return dateNum(a) - dateNum(b); });
      if (future.length) pickDay = future[0];
      else pickDay = datesSeen.slice().sort(function (a, b) { return dateNum(b) - dateNum(a); })[0] || "";
    }
    var dayPeople = people.filter(function (p) {
      return pickDay && dayKey(p.date) === pickDay;
    });
    today = pickDay || today;

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
      if (p.role === "srna") { take(p, { role: "srna", lastRoom: "SRNA" }); return; }
      if ((p.at3p || p.arrival) && !(p.cat && p.room)) { take(p, { role: p.role === "breaker" ? "breaker" : "float", lastRoom: "OB @" + (p.arrival || "3p"), intended: "OB" }); return; }
      if (p.ob || (p.room && /^OB$/i.test(p.room))) { obQueue.push(p); return; }
      if (p.noted) { take(p, { role: "float", lastRoom: p.assign || "MRI" }); return; }
      if (p.cat && p.room) {
        rooms.push({ cat: p.cat, room: p.room, shift: p.shift, name: p.name, closed: false, kind: p.kind, student: !!p.student });
        used[nameKey(p.name)] = 1;
        return;
      }
      if (p.assign && !/^BR$/i.test(p.assign) && !/^call$/i.test(p.assign) && p.role !== "call" && p.role !== "srna" && !p.arrival && !p.noted) unmatched.push({ room: p.assign, staff: p.name });
      take(p, { role: p.role || "float" });
    });
    var obSlots = ["OB 1", "OB 2"];
    obQueue.forEach(function (p, i) {
      if (used[nameKey(p.name)]) return;
      if (i < obSlots.length) {
        p.cat = "fbc";
        p.room = obSlots[i];
        rooms.push({ cat: "fbc", room: obSlots[i], shift: p.shift, name: p.name, closed: false, kind: p.kind, student: !!p.student });
        used[nameKey(p.name)] = 1;
      } else {
        p.room = "";
        p.cat = "";
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
      onDeck: onDeck,
      sheetPeople: dayPeople.filter(function (p) { return p && p.name; }).map(function (p) {
        var locs = (p.cat && p.room) ? [{ cat: p.cat, room: p.room }] : [];
        return {
          roomRaw: p.assign || p.room || p.role || "",
          staffRaw: (p.shift ? p.shift + " " : "") + p.name,
          name: p.name, shift: p.shift || "", student: !!p.student, kind: p.kind || "none",
          locs: locs, role: p.role || ""
        };
      })
    };
  }

  function parseSplitRoster(cells) {
    var people = [];
    var calls = [];
    var seen = {};
    var lastN = "";
    var lastS = "";
    var callWave = "";
    function addRoster(shiftRaw, nameRaw, tower) {
      var nm = flipCommaName(nameRaw);
      if (!nm || !/[A-Za-z]{3,}/.test(nm)) return;
      if (/crna|assign|shift|north|south|orientee|resident|target|call/i.test(nm)) return;
      var lab = shiftRaw ? parseShiftLabel(String(shiftRaw).replace(/\s+/g, "")) : null;
      var shift = lab ? lab.shift : "";
      var st = parseStaff((shift ? shift + " " : "") + nm);
      if (!st || !st.name || isJunkStaff(st.name, st.shift)) return;
      var k = nameKey(st.name);
      if (seen[k]) return;
      seen[k] = 1;
      people.push({
        name: st.name, shift: st.shift || shift, kind: st.kind || "none",
        tower: tower, role: "float", student: !!st.orient, orient: !!st.orient
      });
    }
    var r;
    for (r = 1; r <= 80; r++) {
      var a = cell(cells, "A", r), b = cell(cells, "B", r);
      var e = cell(cells, "E", r), f = cell(cells, "F", r);
      var h = cell(cells, "H", r), i = cell(cells, "I", r);
      if (/^shift$/i.test(a)) lastN = "";
      if (/^shift$/i.test(e)) lastS = "";
      if (b && !/^CRNA/i.test(b) && !/^shift$/i.test(b)) {
        if (a && parseShiftLabel(String(a).replace(/\s+/g, ""))) lastN = a;
        if (lastN || (a && parseShiftLabel(String(a).replace(/\s+/g, "")))) addRoster(lastN || a, b, "NT");
      }
      if (f && !/^CRNA/i.test(f) && !/^shift$/i.test(f)) {
        if (e && parseShiftLabel(String(e).replace(/\s+/g, ""))) lastS = e;
        if (lastS || (e && parseShiftLabel(String(e).replace(/\s+/g, "")))) addRoster(lastS || e, f, "ST");
      }
      if (/general call/i.test(h)) callWave = "General";
      else if (/heart call/i.test(h)) callWave = "Heart";
      else if (/3:30/.test(h)) callWave = "1530";
      else if (/5:30/.test(h)) callWave = "1730";
      else if (/7:30/.test(h) && /call/i.test(h)) callWave = "1930";
      if ((/^1st$/i.test(h) || /^GC\s*\d/i.test(h)) && i && /[A-Za-z]{3,}/.test(i)) {
        var cn = flipCommaName(i);
        if (cn) calls.push({ name: cn, shift: "", lastRoom: (callWave + " " + h.replace(/\s+/g, "")).trim(), role: "call" });
      }
    }
    return {
      kind: "split", date: "", rooms: [], closed: [], unmatched: [],
      people: people, onDeck: [], calls: calls, unsure: [], preview: [], sheetPeople: []
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
    if (/CRNA NORTH/i.test(blob) && /STAFF TARGETS/i.test(blob)) out = parseSplitRoster(cells);
    else if (blob.indexOf("CRNA SCHEDULE") >= 0 && blob.indexOf("NORTH TOWER") >= 0) out = parseWeekday(cells);
    else if (blob.indexOf("NORTH TOWER") >= 0 && blob.indexOf("SOUTH TOWER") >= 0) out = parseWeekend(cells);
    else out = { kind: "unknown", date: "", dates: [], rooms: [], closed: [], unmatched: [], people: [], onDeck: [], unsure: [], preview: [], sheetPeople: [], runners: {} };
    out.file = fileName || "";
    out.cells = cells;
    out.onDeck = out.onDeck || [];
    out.unsure = out.unsure || [];
    out.preview = out.preview || [];
    out.sheetPeople = out.sheetPeople || [];
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
    if (p.role === "wbf") {
      var where = p.intended || p.lastRoom || "";
      if (!where || /^WBF$/i.test(where)) return "WBF";
      return "WBF " + String(where).replace(/([A-Za-z])(\d)/g, "$1 $2");
    }
    if (p.role === "resident") return p.lastRoom || "OB";
    if (p.role === "srna") return "SRNA";
    if (p.role === "evening") return "eves";
    if (p.role === "night") return p.lastRoom || "night";
    if (p.role === "shift") return "";
    if (parseShiftLabel(p.lastRoom)) return "";
    if (p.role === "unplaced" && p.lastRoom) return String(p.lastRoom).replace(/([A-Za-z])(\d)/g, "$1 $2");
    if (p.role === "freed" && p.lastRoom) return "last " + p.lastRoom;
    if (p.role === "breaker") return p.lastRoom ? "BR " + p.lastRoom : "breaker";
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
    return '<span class="staff-line staff-move" data-staff-cat="' + catId + '" data-staff-room="' + String(room).replace(/"/g, "") + '">' + shiftPillHtml(rec.shift) + '<span class="staff-name' + size + (gone ? " gone" : "") + '">' + nm + "</span>" + studentMark(hasStudent(rec)) + start + "</span>";
  }

  function roomBreakKind(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || rec.closed || !rec.name) return "none";
    return rec.kind || breakKind(rec.shift);
  }

  function tintRoomBtn(btn, catId, room) {
    if (!btn) return;
    btn.classList.remove("kind-late", "kind-dinner");
    var afterClose = false;
    try {
      if (typeof currentWindow === "number" && currentWindow === 3) afterClose = true;
      else if (hospitalMins() >= 15 * 60 + 30) afterClose = true;
    } catch (e) {}
    if (!afterClose) return;
    var k = roomBreakKind(catId, room);
    if (k === "late" || k === "dinner") btn.classList.add("kind-" + k);
  }

  g.tintRoomBtn = tintRoomBtn;
  g.roomBreakKind = roomBreakKind;

  function hospitalAfterClose() {
    try {
      if (typeof currentWindow === "number" && currentWindow === 3) return true;
      if (typeof hospitalMins === "function" && hospitalMins() >= 15 * 60 + 30) return true;
    } catch (e) {}
    return false;
  }

  function defaultStaffView() {
    try {
      if (typeof currentRole !== "undefined" && currentRole === "crna" && myAssignType === "breaker") {
        return hospitalAfterClose() ? "jobs" : "board";
      }
    } catch (e) {}
    return "board";
  }

  function getStaffView() {
    try {
      var v = sessionStorage.getItem("anespresso_runner_staffview_v1");
      if (v === "board" || v === "jobs") return v;
    } catch (e) {}
    return defaultStaffView();
  }

  function applyStaffView() {
    var jobs = false;
    try {
      if (typeof currentRole !== "undefined" && currentRole === "crna") jobs = getStaffView() === "jobs";
    } catch (e) {}
    document.body.classList.toggle("staff-view-jobs", jobs);
    document.body.classList.toggle("staff-view-board", !jobs);
  }

  function setStaffView(v) {
    if (v !== "board" && v !== "jobs") return;
    try { sessionStorage.setItem("anespresso_runner_staffview_v1", v); } catch (e) {}
    applyStaffView();
    try { if (typeof renderBreakerJobList === "function") renderBreakerJobList(); } catch (e) {}
    try { if (typeof renderMySitePanel === "function") renderMySitePanel(); } catch (e) {}
  }

  function staffViewToggleHtml() {
    var v = getStaffView();
    return '<div class="staff-view-toggle" role="tablist">' +
      '<button type="button" class="staff-view-btn' + (v === "board" ? " on" : "") + '" onclick="setStaffView(\'board\')">Board</button>' +
      '<button type="button" class="staff-view-btn' + (v === "jobs" ? " on" : "") + '" onclick="setStaffView(\'jobs\')">Jobs</button>' +
      "</div>";
  }

  function collectWindowDue() {
    var items = [];
    if (hospitalAfterClose()) {
      ["late", "dinner"].forEach(function (kind) {
        collectBreakQueue(kind).forEach(function (it) {
          if (it && it.room && it.cat) items.push(it);
        });
      });
      return items;
    }
    var w = typeof currentWindow === "number" ? currentWindow : 0;
    if (typeof CATEGORIES === "undefined") return items;
    CATEGORIES.forEach(function (cat) {
      var rooms = [];
      try { rooms = typeof activeRooms === "function" ? activeRooms(cat) : (cat.rooms || []); } catch (e) { rooms = cat.rooms || []; }
      rooms.forEach(function (room) {
        var rec = occupantOf(cat.id, room);
        if (!rec || !rec.name) return;
        if (rec.shift && !stillInHouse(rec.shift)) return;
        var given = false;
        try { given = !!(state[w] && state[w][cat.id] && state[w][cat.id][room]); } catch (e) {}
        items.push({
          name: rec.name,
          shift: rec.shift || "",
          kind: "window",
          cat: cat.id,
          room: room,
          given: given
        });
      });
    });
    items.sort(function (a, b) {
      if (a.given !== b.given) return a.given ? 1 : -1;
      return String(a.room).localeCompare(String(b.room));
    });
    return items;
  }

  function staffJobsHtml() {
    var urgent = [];
    try {
      Object.keys(sitePrefs || {}).forEach(function (room) {
        if (sitePrefs[room] && sitePrefs[room].bathroom) urgent.push(room);
      });
    } catch (e) {}
    var items = collectWindowDue();
    var due = items.filter(function (it) { return !it.given; });
    var had = items.filter(function (it) { return it.given; });
    function jobRow(label, sub, onclick, kind) {
      var cls = "job-row";
      if (kind === "late") cls += " job-late";
      if (kind === "dinner") cls += " job-dinner";
      if (kind === "urgent") cls += " job-urgent";
      if (kind === "had") cls += " job-info";
      var inner = '<span class="job-main">' + label + "</span>" +
        (sub ? '<span class="job-sub">' + sub + "</span>" : "");
      if (!onclick) return '<div class="' + cls + '">' + inner + "</div>";
      return '<button type="button" class="' + cls + '" onclick="' + onclick + '">' + inner + "</button>";
    }
    function winSub(it) {
      if (it.kind === "dinner") return "dinner";
      if (it.kind === "late") return "late";
      try { return (typeof WIN_LABELS !== "undefined" && WIN_LABELS[currentWindow]) || "due"; } catch (e) { return "due"; }
    }
    var html = "";
    if (urgent.length) {
      html += '<div class="job-sec">Need a break now</div>';
      urgent.forEach(function (room) {
        var nm = "";
        try { nm = (typeof staffLastName === "function") ? (staffLastName(room) || "") : ""; } catch (e) {}
        html += jobRow(escHtml(room) + (nm ? " · " + escHtml(nm) : ""), "Urgent", "promptBathroomComplete('" + String(room).replace(/'/g, "\\'") + "')", "urgent");
      });
    }
    if (due.length) {
      html += '<div class="job-sec">Due this window · ' + due.length + "</div>";
      due.forEach(function (it) {
        var label = (it.name ? escHtml(chipName(it.name)) + " · " : "") + escHtml(it.room);
        html += jobRow(label, winSub(it), "toggleRoom('" + it.cat + "','" + String(it.room).replace(/'/g, "\\'") + "')", it.kind === "dinner" ? "dinner" : (it.kind === "late" ? "late" : ""));
      });
    }
    if (had.length) {
      html += '<div class="job-sec">Already had it · ' + had.length + "</div>";
      had.forEach(function (it) {
        var label = (it.name ? escHtml(chipName(it.name)) + " · " : "") + escHtml(it.room);
        html += jobRow(label, "had it", "", "had");
      });
    }
    if (!urgent.length && !due.length) {
      html += '<div class="job-empty">' + (hospitalAfterClose() ? "No late or dinner breaks due. Switch to Board to see rooms." : "No rooms due this window. Switch to Board to see everyone.") + "</div>";
    }
    return html;
  }

  function paintCrnaJobs() {
    var panel = document.getElementById("my-site-panel");
    if (!panel) return;
    var slot = document.getElementById("staff-jobs-slot");
    if (getStaffView() !== "jobs") {
      if (slot) slot.remove();
      return;
    }
    if (!slot) {
      slot = document.createElement("div");
      slot.id = "staff-jobs-slot";
      slot.className = "staff-jobs-slot";
      panel.appendChild(slot);
    }
    slot.innerHTML = staffJobsHtml();
  }

  function publishStaff() {
    try {
      window.roomStaff = g.roomStaff || {};
      window.onDeck = g.onDeck || [];
      window.reliefPlan = g.reliefPlan || {};
      window.assignmentMeta = g.assignmentMeta || null;
    } catch (e) {}
  }

  function hydrateStaff() {
    try {
      if (window.roomStaff && typeof window.roomStaff === "object") g.roomStaff = window.roomStaff;
      if (Array.isArray(window.onDeck)) g.onDeck = window.onDeck;
      if (window.reliefPlan && typeof window.reliefPlan === "object") g.reliefPlan = window.reliefPlan;
      if ("assignmentMeta" in window) {
        g.assignmentMeta = window.assignmentMeta || null;
        if (g.assignmentMeta && g.assignmentMeta.lateStays) g.lateStays = g.assignmentMeta.lateStays;
        else if (!g.assignmentMeta) g.lateStays = [];
        if (g.assignmentMeta && g.assignmentMeta.calls) g.calls = g.assignmentMeta.calls;
        else if (!g.assignmentMeta) g.calls = [];
      }
    } catch (e) {}
  }

  function applySplitRoster(res) {
    g.onDeck = g.onDeck || [];
    g.calls = g.calls || [];
    var added = 0;
    (res.people || []).forEach(function (p) {
      if (!p || !p.name) return;
      var hits = [];
      try { hits = boardHits(p.name); } catch (e) { hits = []; }
      if (hits.length) {
        hits.forEach(function (h) {
          if (h.where === "room" && g.roomStaff[h.cat] && g.roomStaff[h.cat][h.room] && !g.roomStaff[h.cat][h.room].shift && p.shift) {
            g.roomStaff[h.cat][h.room].shift = p.shift;
            g.roomStaff[h.cat][h.room].kind = p.kind || breakKind(p.shift);
          }
          if (h.where === "deck") {
            (g.onDeck || []).forEach(function (d) {
              if (d && namesMatch(d.name, p.name) && !d.shift && p.shift) {
                d.shift = p.shift;
                d.kind = p.kind || breakKind(p.shift);
              }
            });
          }
        });
        return;
      }
      g.onDeck.push({
        name: p.name, shift: p.shift || "", kind: p.kind || "none",
        lastRoom: p.tower || "", role: p.orient ? "resident" : "float",
        student: !!p.orient, intended: ""
      });
      added++;
    });
    (res.calls || []).forEach(function (c) {
      if (!c || !c.name) return;
      var exists = (g.calls || []).some(function (x) { return x && namesMatch(x.name, c.name) && x.lastRoom === c.lastRoom; });
      if (exists) return;
      g.calls.push({ name: cleanName(c.name), shift: c.shift || "", lastRoom: c.lastRoom || "", role: "call" });
    });
    publishStaff();
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    try { renderOnDeck(); } catch (e) {}
    if (res) res.added = added;
    var msg = added ? ("Roster merged — " + added + " not on the board yet") : "Roster matches the board";
    try { if (typeof showToast === "function") showToast(msg); } catch (e) {}
  }

  function applyAssignmentResult(res) {
    if (!res) return;
    if (res.kind === "split") {
      applySplitRoster(res);
      return;
    }
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
        (res.rooms || []).forEach(function (r) {
          if (!r.cat || !r.room || !r.name || r.closed) return;
          if (!g.roomStaff[r.cat]) g.roomStaff[r.cat] = {};
          if (g.roomStaff[r.cat][r.room] && g.roomStaff[r.cat][r.room].name) return;
          g.roomStaff[r.cat][r.room] = {
            name: r.name, shift: r.shift, kind: r.kind, closed: false,
            student: !!r.student, firstCase: r.firstCase || ""
          };
          openN++;
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
          intended: p.intended || ((p.role === "extra" || p.role === "offsite") ? (p.lastRoom || "") : ""),
          student: !!p.student, firstCase: p.firstCase || "", freedAt: now
        };
      }).filter(function (p) { return p.name && !isJunkStaff(p.name, p.shift); });
      g.lateStays = (res.lateStays || []).map(function (p) {
        return { n: p.n, name: cleanName(p.name), shift: p.shift || "", wave: p.wave, role: "latestay" };
      });
      g.calls = (res.calls || []).map(function (p) {
        return { name: cleanName(p.name), shift: p.shift || "", lastRoom: p.lastRoom || "", role: "call" };
      });
      g.lastSheet = { cells: res.cells || {}, date: res.date, file: res.fileName || res.file || "", kind: res.kind };
      try { localStorage.setItem("anespresso_runner_sheet_v1", JSON.stringify(g.lastSheet)); } catch (e) {}
      g.assignmentMeta = {
        date: res.date, kind: res.kind, file: res.fileName || "", pos: res.pos,
        lateN: (res.late || []).length, dinnerN: (res.dinner || []).length,
        roomsN: openN, closedN: closedN,
        unmatchedN: (res.unmatched || []).length, deckN: g.onDeck.length, appliedAt: now,
        lateStays: g.lateStays,
        calls: g.calls,
        runners: res.runners || null,
        cleared: false
      };
      g.reliefPlan = {};
      publishStaff();
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
    var tag = p.given ? "had it" : (p.role === "breaker" ? "breaker" : (p.room || ""));
    var last = tag ? '<span class="staff-last">' + tag + "</span>" : "";
    var num = p.given ? "" : '<span class="breakq-n">' + n + "</span>";
    return '<button type="button" class="breakq-item' + (p.given ? " given" : "") + '"' +
      (p.cat ? ' data-qcat="' + p.cat + '"' : "") +
      (p.room ? ' data-qroom="' + p.room + '"' : "") + ">" +
      num + shiftPillHtml(p.shift) +
      '<span class="staff-name">' + chipName(p.name || p.room || "") + "</span>" + last +
      "</button>";
  }

  function queueBlockHtml(label, items) {
    var due = items.filter(function (p) { return !p.given; });
    var had = items.filter(function (p) { return p.given; });
    if (!items.length) {
      return '<div class="late-board-row"><strong>' + label + "</strong> none in house</div>";
    }
    var rows = due.map(function (p, i) { return queueRowHtml(p, i + 1); }).join("");
    var hadRows = had.map(function (p) { return queueRowHtml(p, 0); }).join("");
    return '<div class="late-board-row"><strong>' + label + "</strong> " + due.length + " due" +
      (had.length ? " · " + had.length + " had it" : "") + "</div>" +
      (rows ? '<div class="breakq-list">' + rows + "</div>" : "") +
      (hadRows ? '<div class="breakq-had-label">Already had it</div><div class="breakq-list breakq-had">' + hadRows + "</div>" : "");
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
    return deckBucket(p) === "out";
  }

  function deckBucket(p) {
    if (!p) return "out";
    if (p.role === "call") return "now";
    var h = hospitalHour();
    var t = personHours(p);
    var start = t.start;
    var end = t.end;
    if (t.unknown) {
      if (start > h) return "later";
      return "now";
    }
    if (end > 24) {
      var mornEnd = end - 24;
      if (h >= start || h < mornEnd) return "now";
      return "out";
    }
    if (start > h) return "later";
    if (h >= end) return "out";
    return "now";
  }

  function idleMs(p) {
    var now = Date.now();
    if (p && p.freedAt) return Math.max(0, now - p.freedAt);
    var start = shiftStartHour(p && p.shift);
    if (start == null || start >= 99) return 0;
    var d = new Date();
    try { if (typeof hospitalNow === "function") d = hospitalNow(); } catch (e) {}
    var hr = start > 24 ? start - 24 : start;
    var t = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Math.floor(hr), 0, 0, 0).getTime();
    if (start >= 19 && hospitalHour() < 12) t -= 86400000;
    return Math.max(0, now - t);
  }

  function sortByIdle(arr) {
    arr.sort(function (a, b) {
      var d = idleMs(b) - idleMs(a);
      if (d) return d;
      return String(chipName(a && a.name)).localeCompare(String(chipName(b && b.name)));
    });
    return arr;
  }

  function longestIdleDeck(deck) {
    var now = (deck || []).filter(function (p) {
      return deckBucket(p) === "now" && !isLeavingSoon(p.shift, p);
    });
    if (!now.length) return null;
    sortByIdle(now);
    return now[0] || null;
  }

  function todayRoster() {
    var out = [];
    var seen = {};
    function add(name, shift, kind, cat, room, role) {
      if (!name) return;
      var k = nameKey(name);
      if (!k || seen[k]) return;
      seen[k] = 1;
      out.push({ name: name, shift: shift || "", kind: kind, cat: cat || "", room: room || "", role: role || "" });
    }
    var cats = (typeof CATEGORIES !== "undefined") ? CATEGORIES : [];
    cats.forEach(function (c) {
      var staff = (g.roomStaff || {})[c.id] || {};
      (c.rooms || []).forEach(function (r) {
        var rec = staff[r];
        if (rec && rec.name && !rec.closed) add(rec.name, rec.shift, "room", c.id, r, "");
      });
    });
    (g.onDeck || []).forEach(function (p) {
      if (p && p.name) add(p.name, p.shift, "deck", "", p.lastRoom || "", p.role || "");
    });
    out.sort(function (a, b) {
      return lastName(a.name).localeCompare(lastName(b.name));
    });
    return out;
  }

  function findStaffByName(name) {
    var k = nameKey(name);
    if (!k) return null;
    var list = todayRoster();
    var i, hits;
    for (i = 0; i < list.length; i++) {
      if (nameKey(list[i].name) === k) return list[i];
    }
    var lastK = nameKey(lastName(name));
    if (!lastK) return null;
    hits = [];
    for (i = 0; i < list.length; i++) {
      if (nameKey(lastName(list[i].name)) === lastK) hits.push(list[i]);
    }
    return hits.length === 1 ? hits[0] : null;
  }

  function splitDeck(list) {
    var now = [], later = [], gone = [];
    (list || []).forEach(function (p) {
      var b = deckBucket(p);
      if (b === "out") gone.push(p);
      else if (b === "later") later.push(p);
      else now.push(p);
    });
    sortByIdle(now);
    sortByShiftLen(later);
    sortByShiftLen(gone);
    return { now: now, later: later, gone: gone };
  }

  function lateStayByName() {
    var map = {};
    (g.lateStays || (g.assignmentMeta && g.assignmentMeta.lateStays) || []).forEach(function (p) {
      if (p && p.name) map[nameKey(p.name)] = p;
    });
    return map;
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
    var runner = false;
    try { runner = typeof currentRole !== "undefined" && currentRole === "runner"; } catch (e) {}
    if (!runner) {
      if (card) card.remove();
      return;
    }
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
    var parts = splitDeck(list);
    var here = parts.now;
    var later = parts.later;
    var gone = parts.gone;
    var sel = g.selectedDeck || null;
    function chipHtml(p, mark) {
      var key = nameKey(p.name);
      var tag = deckTag(p);
      var last = tag ? '<span class="staff-last">' + tag + "</span>" : "";
      var on = sel === key ? " selected" : "";
      return '<button type="button" class="ondeck-chip' + (mark ? " " + mark : "") + on + '" data-deck="' + key + '">' +
        shiftPillHtml(p.shift) + '<span class="staff-name">' + chipName(p.name) + "</span>" + studentMark(hasStudent(p)) + last +
        (mark === "out" ? '<span class="staff-last">out</span>' : (mark === "later" ? '<span class="staff-last">later</span>' : "")) +
        "</button>";
    }
    var hereChips = here.map(function (p) { return chipHtml(p, ""); }).join("");
    var laterChips = later.map(function (p) { return chipHtml(p, "later"); }).join("");
    var goneChips = gone.map(function (p) { return chipHtml(p, "out"); }).join("");
    var laterOpen = !!g.deckLaterOpen;
    var laterBtn = later.length
      ? '<button type="button" class="ondeck-out-toggle" data-fold="later" id="ondeck-later-toggle" aria-expanded="' + (laterOpen ? "true" : "false") + '">' +
        (laterOpen ? "Hide later · " : "Later · ") + later.length + (laterOpen ? " ▴" : " ▾") + "</button>" +
        (laterOpen ? '<div class="ondeck-grid ondeck-later-grid">' + laterChips + "</div>" : "")
      : "";
    var open = !!g.deckOutOpen;
    var outBtn = gone.length
      ? '<button type="button" class="ondeck-out-toggle" data-fold="out" id="ondeck-out-toggle" aria-expanded="' + (open ? "true" : "false") + '">' +
        (open ? "Hide out · " : "Out · ") + gone.length + (open ? " ▴" : " ▾") + "</button>" +
        (open ? '<div class="ondeck-grid ondeck-out-grid">' + goneChips + "</div>" : "")
      : "";
    var canMove = canMoveStaff();
    var held = heldFrom();
    var hint = !canMove
      ? ""
      : (g.moveHint
        ? '<div class="ondeck-hint">' + g.moveHint + "</div>"
        : (g.fillTarget
          ? '<div class="ondeck-hint">Tap who goes in ' + g.fillTarget.room + "</div>"
          : (sel
            ? '<div class="ondeck-hint">Tap a room to place them · tap the chip again to cancel</div>'
            : '<div class="ondeck-hint">Longest idle first · tap someone, then tap a room</div>')));
    card.innerHTML =
      '<div class="cat-header-row"><div class="cup-indicator">☕</div><div class="cat-info">' +
      '<div class="cat-name">On deck</div>' +
      '<div class="cat-full-name">Free now · longest idle first</div></div>' +
      '<div class="cat-progress-label"><div class="cat-pct">' + here.length + '</div>' +
      '<div class="cat-count">now</div></div></div>' +
      hint +
      '<div class="ondeck-grid">' + (hereChips || '<span class="roster-empty">Nobody free right now</span>') + "</div>" +
      laterBtn +
      outBtn;
    function inBucket(arr, key) {
      return (arr || []).some(function (p) { return nameKey(p.name) === key; });
    }
    function toggleFold(which) {
      if (which === "later") {
        g.deckLaterOpen = !g.deckLaterOpen;
        if (!g.deckLaterOpen && sel && inBucket(later, sel)) g.selectedDeck = null;
      } else if (which === "out") {
        g.deckOutOpen = !g.deckOutOpen;
        if (!g.deckOutOpen && sel && inBucket(gone, sel)) g.selectedDeck = null;
      }
      document.body.classList.toggle("assigning", !!(g.selectedDeck || g.heldMove));
      renderOnDeck();
    }
    if (!canMove) {
      g.selectedDeck = null;
      g.fillTarget = null;
    }
    card.querySelectorAll("[data-deck]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        if (!canMoveStaff()) return;
        var k = el.getAttribute("data-deck");
        if (g.fillTarget && g.fillTarget.cat && g.fillTarget.room) {
          g.selectedDeck = k;
          var tgt = g.fillTarget;
          g.fillTarget = null;
          assignSelectedTo(tgt.cat, tgt.room);
          return;
        }
        g.heldMove = null;
        g.selectedDeck = g.selectedDeck === k ? null : k;
        document.body.classList.toggle("assigning", !!(g.selectedDeck || g.heldMove));
        renderOnDeck();
        renderShiftChange();
      };
    });
    card.querySelectorAll("[data-fold]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        toggleFold(el.getAttribute("data-fold"));
      };
    });
    document.body.classList.toggle("assigning", !!(g.selectedDeck || g.heldMove));
  }

  function rememberClose(catId, room, rec) {
    g.undoClose = {
      cat: catId,
      room: room,
      rec: rec && rec.name ? {
        name: rec.name, shift: rec.shift || "", kind: rec.kind || "none",
        student: !!rec.student, closed: false
      } : null,
      ts: Date.now()
    };
    try { renderUndoBar(); } catch (e) {}
  }

  function undoLastClose() {
    var u = g.undoClose;
    if (!u || !u.cat || !u.room) return;
    g.roomStaff = g.roomStaff || {};
    g.roomStaff[u.cat] = g.roomStaff[u.cat] || {};
    if (u.rec && u.rec.name) {
      var k = nameKey(u.rec.name);
      g.onDeck = (g.onDeck || []).filter(function (p) { return nameKey(p.name) !== k; });
      g.roomStaff[u.cat][u.room] = {
        name: u.rec.name, shift: u.rec.shift, kind: u.rec.kind,
        closed: false, student: !!u.rec.student
      };
    }
    try { applyRoomActive(u.cat, u.room, true); } catch (e) {}
    var label = (u.rec && u.rec.name) ? chipName(u.rec.name) + " back in " + u.room : u.room + " reopened";
    g.undoClose = null;
    try { persistStaff(); } catch (e) {
      try { if (typeof saveShared === "function") saveShared(); } catch (e2) {}
      try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e2) {}
      renderOnDeck();
      paintWkndStaff();
    }
    renderUndoBar();
    try { if (typeof showToast === "function") showToast(label); } catch (e) {}
  }

  function renderUndoBar() {
    var u = g.undoClose;
    var el = document.getElementById("undo-close-bar");
    if (!u || !u.rec) {
      if (el) el.remove();
      var desk = document.getElementById("desk-undo-btn");
      if (desk) desk.style.display = "none";
      return;
    }
    if (!el) {
      el = document.createElement("button");
      el.id = "undo-close-bar";
      el.type = "button";
      var app = document.getElementById("app") || document.body;
      app.appendChild(el);
    }
    el.textContent = "Undo · " + chipName(u.rec.name) + " back in " + u.room;
    el.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      undoLastClose();
    };
    var desk = document.getElementById("desk-undo-btn");
    if (desk) {
      desk.style.display = "";
      desk.textContent = "Undo · " + chipName(u.rec.name) + " → " + u.room;
    }
  }

  function pushOnDeckFromRoom(catId, room) {
    var rec = g.roomStaff && g.roomStaff[catId] && g.roomStaff[catId][room];
    if (!rec || !rec.name) return;
    rememberClose(catId, room, rec);
    g.onDeck = g.onDeck || [];
    var k = nameKey(rec.name);
    g.onDeck = g.onDeck.filter(function (p) { return nameKey(p.name) !== k; });
    g.onDeck.unshift({
      name: rec.name, shift: rec.shift || "", kind: rec.kind || "none",
      lastRoom: room, role: "freed", student: !!rec.student, freedAt: Date.now()
    });
    delete g.roomStaff[catId][room];
    publishStaff();
    renderOnDeck();
  }

  function closeRoomIfEmpty(cat, room) {
    if (!cat || !room) return false;
    if (occupantOf(cat, room)) return false;
    try { deactivateRoomNow(cat, room); } catch (e) {}
    return true;
  }

  function sendOccupantToDeck(cat, room) {
    var rec = occupantOf(cat, room);
    if (!rec) return false;
    pushOnDeckFromRoom(cat, room);
    deactivateRoomNow(cat, room);
    return true;
  }

  function cloneStaffRec(rec) {
    if (!rec) return { name: "", shift: "", kind: "none", closed: false };
    return {
      name: rec.name, shift: rec.shift || "", kind: rec.kind || "none",
      closed: false, student: !!rec.student, firstCase: rec.firstCase || "",
      early: !!rec.early, orient: !!rec.orient
    };
  }

  function heldFrom() {
    if (g.heldMove && g.heldMove.cat && g.heldMove.room) {
      return {
        kind: "room", cat: g.heldMove.cat, room: g.heldMove.room,
        intent: g.heldMove.intent || "place", closeFrom: !!g.heldMove.closeFrom
      };
    }
    if (g.selectedDeck) return { kind: "deck", key: g.selectedDeck, intent: "place" };
    return null;
  }

  function clearHeld() {
    g.selectedDeck = null;
    g.heldMove = null;
    g.moveHint = "";
    try { document.body.classList.remove("assigning"); } catch (e) {}
    refreshMoveHints();
  }

  function refreshMoveHints() {
    var desk = document.querySelector("#weekend-overlay.show .sites-fn-hint");
    if (desk && g.sitesFn === "rooms") {
      desk.textContent = g.moveHint || "Tap a room to open or close it. Closing a staffed room asks where they go. Tap a name to move them.";
    }
  }

  function setHeldRoom(cat, room, intent, closeFrom) {
    g.selectedDeck = null;
    g.heldMove = { cat: cat, room: room, intent: intent || "place", closeFrom: !!closeFrom };
    var rec = occupantOf(cat, room);
    var nm = rec ? chipName(rec.name) : "them";
    g.moveHint = intent === "swap"
      ? "Tap the room " + nm + " swaps with · tap " + room + " to cancel"
      : "Tap the room " + nm + " moves to · tap " + room + " to cancel";
    try { document.body.classList.add("assigning"); } catch (e) {}
    if (g.sitesFn === "people") {
      g.sitesFn = "rooms";
      try { if (typeof enhanceEditDayModal === "function") enhanceEditDayModal(true); } catch (e) {}
    }
    try { if (typeof showToast === "function") showToast(g.moveHint); } catch (e) {}
    renderOnDeck();
    refreshMoveHints();
  }

  function sendRecToDeck(rec, lastRoom, select) {
    if (!rec || !rec.name) return;
    g.onDeck = g.onDeck || [];
    var k = nameKey(rec.name);
    g.onDeck = g.onDeck.filter(function (p) { return nameKey(p.name) !== k; });
    g.onDeck.unshift({
      name: rec.name, shift: rec.shift || "", kind: rec.kind || "none",
      lastRoom: lastRoom || "", role: "freed", student: !!rec.student,
      firstCase: rec.firstCase || "", freedAt: Date.now()
    });
    if (select) {
      g.selectedDeck = k;
      g.heldMove = null;
      g.moveHint = "Tap a room to place " + chipName(rec.name);
      try { document.body.classList.add("assigning"); } catch (e) {}
      if (g.sitesFn === "people") {
        g.sitesFn = "rooms";
        try { if (typeof enhanceEditDayModal === "function") enhanceEditDayModal(true); } catch (e) {}
      }
    }
  }

  function afterStaffMove() {
    persistStaff();
    try { if (typeof updateActivePOCCounter === "function") updateActivePOCCounter(); } catch (e) {}
    refreshMoveHints();
  }

  function finishPlant(held, toCat, toRoom, how) {
    if (!held || !toCat || !toRoom) return false;
    g.roomStaff = g.roomStaff || {};
    g.roomStaff[toCat] = g.roomStaff[toCat] || {};
    var occ = occupantOf(toCat, toRoom);
    var mover = null;
    var moverName = "";
    if (held.kind === "deck") {
      var list = g.onDeck || [];
      var idx = -1;
      for (var i = 0; i < list.length; i++) {
        if (nameKey(list[i].name) === held.key) { idx = i; break; }
      }
      if (idx < 0) { clearHeld(); return false; }
      mover = list[idx];
      moverName = mover.name;
      list.splice(idx, 1);
      if (how === "swap" && occ) sendRecToDeck(occ, toRoom, true);
      else if (how === "hand" && occ) sendRecToDeck(occ, toRoom, true);
      else if (how === "bump" && occ) sendRecToDeck(occ, toRoom, false);
      g.roomStaff[toCat][toRoom] = cloneStaffRec(mover);
    } else {
      mover = occupantOf(held.cat, held.room);
      if (!mover) { clearHeld(); return false; }
      moverName = mover.name;
      if (how === "swap" && occ) {
        var a = cloneStaffRec(mover);
        var b = cloneStaffRec(occ);
        g.roomStaff[held.cat] = g.roomStaff[held.cat] || {};
        g.roomStaff[held.cat][held.room] = b;
        g.roomStaff[toCat][toRoom] = a;
        try {
          if (typeof catEditState !== "undefined") {
            if (catEditState[held.cat]) catEditState[held.cat].deletedRooms.delete(held.room);
            if (catEditState[toCat]) catEditState[toCat].deletedRooms.delete(toRoom);
          }
        } catch (e) {}
        var keep = how === "swap" && held.kind === "deck";
        if (!keep) clearHeld();
        afterStaffMove();
        try { if (typeof showToast === "function") showToast(chipName(a.name) + " ↔ " + chipName(b.name)); } catch (e2) {}
        return true;
      }
      if (how === "hand" && occ) sendRecToDeck(occ, toRoom, true);
      else if (how === "bump" && occ) sendRecToDeck(occ, toRoom, false);
      g.roomStaff[toCat][toRoom] = cloneStaffRec(mover);
      g.roomStaff[held.cat] = g.roomStaff[held.cat] || {};
      delete g.roomStaff[held.cat][held.room];
      closeRoomIfEmpty(held.cat, held.room);
    }
    try {
      if (typeof catEditState !== "undefined" && catEditState[toCat]) {
        catEditState[toCat].deletedRooms.delete(toRoom);
        if (catEditState[toCat].deletedEvents) catEditState[toCat].deletedEvents[toRoom] = { deleted: false, ts: Date.now() };
      }
    } catch (e) {}
    var keepHand = (how === "hand" || (how === "swap" && held.kind === "deck")) && occ;
    if (!keepHand) clearHeld();
    else {
      g.heldMove = null;
    }
    afterStaffMove();
    try {
      if (typeof showToast === "function") {
        if (keepHand) showToast(chipName(moverName) + " → " + toRoom + " · place " + chipName(occ.name) + " next");
        else showToast(chipName(moverName) + " → " + toRoom);
      }
    } catch (e4) {}
    return true;
  }

  function closePlantAsk() {
    var ov = document.getElementById("plant-ask-overlay");
    if (ov) ov.remove();
  }

  function askOccupiedPlant(held, toCat, toRoom) {
    var occ = occupantOf(toCat, toRoom);
    var mover = held.kind === "deck"
      ? (g.onDeck || []).filter(function (p) { return nameKey(p.name) === held.key; })[0]
      : occupantOf(held.cat, held.room);
    if (!occ || !mover) return false;
    closePlantAsk();
    var ov = document.createElement("div");
    ov.id = "plant-ask-overlay";
    ov.className = "deact-sheet-overlay";
    var swapTitle = held.kind === "deck"
      ? "Swap — " + chipName(mover.name) + " in " + toRoom + ", keep " + chipName(occ.name) + " in your hand"
      : "Swap — they trade rooms";
    var swapSub = held.kind === "deck"
      ? chipName(occ.name) + " stays selected so you can tap the next room"
      : chipName(mover.name) + " ↔ " + chipName(occ.name);
    ov.innerHTML =
      '<div class="deact-sheet">' +
      '<div class="deact-kicker">' + toRoom + " is staffed</div>" +
      '<div class="deact-who">' + shiftPillHtml(mover.shift) + '<span class="staff-name">' + chipName(mover.name) + "</span>" +
      '<span class="staff-last">→ ' + toRoom + "</span></div>" +
      '<div class="deact-ask">Now: ' + shiftPillHtml(occ.shift) + " " + chipName(occ.name) + "</div>" +
      '<button type="button" class="deact-choice" data-act="swap"><span class="deact-choice-title">' + swapTitle + '</span><span class="deact-choice-sub">' + swapSub + '</span></button>' +
      '<button type="button" class="deact-choice" data-act="hand"><span class="deact-choice-title">Move ' + chipName(occ.name) + ' to another room</span><span class="deact-choice-sub">' + chipName(mover.name) + ' takes ' + toRoom + ' · then tap where ' + chipName(occ.name) + ' goes</span></button>' +
      '<button type="button" class="deact-choice" data-act="bump"><span class="deact-choice-title">Bump ' + chipName(occ.name) + ' to deck</span><span class="deact-choice-sub">' + chipName(mover.name) + ' in ' + toRoom + ' · ' + chipName(occ.name) + ' free</span></button>' +
      '<button type="button" class="deact-cancel" data-act="cancel">Cancel</button></div>';
    ov.addEventListener("click", function (e) { if (e.target === ov) closePlantAsk(); });
    ov.querySelector('[data-act="cancel"]').onclick = closePlantAsk;
    ov.querySelector('[data-act="swap"]').onclick = function () {
      closePlantAsk();
      finishPlant(held, toCat, toRoom, "swap");
    };
    ov.querySelector('[data-act="hand"]').onclick = function () {
      closePlantAsk();
      finishPlant(held, toCat, toRoom, "hand");
    };
    ov.querySelector('[data-act="bump"]').onclick = function () {
      closePlantAsk();
      finishPlant(held, toCat, toRoom, "bump");
    };
    document.body.appendChild(ov);
    return true;
  }

  function plantHeldInto(toCat, toRoom) {
    var held = heldFrom();
    if (!held) return false;
    if (held.kind === "room" && held.cat === toCat && held.room === toRoom) {
      clearHeld();
      try { if (typeof showToast === "function") showToast("Cancelled"); } catch (e) {}
      renderOnDeck();
      return true;
    }
    var occ = occupantOf(toCat, toRoom);
    if (held.intent === "swap") {
      if (!occ) return finishPlant(held, toCat, toRoom, "empty");
      return finishPlant(held, toCat, toRoom, "swap");
    }
    if (occ) return askOccupiedPlant(held, toCat, toRoom);
    return finishPlant(held, toCat, toRoom, "empty");
  }

  function tryHandleRoomTap(catId, room, ev) {
    if (!canMoveStaff()) return false;
    if (heldFrom()) return plantHeldInto(catId, room);
    var t = ev && ev.target;
    if (t && t.closest && t.closest(".staff-move")) {
      if (ev.preventDefault) ev.preventDefault();
      if (ev.stopPropagation) ev.stopPropagation();
      openPersonSheet("room", "", catId, room);
      return true;
    }
    return false;
  }

  function assignSelectedTo(catId, room) {
    if (!canMoveStaff()) {
      clearHeld();
      return false;
    }
    if (!g.selectedDeck && !g.heldMove) return false;
    return plantHeldInto(catId, room);
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
    var lsMap = lateStayByName();
    var byName = {};
    rooms.forEach(function (r) {
      var k = nameKey(r.out && r.out.name);
      if (!k) return;
      if (!byName[k]) {
        byName[k] = r;
        byName[k].extraRooms = [];
      } else {
        byName[k].extraRooms.push({ cat: r.cat, room: r.room });
      }
    });
    var merged = Object.keys(byName).map(function (k) {
      var r = byName[k];
      var extras = r.extraRooms || [];
      r.label = extras.length
        ? [r.room].concat(extras.map(function (x) { return x.room; })).join(" + ")
        : r.room;
      var ls = lsMap[k];
      if (ls) {
        r.latestay = true;
        r.lsN = ls.n;
      }
      return r;
    });
    return sortReliefRooms(merged);
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
    if (purpose === "relief" && !planned) {
      var idle = longestIdleDeck(list.deck);
      if (idle) {
        planned = { name: idle.name, from: "deck", fromCat: "", fromRoom: "", shift: idle.shift };
        plannedKey = nameKey(idle.name);
      }
    }

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
        waveClock(list.wave);
    }

    var suggestHtml = "";
    if (purpose === "relief" && planned && planned.name) {
      suggestHtml =
        '<button type="button" class="sites-choice relief-suggest" id="relief-suggest-btn">' +
        "<strong>Send " + chipName(planned.name) + "</strong>" +
        "<span>" + (planned.from === "room" ? (planned.fromRoom || "in a room") : "On deck · longest idle") +
        " · tap to place</span></button>";
    }

    var ov = document.createElement("div");
    ov.id = "relief-roster-overlay";
    ov.className = "relief-roster-overlay";
    ov.innerHTML =
      '<div class="relief-roster-card">' +
      '<div class="relief-roster-title">' + title + "</div>" +
      '<div class="relief-roster-sub">' + sub + "</div>" +
      suggestHtml +
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
    var sug = ov.querySelector("#relief-suggest-btn");
    if (sug && planned) {
      sug.onclick = function (ev) {
        ev.stopPropagation();
        if (planned.from === "room") pickRelief("room", planned.fromCat, planned.fromRoom, plannedKey, cat, room);
        else pickRelief("deck", "", "", plannedKey, cat, room);
      };
    }
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
    var held = from === "deck"
      ? { kind: "deck", key: fromKey, intent: "place" }
      : { kind: "room", cat: fromCat, room: fromRoom, intent: "place" };
    closeReliefRoster();
    var occ = occupantOf(toCat, toRoom);
    if (occ) { askOccupiedPlant(held, toCat, toRoom); return; }
    finishPlant(held, toCat, toRoom, "empty");
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
    var held = { kind: "room", cat: src.cat, room: src.room, closeFrom: true, intent: "place" };
    g.pendingPour = null;
    setEditSitesHint("");
    var occ = occupantOf(toCat, toRoom);
    if (occ) {
      askOccupiedPlant(held, toCat, toRoom);
      paintEditDayButtons();
      return;
    }
    finishPlant(held, toCat, toRoom, "empty");
    paintEditDayButtons();
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
        '<span class="relief-room">' + (j.label || j.room) + "</span>" +
        shiftPillHtml(j.out.shift) +
        '<span class="staff-name">' + outNm + "</span>" +
        (j.latestay ? '<span class="staff-last">LS #' + j.lsN + "</span>" : "") +
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
      fn === "people" ? "Search a name, then send them to deck, move, or swap. Free now is longest-idle first."
      : fn === "relief" ? "Next wave is open. Later hours and doctors are folded. Late stays show once with LS #."
      : "Tap a room to open or close it. Closing a staffed room asks where they go. Tap a name to move them.";
    var viewBtn = (g.lastSheet && g.lastSheet.cells) || (g.assignmentMeta && g.assignmentMeta.cells)
      ? '<button type="button" class="desk-upload-btn" id="desk-view-sheet">View uploaded sheet</button>'
      : "";
    var undoBtn = g.undoClose && g.undoClose.rec
      ? '<button type="button" class="desk-undo-btn" id="desk-undo-btn">Undo · ' + chipName(g.undoClose.rec.name) + " → " + g.undoClose.room + "</button>"
      : '<button type="button" class="desk-undo-btn" id="desk-undo-btn" style="display:none"></button>';
    return '<div class="sites-fn-tabs">' +
      tab("rooms", "Rooms") +
      tab("people", "People") +
      tab("relief", "Relief") +
      "</div>" +
      '<div class="sites-fn-hint">' + hint + "</div>" +
      undoBtn +
      '<button type="button" class="desk-upload-btn" id="desk-upload-btn">Upload assignment sheet</button>' +
      viewBtn +
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
        '<span class="relief-room">' + (j.label || j.room) + "</span>" +
        shiftPillHtml(j.out.shift) +
        '<span class="staff-name">' + chipName(j.out.name) + "</span>" +
        (j.latestay ? '<span class="staff-last">LS #' + j.lsN + "</span>" : "") +
        '<span class="relief-arrow">→</span>' +
        (j.inn ? shiftPillHtml(j.inn.shift) : "") +
        '<span class="staff-name' + (j.inn ? "" : " missing") + '">' + inNm + "</span>" +
        place + "</button>";
    }
    var hour = hospitalHour();
    g.reliefOpen = g.reliefOpen || {};
    var blocks = waves.map(function (block, i) {
      var explicit = g.reliefOpen[block.wave];
      var open = explicit === undefined ? (i === 0) : !!explicit;
      var chips = open ? block.rooms.map(chipHtml).join("") : "";
      var title = "Out at " + waveClock(block.wave) + " · " + block.rooms.length;
      return '<button type="button" class="relief-wave-toggle" data-fold="wave" data-wave="' + block.wave + '" aria-expanded="' + (open ? "true" : "false") + '">' +
        (open ? "Hide · " : "") + title + (open ? " ▴" : " ▾") + "</button>" +
        (open ? '<div class="relief-need-grid">' + chips + "</div>" : "");
    }).join("");
    if (docs.length) {
      var docsOpen = g.docsOpen === undefined ? (hour >= 16) : !!g.docsOpen;
      blocks += '<button type="button" class="relief-wave-toggle" data-fold="wave" data-wave="md" aria-expanded="' + (docsOpen ? "true" : "false") + '">' +
        (docsOpen ? "Hide · " : "") + "Doctors · usually 5:30 · " + docs.length + (docsOpen ? " ▴" : " ▾") + "</button>";
      if (docsOpen) blocks += '<div class="relief-need-grid">' + docs.map(chipHtml).join("") + "</div>";
    }
    pane.innerHTML =
      (blocks || '<div class="roster-empty">Nobody leaving yet</div>') +
      queueBlockHtml("Late (M+S)", late) +
      queueBlockHtml("Dinner (Q+W+E)", dinner);
    pane.querySelectorAll(".relief-wave-toggle").forEach(function (el) {
      el.onclick = function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        var w = el.getAttribute("data-wave");
        if (w === "md") g.docsOpen = !(g.docsOpen === undefined ? (hospitalHour() >= 16) : g.docsOpen);
        else {
          g.reliefOpen = g.reliefOpen || {};
          var n = parseInt(w, 10);
          var was = el.getAttribute("aria-expanded") === "true";
          g.reliefOpen[n] = !was;
        }
        renderDeskRelief();
      };
    });
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
      rooms.style.display = "none";
    }
  }

  function renderDeskPeople() {
    var pane = document.getElementById("desk-people-pane");
    if (!pane) return;
    var list = (window.onDeck || g.onDeck || []).filter(function (p) {
      return p && p.name && !isJunkStaff(p.name, p.shift);
    });
    var parts = splitDeck(list);
    var here = parts.now, later = parts.later, gone = parts.gone;
    var q = nameKey(g.deskPeopleQ || "");
    function pill(p, mark) {
      var tag = deckTag(p);
      return '<button type="button" class="roster-pill' + (mark ? " " + mark : "") + '" data-deckkey="' + nameKey(p.name) + '">' +
        shiftPillHtml(p.shift) + '<span class="staff-name">' + chipName(p.name) + "</span>" +
        studentMark(hasStudent(p)) + (tag ? '<span class="staff-last">' + tag + "</span>" : "") +
        (mark === "out" ? '<span class="staff-last">out</span>' : (mark === "later" ? '<span class="staff-last">later</span>' : "")) +
        "</button>";
    }
    var hereChips = here.map(function (p) { return pill(p, ""); }).join("");
    var laterChips = later.map(function (p) { return pill(p, "later"); }).join("");
    var goneChips = gone.map(function (p) { return pill(p, "out"); }).join("");
    var laterOpen = !!g.deskLaterOpen;
    var laterBlock = later.length
      ? '<button type="button" class="ondeck-out-toggle" data-fold="later" id="desk-later-toggle" aria-expanded="' + (laterOpen ? "true" : "false") + '">' +
        (laterOpen ? "Hide later · " : "Later · ") + later.length + (laterOpen ? " ▴" : " ▾") + "</button>" +
        (laterOpen ? '<div class="roster-grid">' + laterChips + "</div>" : "")
      : "";
    var open = !!g.deskOutOpen;
    var outBlock = gone.length
      ? '<button type="button" class="ondeck-out-toggle" data-fold="out" id="desk-out-toggle" aria-expanded="' + (open ? "true" : "false") + '">' +
        (open ? "Hide out · " : "Out · ") + gone.length + (open ? " ▴" : " ▾") + "</button>" +
        (open ? '<div class="roster-grid">' + goneChips + "</div>" : "")
      : "";
    var inRoom = [];
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = occupantOf(cat, room);
        if (!rec) return;
        inRoom.push({ cat: cat, room: room, rec: rec });
      });
    });
    inRoom.sort(function (a, b) {
      var da = catOrder(a.cat), db = catOrder(b.cat);
      if (da !== db) return da - db;
      return String(chipName(a.rec.name)).localeCompare(String(chipName(b.rec.name)));
    });
    var roomChips = inRoom.map(function (x) {
      return '<button type="button" class="roster-pill" data-roomkey="' + x.cat + "|" + x.room + '" data-deckkey="' + nameKey(x.rec.name) + '">' +
        shiftPillHtml(x.rec.shift) + '<span class="staff-name">' + chipName(x.rec.name) + "</span>" +
        studentMark(hasStudent(x.rec)) + '<span class="staff-last">' + x.room + "</span></button>";
    }).join("");
    var callList = (g.calls || (g.assignmentMeta && g.assignmentMeta.calls) || []).filter(function (p) {
      return p && p.name;
    });
    var callChips = callList.map(function (p) {
      var hits = boardHits(p.name);
      var roomHit = hits.filter(function (h) { return h.where === "room"; })[0];
      var cls = ' class="roster-pill" data-deckkey="' + nameKey(p.name) + '"';
      if (roomHit) cls += ' data-roomkey="' + roomHit.cat + "|" + roomHit.room + '"';
      return '<button type="button"' + cls + ">" +
        shiftPillHtml(p.shift) + '<span class="staff-name">' + chipName(p.name) + "</span>" +
        '<span class="staff-last">' + (p.lastRoom || "call") + (roomHit ? " · " + roomHit.room : "") + "</span></button>";
    }).join("");
    var callBlock = callChips
      ? '<div class="assign-cat-label">Call</div><div class="roster-grid">' + callChips + "</div>"
      : "";
    pane.innerHTML =
      '<input id="desk-people-q" class="desk-people-search" type="search" placeholder="Search a name" autocomplete="off" autocorrect="off" spellcheck="false">' +
      '<button type="button" class="desk-add-btn" id="desk-add-staff">Add someone who came in</button>' +
      '<div class="assign-cat-label">Free now · longest idle first</div>' +
      '<div class="roster-grid">' + (hereChips || '<span class="roster-empty">Nobody free right now</span>') + "</div>" +
      laterBlock +
      outBlock +
      callBlock +
      '<div class="assign-cat-label">In a room</div>' +
      '<div class="roster-grid">' + (roomChips || '<span class="roster-empty">No one in a room</span>') + "</div>";
    var add = document.getElementById("desk-add-staff");
    if (add) add.onclick = function () { openAddStaff("", ""); };
    var tog = document.getElementById("desk-out-toggle");
    if (tog) tog.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      g.deskOutOpen = !g.deskOutOpen;
      renderDeskPeople();
    };
    var laterTog = document.getElementById("desk-later-toggle");
    if (laterTog) laterTog.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      g.deskLaterOpen = !g.deskLaterOpen;
      renderDeskPeople();
    };
    pane.querySelectorAll("[data-deckkey]").forEach(function (el) {
      el.onclick = function (ev) {
        ev.stopPropagation();
        var rk = el.getAttribute("data-roomkey");
        if (rk) {
          var bits = rk.split("|");
          openPersonSheet("room", "", bits[0], bits.slice(1).join("|"));
          return;
        }
        openPersonSheet("deck", el.getAttribute("data-deckkey"), "", "");
      };
    });
    var inp = document.getElementById("desk-people-q");
    if (inp) {
      inp.value = g.deskPeopleQ || "";
      inp.oninput = function () {
        g.deskPeopleQ = inp.value;
        var qk = nameKey(inp.value);
        var needLater = qk && !g.deskLaterOpen && later.some(function (p) { return nameKey(p.name).indexOf(qk) !== -1; });
        var needOut = qk && !g.deskOutOpen && gone.some(function (p) { return nameKey(p.name).indexOf(qk) !== -1; });
        if (needLater || needOut) {
          if (needLater) g.deskLaterOpen = true;
          if (needOut) g.deskOutOpen = true;
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
    hydrateStaff();
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
    var view = document.getElementById("desk-view-sheet");
    if (view) view.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      showSheetViewer();
    };
    var undoBtn = document.getElementById("desk-undo-btn");
    if (undoBtn) undoBtn.onclick = function (ev) {
      ev.preventDefault();
      ev.stopPropagation();
      undoLastClose();
      enhanceEditDayModal(true);
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
    if (rooms) rooms.style.display = (onRelief || onPeople) ? "none" : "";
    if (pane) {
      pane.style.display = onRelief ? "block" : "none";
      if (onRelief) renderDeskRelief();
    }
    if (people) {
      people.style.display = onPeople ? "block" : "none";
      if (onPeople) renderDeskPeople();
    }
    var go = document.querySelector("#weekend-overlay .weekend-go-btn");
    if (go) go.style.display = "none";
    var back = document.getElementById("weekend-back-btn");
    if (back) {
      back.style.display = "block";
      back.textContent = "Close";
      back.onclick = function (ev) {
        if (ev) { ev.preventDefault(); ev.stopPropagation(); }
        try {
          if (typeof confirmEditActiveSites === "function") confirmEditActiveSites();
          else {
            document.getElementById("weekend-overlay").classList.remove("show");
            if (typeof _resetModalToDefaults === "function") _resetModalToDefaults();
          }
        } catch (e) {}
      };
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
      '<div class="relief-roster-sub">Removes names, on-deck, relief picks, and which rooms were open or closed. You\'ll upload the sheet again.</div>' +
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

  function resetRoomsAfterClear() {
    var now = Date.now();
    var weekend = false;
    try { weekend = typeof isWeekend === "function" && isWeekend(); } catch (e) {}
    if (typeof CATEGORIES === "undefined") return;
    CATEGORIES.forEach(function (c) {
      var rooms = c.rooms || [];
      if (typeof catEditState !== "undefined" && catEditState[c.id]) {
        var es = catEditState[c.id];
        es.deletedRooms = new Set(weekend ? rooms.slice() : []);
        es.deletedEvents = {};
        rooms.forEach(function (r) {
          es.deletedEvents[r] = { deleted: !!weekend, ts: now };
        });
      }
      if (typeof weekendActive !== "undefined") {
        weekendActive[c.id] = weekendActive[c.id] || {};
        rooms.forEach(function (r) {
          weekendActive[c.id][r] = !weekend;
          var btn = document.getElementById("wknd-" + c.id + "-" + r.replace(/\s/g, "_"));
          if (btn) {
            btn.classList.toggle("active", !weekend);
            btn.classList.toggle("room-inactive-marker", !!weekend);
            var line = btn.querySelector(".staff-line");
            if (line) line.remove();
          }
        });
      }
    });
    try { if (typeof updateActivePOCCounter === "function") updateActivePOCCounter(); } catch (e) {}
    try { if (typeof updateSummary === "function") updateSummary(); } catch (e) {}
  }

  function clearAssignmentBoard() {
    g.roomStaff = {};
    g.onDeck = [];
    g.reliefPlan = {};
    g.lateStays = [];
    g.calls = [];
    g.lastSheet = null;
    g.undoClose = null;
    try { localStorage.removeItem("anespresso_runner_sheet_v1"); } catch (e) {}
    if (typeof CATEGORIES !== "undefined") {
      CATEGORIES.forEach(function (c) { g.roomStaff[c.id] = {}; });
    }
    g.assignmentMeta = {
      cleared: true,
      appliedAt: Date.now(),
      date: (typeof todayStr === "function") ? todayStr() : "",
      file: "",
      roomsN: 0, closedN: 0, deckN: 0, lateStays: [], calls: []
    };
    resetRoomsAfterClear();
    publishStaff();
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try { if (typeof refreshBoard === "function") refreshBoard(); } catch (e) {}
    try { renderOnDeck(); } catch (e) {}
    try { paintWkndStaff(); } catch (e) {}
    try { renderDeskPeople(); } catch (e) {}
    try { renderDeskRelief(); } catch (e) {}
    try { renderShiftChange(); } catch (e) {}
    try { renderLateBoardBar(); } catch (e) {}
    try { if (typeof showToast === "function") showToast("Board cleared"); } catch (e) {}
    try {
      if (document.getElementById("weekend-overlay") && document.getElementById("weekend-overlay").classList.contains("show")) {
        paintWkndStaff();
      }
    } catch (e) {}
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
        var held = { kind: "room", cat: fromCat, room: fromRoom, closeFrom: true, intent: "place" };
        var occ = occupantOf(toCat, toRoom);
        closeSitesSheet();
        if (occ) { askOccupiedPlant(held, toCat, toRoom); return; }
        finishPlant(held, toCat, toRoom, "empty");
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
          var held = from === "deck"
            ? { kind: "deck", key: el.getAttribute("data-rkey"), intent: "place" }
            : { kind: "room", cat: el.getAttribute("data-rcat"), room: el.getAttribute("data-rroom"), intent: "place" };
          closeSitesSheet();
          var occ = occupantOf(cat, room);
          if (occ) { askOccupiedPlant(held, cat, room); return; }
          finishPlant(held, cat, room, "empty");
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
    publishStaff();
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
      if (t.closest("#assign-upload-overlay, .assign-upload-overlay, .assign-apply, .assign-cancel, #remove-staff-overlay, #sites-sheet-overlay, #force-update-overlay, #plant-ask-overlay, .sheet-handle, [data-fold], .ondeck-out-toggle, .relief-wave-toggle")) return;
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
    var moves = "";
    if (canMoveStaff()) {
      if (kind === "deck") {
        moves += '<button type="button" class="sites-choice" data-se="place"><strong>Place in a room</strong><span>Tap a room next — occupied rooms will ask swap or bump</span></button>';
      } else {
        moves += '<button type="button" class="sites-choice" data-se="deck"><strong>Send to deck</strong><span>Free — ready for breaks or relief</span></button>';
        moves += '<button type="button" class="sites-choice" data-se="move"><strong>Move to a room</strong><span>They stay assigned — tap the next OR</span></button>';
        moves += '<button type="button" class="sites-choice" data-se="swap"><strong>Swap with another room</strong><span>Tap the room they trade with</span></button>';
      }
    }
    ov.innerHTML =
      '<div class="sites-sheet-card">' +
      '<div class="relief-roster-title">' + chipName(rec.name) + "</div>" +
      '<div class="relief-roster-sub">' + shiftPillHtml(rec.shift) +
      (kind === "deck" ? '<span class="staff-last">on deck</span>' : '<span class="roster-loc">' + room + "</span>") +
      "</div>" +
      moves +
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
    var place = ov.querySelector('[data-se="place"]');
    if (place) place.onclick = function () {
      closeSitesSheet();
      g.heldMove = null;
      g.selectedDeck = key;
      g.moveHint = "Tap a room to place " + chipName(rec.name);
      try { document.body.classList.add("assigning"); } catch (e) {}
      if (g.sitesFn === "people") {
        g.sitesFn = "rooms";
        try { if (typeof enhanceEditDayModal === "function") enhanceEditDayModal(true); } catch (e) {}
      }
      renderOnDeck();
      refreshMoveHints();
      try { if (typeof showToast === "function") showToast(g.moveHint); } catch (e) {}
    };
    var toDeck = ov.querySelector('[data-se="deck"]');
    if (toDeck) toDeck.onclick = function () {
      closeSitesSheet();
      sendOccupantToDeck(cat, room);
      persistStaff();
      try { if (typeof showToast === "function") showToast(chipName(rec.name) + " → on deck · " + room + " closed"); } catch (e) {}
    };
    var move = ov.querySelector('[data-se="move"]');
    if (move) move.onclick = function () {
      closeSitesSheet();
      setHeldRoom(cat, room, "place", true);
    };
    var swap = ov.querySelector('[data-se="swap"]');
    if (swap) swap.onclick = function () {
      closeSitesSheet();
      setHeldRoom(cat, room, "swap", false);
    };
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
      window.weekendToggleRoom = function (catId, room, btn, ev) {
        ev = ev || (typeof window !== "undefined" ? window.event : null);
        if (tryHandleRoomTap(catId, room, ev)) return;
        if (isEditDayModal() || (document.getElementById("weekend-day-badge") || {}).textContent === "Active Sites") {
          if (g.sitesFn === "staffing" || g.sitesFn === "people") {
            openStaffingSheet(catId, room);
            return;
          }
          if (g.sitesFn === "relief") return;
          var occ = occupantOf(catId, room);
          if (occ) {
            openPourSheet(catId, room);
            return;
          }
          var on = true;
          try {
            if (typeof catEditState !== "undefined" && catEditState[catId] && catEditState[catId].deletedRooms) {
              on = !catEditState[catId].deletedRooms.has(room);
            }
          } catch (e) {}
          try { applyRoomActive(catId, room, !on); } catch (e2) {
            try { if (on) deactivateRoomNow(catId, room); } catch (e3) {}
          }
          persistStaff();
          return;
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
          hydrateStaff();
          if (!g.lastSheet) {
            try {
              var raw = localStorage.getItem("anespresso_runner_sheet_v1");
              if (raw) g.lastSheet = JSON.parse(raw);
            } catch (e2) {}
          }
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
        hydrateStaff();
        origOpen.apply(this, arguments);
        scrollDeskToTop();
        setTimeout(function () { try { enhanceEditDayModal(); scrollDeskToTop(); } catch (e) {} }, 0);
      };
    }
    if (typeof buildPayload === "function" && !g._buildWrapped) {
      g._buildWrapped = true;
      var origBuildPay = buildPayload;
      window.buildPayload = function () {
        publishStaff();
        var p = origBuildPay.apply(this, arguments);
        if (p && p.assignmentMeta) {
          var m = {};
          Object.keys(p.assignmentMeta).forEach(function (k) {
            if (k !== "cells" && k !== "preview" && k !== "unsure") m[k] = p.assignmentMeta[k];
          });
          p.assignmentMeta = m;
        }
        return p;
      };
    }
  }

  function wrapDeactivate() {
    if (typeof toggleDeleteRoom === "function" && !g._deckWrapped) {
      g._deckWrapped = true;
      var orig = toggleDeleteRoom;
      window.toggleDeleteRoom = function (catId, room) {
        if (heldFrom() && assignSelectedTo(catId, room)) return;
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
        if (tryHandleRoomTap(catId, room, ev)) return;
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

  function refreshAfterAssign() {
    try { if (typeof saveShared === "function") saveShared(); } catch (e) {}
    try {
      if (typeof buildBoard === "function" && !document.getElementById("card-nt")) buildBoard();
      else if (typeof refreshBoard === "function") refreshBoard();
    } catch (e) {}
    try { if (typeof updateActivePOCCounter === "function") updateActivePOCCounter(); } catch (e) {}
    try { renderLateBoardBar(); } catch (e) {}
    try { renderOnDeck(); } catch (e) {}
    try { renderShiftChange(); } catch (e) {}
    try { paintWkndStaff(); } catch (e) {}
    try { if (typeof renderDeskPeople === "function") renderDeskPeople(); } catch (e) {}
    try { if (typeof enhanceEditDayModal === "function") enhanceEditDayModal(true); } catch (e) {}
  }

  function boardHits(name) {
    var hits = [];
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = g.roomStaff[cat][room];
        if (rec && rec.name && !rec.closed && namesMatch(rec.name, name)) {
          hits.push({ cat: cat, room: room, rec: rec, where: "room" });
        }
      });
    });
    (g.onDeck || []).forEach(function (p) {
      if (p && p.name && namesMatch(p.name, name)) {
        hits.push({ where: "deck", rec: p, lastRoom: p.lastRoom || "" });
      }
    });
    return hits;
  }

  function liveOccupant(cat, room) {
    var rec = g.roomStaff && g.roomStaff[cat] && g.roomStaff[cat][room];
    if (!rec || rec.closed || !rec.name) return null;
    return rec;
  }

  function auditBoardVsSheet(res) {
    var issues = [];
    var seen = {};
    function add(it) {
      var k = (it.reason || "") + "|" + nameKey(it.name || "") + "|" + String(it.roomRaw || "");
      if (seen[k]) return;
      seen[k] = 1;
      issues.push(it);
    }
    function describeHits(hits) {
      if (!hits || !hits.length) return "not on the board";
      return hits.map(function (h) {
        if (h.where === "deck") return "on deck" + (h.lastRoom ? " (" + h.lastRoom + ")" : "");
        return h.room;
      }).join(", ");
    }
    (res.sheetPeople || []).forEach(function (p) {
      if (!p || !p.name) return;
      var hits = boardHits(p.name);
      var anywhere = p.role === "wbf" || p.role === "breaker" || p.role === "midnight" || p.role === "call" || p.role === "latestay" || p.role === "resident" || p.role === "evening" || p.role === "night" || p.role === "srna" || p.role === "runner" || p.role === "extra" || p.role === "offsite" || p.role === "float";
      if (p.role === "shift" || parseShiftLabel(p.roomRaw)) {
        if (hits.length) return;
        var guessed = (parseShiftLabel(p.roomRaw) || {}).shift || p.shift || "";
        add({
          reason: "shift",
          roomRaw: p.roomRaw,
          staffRaw: p.staffRaw,
          name: p.name,
          shift: p.shift || guessed,
          student: p.student,
          kind: p.kind,
          suggestions: [],
          guessedShift: guessed,
          cellRefs: p.cellRefs || [],
          boardNow: describeHits(hits)
        });
        add({
          reason: "need-room",
          roomRaw: p.roomRaw,
          staffRaw: p.staffRaw,
          name: p.name,
          shift: p.shift || guessed,
          student: p.student,
          kind: p.kind,
          suggestions: [],
          cellRefs: p.cellRefs || [],
          boardNow: describeHits(hits)
        });
        return;
      }
      if (anywhere) {
        if (p.role === "latestay") {
          var onLs = (g.lateStays || []).some(function (x) { return x && namesMatch(x.name, p.name); });
          if (onLs) return;
        }
        if (!hits.length) {
          add({
            reason: "missing", roomRaw: p.roomRaw, staffRaw: p.staffRaw, name: p.name,
            shift: p.shift, student: p.student, kind: p.kind, suggestions: suggestFor(p.roomRaw),
            boardNow: "not on the board"
          });
        }
        return;
      }
      if (p.locs && p.locs.length) {
        p.locs.forEach(function (loc) {
          var rec = liveOccupant(loc.cat, loc.room);
          if (rec && namesMatch(rec.name, p.name)) return;
          add({
            reason: "mismatch",
            roomRaw: loc.room,
            staffRaw: p.staffRaw,
            name: p.name,
            shift: p.shift,
            student: p.student,
            kind: p.kind,
            suggestions: [{ cat: loc.cat, room: loc.room }].concat(suggestFor(p.roomRaw)),
            cellRefs: p.cellRefs || [],
            boardNow: rec && rec.name ? rec.name + " in " + loc.room : loc.room + " empty"
          });
        });
        if (!p.shift && !p.student && coreShift(p.shift) !== "Dr" && !/SRNA|RESIDENT/i.test(p.roomRaw || "")) {
          var placedOk = p.locs.some(function (loc) {
            var rec = liveOccupant(loc.cat, loc.room);
            return rec && namesMatch(rec.name, p.name);
          });
          if (placedOk) {
            add({
              reason: "need-shift",
              roomRaw: p.roomRaw || (p.locs[0] && p.locs[0].room) || "",
              staffRaw: p.staffRaw,
              name: p.name,
              shift: "",
              student: p.student,
              kind: p.kind,
              suggestions: [],
              boardNow: describeHits(hits)
            });
          }
        }
        return;
      }
      if (!hits.length) {
        add({
          reason: "missing", roomRaw: p.roomRaw, staffRaw: p.staffRaw, name: p.name,
          shift: p.shift, student: p.student, kind: p.kind, suggestions: suggestFor(p.roomRaw),
          boardNow: "not on the board"
        });
      } else if (p.roomRaw && !skipRoomLabel(p.roomRaw) && !/^(BR|night|call|WBF|on deck)$/i.test(p.roomRaw)) {
        add({
          reason: "unknown-room", roomRaw: p.roomRaw, staffRaw: p.staffRaw, name: p.name,
          shift: p.shift, student: p.student, kind: p.kind, suggestions: suggestFor(p.roomRaw),
          cellRefs: p.cellRefs || [],
          boardNow: describeHits(hits)
        });
      }
    });
    var onSheet = res.sheetPeople || [];
    function sheetHas(name) {
      return onSheet.some(function (p) { return p && namesMatch(p.name, name); });
    }
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = liveOccupant(cat, room);
        if (!rec) return;
        if (sheetHas(rec.name)) return;
        add({
          reason: "extra", roomRaw: room, staffRaw: rec.name, name: rec.name,
          shift: rec.shift, student: !!rec.student, kind: rec.kind || "none",
          cat: cat, room: room, suggestions: [], boardNow: room
        });
      });
    });
    (g.onDeck || []).forEach(function (p) {
      if (!p || !p.name) return;
      if (sheetHas(p.name)) return;
      add({
        reason: "extra", roomRaw: p.lastRoom || "on deck", staffRaw: p.name, name: p.name,
        shift: p.shift, student: !!p.student, kind: p.kind || "none",
        suggestions: [], boardNow: "on deck"
      });
    });
    return issues;
  }

  function placeUnsureChoice(item, loc) {
    if (!item || !item.name) return;
    if (!loc) return;
    if (!g.roomStaff) g.roomStaff = {};
    if (!g.roomStaff[loc.cat]) g.roomStaff[loc.cat] = {};
    var occ = g.roomStaff[loc.cat][loc.room];
    if (occ && occ.name && !namesMatch(occ.name, item.name)) {
      g.onDeck = g.onDeck || [];
      g.onDeck.unshift({
        name: occ.name, shift: occ.shift || "", kind: occ.kind || "none",
        lastRoom: loc.room, role: "freed", student: !!occ.student
      });
    }
    g.roomStaff[loc.cat][loc.room] = {
      name: item.name, shift: item.shift || "", kind: item.kind || breakKind(item.shift),
      closed: false, student: !!item.student, firstCase: item.firstCase || extractRoomTime(item.roomRaw) || ""
    };
    try {
      var es = typeof catEditState !== "undefined" && catEditState[loc.cat];
      if (es && es.deletedRooms && typeof es.deletedRooms.delete === "function") {
        es.deletedRooms.delete(loc.room);
        if (es.deletedEvents) es.deletedEvents[loc.room] = { deleted: false, ts: Date.now() };
      }
    } catch (e) {}
    var k = nameKey(item.name);
    g.onDeck = (g.onDeck || []).filter(function (p) {
      return nameKey(p.name) !== k;
    });
    refreshAfterAssign();
  }

  function removeExtra(item) {
    if (!item) return;
    if (item.cat && item.room && g.roomStaff[item.cat]) {
      g.roomStaff[item.cat][item.room] = { name: "", shift: "", kind: "none", closed: true };
      try {
        var es = typeof catEditState !== "undefined" && catEditState[item.cat];
        if (es && es.deletedRooms && typeof es.deletedRooms.add === "function") {
          es.deletedRooms.add(item.room);
          if (es.deletedEvents) es.deletedEvents[item.room] = { deleted: true, ts: Date.now() };
        }
      } catch (e) {}
    }
    var k = nameKey(item.name);
    g.onDeck = (g.onDeck || []).filter(function (p) { return nameKey(p.name) !== k; });
    refreshAfterAssign();
  }

  function applyShiftGuess(item, letter) {
    if (!item || !item.name) return;
    var lab = parseShiftLabel(letter) || { shift: letter, kind: breakKind(letter) };
    function patch(rec) {
      if (!rec) return;
      rec.shift = lab.shift;
      rec.kind = lab.kind || breakKind(lab.shift);
    }
    (g.onDeck || []).forEach(function (p) {
      if (!namesMatch(p.name, item.name)) return;
      patch(p);
      if (parseShiftLabel(p.lastRoom)) {
        p.lastRoom = "";
        if (p.role === "unplaced") p.role = "shift";
      }
    });
    Object.keys(g.roomStaff || {}).forEach(function (cat) {
      Object.keys(g.roomStaff[cat] || {}).forEach(function (room) {
        var rec = g.roomStaff[cat][room];
        if (rec && rec.name && namesMatch(rec.name, item.name)) patch(rec);
      });
    });
    refreshAfterAssign();
  }

  function showAssignReview(res, file) {
    var existing = document.getElementById("assign-review-overlay");
    if (existing) existing.remove();
    var items = (res.unsure || []).slice();
    if (!items.length) return;
    var preview = res.preview || [];
    var cells = res.cells || {};
    var ov = document.createElement("div");
    ov.id = "assign-review-overlay";
    ov.className = "assign-review-overlay";
    function esc(s) {
      return String(s || "").replace(/[&<>"]/g, function (c) {
        return { "&": "&" + "amp;", "<": "&" + "lt;", ">": "&" + "gt;", '"': "&" + "quot;" }[c];
      });
    }
    function colToN(s) {
      var n = 0;
      for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
      return n;
    }
    function nToCol(n) {
      var s = "";
      while (n > 0) {
        s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    }
    function cellMatches(val, q) {
      if (!q || !val) return false;
      var v = String(val).trim();
      if (!v) return false;
      function whole(hay, needle) {
        if (!needle) return false;
        var n = String(needle).trim();
        if (!n) return false;
        if (n.length <= 2) return hay.toUpperCase() === n.toUpperCase();
        var re = new RegExp("(^|[^A-Za-z])" + n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "([^A-Za-z]|$)", "i");
        return re.test(hay);
      }
      if (q.roomRaw && whole(v, q.roomRaw)) return true;
      var ln = lastName(q.name);
      if (ln && ln.length > 2 && whole(v, ln)) return true;
      return false;
    }
    function spreadHtml() {
      var refs = Object.keys(cells);
      if (!refs.length) return "";
      var maxC = 0, maxR = 0, minR = 99;
      refs.forEach(function (ref) {
        var m = ref.match(/^([A-Z]+)(\d+)$/);
        if (!m) return;
        maxC = Math.max(maxC, colToN(m[1]));
        maxR = Math.max(maxR, parseInt(m[2], 10));
        minR = Math.min(minR, parseInt(m[2], 10));
      });
      maxC = Math.min(maxC, 12);
      maxR = Math.min(maxR, 48);
      minR = Math.min(minR, 1);
      var html = '<table class="rev-spread"><tbody>';
      for (var r = minR; r <= maxR; r++) {
        html += "<tr>";
        for (var c = 1; c <= maxC; c++) {
          var ref = nToCol(c) + r;
          html += '<td data-ref="' + ref + '">' + esc(cells[ref] || "") + "</td>";
        }
        html += "</tr>";
      }
      return html + "</tbody></table>";
    }
    function orderedPreviewHtml(q) {
      var order = { NT: 0, ST: 1, "STE": 2, CCS: 3, FBC: 4, ENDO: 5, EP: 6, Mid: 7, WBF: 8, MN: 9 };
      var rows = preview.slice().sort(function (a, b) {
        var da = order[a.section] != null ? order[a.section] : 20;
        var db = order[b.section] != null ? order[b.section] : 20;
        if (da !== db) return da - db;
        return 0;
      });
      var html = "";
      var lastSec = "";
      rows.forEach(function (row) {
        if (skipRoomLabel(row.room) && !row.staff) return;
        if (!row.staff && !row.room) return;
        if (row.section !== lastSec) {
          lastSec = row.section;
          html += '<div class="rev-sec">' + esc(row.section) + "</div>";
        }
        var on = q && cellMatches(row.room + " " + row.staff, q);
        html += '<div class="rev-row' + (on ? " on" : "") + '"><span class="rev-room">' +
          esc(row.room) + '</span><span class="rev-staff">' + esc(row.staff) + "</span></div>";
      });
      return html || '<div class="rev-empty">No sheet rows to show.</div>';
    }
    function attachSheetZoom(viewport, inner) {
      var scale = 1, tx = 8, ty = 8;
      var min = 0.45, max = 4;
      inner.style.transformOrigin = "0 0";
      function apply() {
        inner.style.transform = "translate(" + tx + "px," + ty + "px) scale(" + scale + ")";
      }
      var pointers = {};
      var startDist = 0, startScale = 1, lastPan = null;
      function dist(a, b) {
        var dx = a.x - b.x, dy = a.y - b.y;
        return Math.sqrt(dx * dx + dy * dy);
      }
      function pts() {
        return Object.keys(pointers).map(function (k) { return pointers[k]; });
      }
      viewport.addEventListener("pointerdown", function (e) {
        e.preventDefault();
        try { viewport.setPointerCapture(e.pointerId); } catch (err) {}
        pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
        var p = pts();
        if (p.length >= 2) {
          startDist = dist(p[0], p[1]);
          startScale = scale;
          lastPan = null;
        } else {
          lastPan = { x: e.clientX, y: e.clientY };
        }
      });
      viewport.addEventListener("pointermove", function (e) {
        if (!pointers[e.pointerId]) return;
        e.preventDefault();
        pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
        var p = pts();
        if (p.length >= 2 && startDist > 8) {
          var ns = Math.min(max, Math.max(min, startScale * (dist(p[0], p[1]) / startDist)));
          var mid = { x: (p[0].x + p[1].x) / 2, y: (p[0].y + p[1].y) / 2 };
          var vr = viewport.getBoundingClientRect();
          var px = mid.x - vr.left, py = mid.y - vr.top;
          var k = ns / scale;
          tx = px - k * (px - tx);
          ty = py - k * (py - ty);
          scale = ns;
          apply();
        } else if (p.length === 1 && lastPan) {
          tx += e.clientX - lastPan.x;
          ty += e.clientY - lastPan.y;
          lastPan = { x: e.clientX, y: e.clientY };
          apply();
        }
      });
      function up(e) {
        delete pointers[e.pointerId];
        lastPan = null;
        var p = pts();
        if (p.length >= 2) {
          startDist = dist(p[0], p[1]);
          startScale = scale;
        }
      }
      viewport.addEventListener("pointerup", up);
      viewport.addEventListener("pointercancel", up);
      viewport.addEventListener("wheel", function (e) {
        e.preventDefault();
        var vr = viewport.getBoundingClientRect();
        var px = e.clientX - vr.left, py = e.clientY - vr.top;
        var ns = Math.min(max, Math.max(min, scale * (e.deltaY < 0 ? 1.1 : 0.9)));
        var k = ns / scale;
        tx = px - k * (px - tx);
        ty = py - k * (py - ty);
        scale = ns;
        apply();
      }, { passive: false });
      return {
        apply: apply,
        fit: function () {
          var vr = viewport.getBoundingClientRect();
          var w = inner.scrollWidth || inner.offsetWidth || 1;
          scale = Math.min(1.05, Math.max(0.5, (vr.width - 16) / w));
          tx = 8; ty = 8;
          inner.style.transition = "";
          apply();
        },
        panTo: function (el) {
          if (!el) return;
          var vr = viewport.getBoundingClientRect();
          var hr = el.getBoundingClientRect();
          tx += (vr.left + vr.width * 0.5) - (hr.left + hr.width * 0.5);
          ty += (vr.top + vr.height * 0.42) - (hr.top + hr.height * 0.5);
          inner.style.transition = "transform .28s ease";
          apply();
          setTimeout(function () { inner.style.transition = ""; }, 280);
        }
      };
    }

    var dateBit = res.date ? " · " + res.date : "";
    var fileBit = (file && file.name) ? " · " + file.name : (res.file ? " · " + res.file : "");
    var grid = spreadHtml();
    ov.innerHTML =
      '<div class="assign-review-card">' +
      '<div class="rev-handle"><span></span></div>' +
      "<h3>Check the sheet</h3>" +
      '<p class="rev-lead" id="rev-lead"></p>' +
      '<p class="rev-date">' + esc((res.kind || "sheet") + dateBit + fileBit) + " · pinch to zoom</p>" +
      '<div class="rev-split">' +
      (grid
        ? '<div class="rev-sheet-zoom" id="rev-zoom"><div class="rev-sheet-inner" id="rev-inner">' + grid + "</div></div>"
        : '<div class="rev-sheet" id="rev-list"></div>') +
      '<div class="rev-q" id="rev-q"></div></div>' +
      '<button type="button" class="rev-done" id="rev-done">Skip the rest</button></div>';
    document.body.appendChild(ov);
    var zoom = null;
    var zoomEl = document.getElementById("rev-zoom");
    var innerEl = document.getElementById("rev-inner");
    var creating = false;
    var createCat = "";
    if (zoomEl && innerEl) {
      zoom = attachSheetZoom(zoomEl, innerEl);
      setTimeout(function () { if (zoom) zoom.fit(); }, 30);
    }

    function createFormHtml(q) {
      var guessed = guessCatForLabel((q && q.roomRaw) || "");
      if (!createCat) createCat = (guessed && guessed.cat) || "nora";
      var name = stripRoomTime((q && q.roomRaw) || "") || (guessed && guessed.room) || "";
      var cats = "";
      (typeof CATEGORIES !== "undefined" ? CATEGORIES : []).forEach(function (c) {
        cats += '<button type="button" class="rev-cat' + (c.id === createCat ? " on" : "") + '" data-create-cat="' + esc(c.id) + '">' + esc(c.name) + "</button>";
      });
      return '<div class="rev-create" id="rev-create">' +
        '<div class="rev-create-label">Name the pill for today</div>' +
        '<input class="rev-create-input" id="rev-create-name" maxlength="24" value="' + esc(name) + '" autocomplete="off">' +
        '<div class="rev-create-label">Which suite?</div>' +
        '<div class="rev-create-cats">' + cats + "</div>" +
        '<button type="button" class="rev-create-go" data-create-go="1">Add for today</button></div>';
    }
    function bindCreateForm(q) {
      var box = document.getElementById("rev-create");
      if (!box || !q) return;
      box.querySelectorAll("[data-create-cat]").forEach(function (btn) {
        btn.onclick = function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          createCat = btn.getAttribute("data-create-cat") || createCat;
          box.querySelectorAll("[data-create-cat]").forEach(function (b) {
            b.classList.toggle("on", b.getAttribute("data-create-cat") === createCat);
          });
        };
      });
      var go = box.querySelector("[data-create-go]");
      if (go) {
        go.onclick = function (ev) {
          ev.preventDefault();
          ev.stopPropagation();
          var inp = document.getElementById("rev-create-name");
          var name = inp ? String(inp.value || "").replace(/\s+/g, " ").trim() : "";
          if (!name) {
            try { if (typeof showToast === "function") showToast("Name the room first"); } catch (e) {}
            return;
          }
          if (!createCat) createCat = "nora";
          var loc = null;
          try {
            if (typeof addCustomRoomForDay === "function") {
              loc = addCustomRoomForDay(createCat, name, {
                sheetRaw: q.roomRaw || name,
                who: q.name || "",
                source: "review"
              });
            }
          } catch (e) {}
          if (!loc) loc = { cat: createCat, room: name };
          var cur = items.shift();
          creating = false;
          createCat = "";
          if (cur) placeUnsureChoice(cur, loc);
          paint();
        };
      }
    }

    function paint() {
      var left = items.length;
      var q = items[0];
      var lead = document.getElementById("rev-lead");
      var qEl = document.getElementById("rev-q");
      var done = document.getElementById("rev-done");
      if (lead) lead.textContent = left ? left + " assignment" + (left === 1 ? "" : "s") + " need a tap" : "All set";
      if (done) done.textContent = left ? "Skip the rest" : "Done";
      var chips = "";
      var why = "";
      if (q) {
        if (q.reason === "mismatch") why = "Sheet vs board don't match";
        else if (q.reason === "missing") why = "On the sheet, not on the board";
        else if (q.reason === "extra") why = "On the board, not on the sheet";
        else if (q.reason === "shift") why = "Is this their shift?";
        else if (q.reason === "need-shift") why = "No shift letter found";
        else if (q.reason === "need-room") why = "Where are they assigned?";
        else why = "This label isn't a room we know";
        if (q.reason === "shift") {
          var gs = q.guessedShift || q.roomRaw || q.shift || "";
          chips += '<button type="button" class="rev-chip" data-shift="' + esc(gs) + '">Yes — ' + esc(gs) + " shift</button>";
        }
        if (q.reason === "need-shift") {
          ["D", "d", "M", "S", "Q", "W", "E", "N", "t", "e"].forEach(function (lettr) {
            chips += '<button type="button" class="rev-chip" data-shift="' + lettr + '">' + lettr + "</button>";
          });
        }
        if (q.reason !== "shift") {
          (q.suggestions || []).forEach(function (s) {
            if (!s || !s.room) return;
            chips += '<button type="button" class="rev-chip" data-cat="' + esc(s.cat) + '" data-room="' + esc(s.room) + '">' + esc(s.room) + "</button>";
          });
        }
        if (q.reason === "extra") {
          chips += '<button type="button" class="rev-chip" data-remove="1">Take off the board</button>';
          chips += '<button type="button" class="rev-chip deck" data-keep="1">Keep</button>';
        } else if (q.reason === "need-room") {
          chips += '<button type="button" class="rev-chip deck" data-deck="1">On deck — no room</button>';
        } else if (q.reason === "need-shift") {
          chips += '<button type="button" class="rev-chip deck" data-deck="1">No shift</button>';
        } else if (q.reason === "shift") {
          chips += '<button type="button" class="rev-chip deck" data-keep="1">Not a shift</button>';
        } else {
          chips += '<button type="button" class="rev-chip deck" data-deck="1">Keep on deck</button>';
        }
        if (q.reason !== "shift" && q.reason !== "need-shift" && q.reason !== "extra") {
          chips += '<button type="button" class="rev-chip create" data-create="1">Create a room…</button>';
        }
      }
      if (qEl) {
        qEl.innerHTML = q
          ? ('<div class="rev-why">' + esc(why) + "</div>" +
            '<div class="rev-ask">' + (q.reason === "need-room"
              ? "No room on the sheet"
              : ('Sheet said <strong>' + esc(q.roomRaw) + "</strong>")) + "</div>" +
            '<div class="rev-who">' + shiftPillHtml(q.shift) + " " + esc(q.name) + (q.student ? " ^" : "") + "</div>" +
            '<div class="rev-now">Board now: ' + esc(q.boardNow || "—") + "</div>" +
            '<div class="rev-chips">' + chips + "</div>" +
            (creating ? createFormHtml(q) : ""))
          : '<div class="rev-ask">Board is ready.</div>';
        qEl.querySelectorAll(".rev-chip").forEach(function (btn) {
          btn.onclick = function (ev) {
            ev.preventDefault();
            ev.stopPropagation();
            if (btn.getAttribute("data-create")) {
              creating = true;
              createCat = "";
              paint();
              var inp = document.getElementById("rev-create-name");
              if (inp) setTimeout(function () { try { inp.focus(); inp.select(); } catch (e) {} }, 40);
              return;
            }
            var cur = items.shift();
            if (!cur) { paint(); return; }
            creating = false;
            if (btn.getAttribute("data-remove")) removeExtra(cur);
            else if (btn.getAttribute("data-shift")) applyShiftGuess(cur, btn.getAttribute("data-shift"));
            else if (btn.getAttribute("data-keep") || btn.getAttribute("data-deck")) { /* leave */ }
            else placeUnsureChoice(cur, { cat: btn.getAttribute("data-cat"), room: btn.getAttribute("data-room") });
            paint();
          };
        });
        bindCreateForm(q);
      }
      if (innerEl) {
        innerEl.querySelectorAll("td.on").forEach(function (td) { td.classList.remove("on"); });
        var first = null;
        if (q) {
          var want = {};
          (q.cellRefs || []).forEach(function (ref) {
            if (ref) want[String(ref).toUpperCase()] = 1;
          });
          var useRefs = Object.keys(want).length > 0;
          innerEl.querySelectorAll("td").forEach(function (td) {
            var ref = (td.getAttribute("data-ref") || "").toUpperCase();
            var on = useRefs ? !!want[ref] : cellMatches(td.textContent, q);
            if (on) {
              td.classList.add("on");
              if (!first) first = td;
            }
          });
        }
        if (zoom && first) setTimeout(function () { zoom.panTo(first); }, 40);
      } else {
        var list = document.getElementById("rev-list");
        if (list) {
          list.innerHTML = orderedPreviewHtml(q);
          var hit = list.querySelector(".rev-row.on");
          if (hit) hit.scrollIntoView({ block: "center", behavior: "smooth" });
        }
      }
      if (done) {
        done.onclick = function (ev) {
          ev.preventDefault();
          ov.remove();
          try { if (typeof showToast === "function") showToast("Sheet applied"); } catch (e) {}
        };
      }
    }
    paint();
  }

  function showSheetViewer() {
    var sheet = g.lastSheet || (g.assignmentMeta && { cells: g.assignmentMeta.cells, date: g.assignmentMeta.date, file: g.assignmentMeta.file, kind: g.assignmentMeta.kind });
    if (!sheet || !sheet.cells || !Object.keys(sheet.cells).length) {
      try { if (typeof showToast === "function") showToast("No sheet uploaded yet"); } catch (e) {}
      return;
    }
    var existing = document.getElementById("sheet-view-overlay");
    if (existing) existing.remove();
    var cells = sheet.cells;
    function esc(s) {
      return String(s || "").replace(/[&<>"]/g, function (c) {
        return { "&": "&" + "amp;", "<": "&" + "lt;", ">": "&" + "gt;", '"': "&" + "quot;" }[c];
      });
    }
    function colToN(s) {
      var n = 0;
      for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
      return n;
    }
    function nToCol(n) {
      var s = "";
      while (n > 0) {
        s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    }
    var refs = Object.keys(cells);
    var maxC = 0, maxR = 0, minR = 99;
    refs.forEach(function (ref) {
      var m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) return;
      maxC = Math.max(maxC, colToN(m[1]));
      maxR = Math.max(maxR, parseInt(m[2], 10));
      minR = Math.min(minR, parseInt(m[2], 10));
    });
    maxC = Math.min(maxC, 12);
    maxR = Math.min(maxR, 48);
    minR = Math.min(minR, 1);
    var grid = '<table class="rev-spread"><tbody>';
    for (var r = minR; r <= maxR; r++) {
      grid += "<tr>";
      for (var c = 1; c <= maxC; c++) {
        var ref = nToCol(c) + r;
        grid += "<td>" + esc(cells[ref] || "") + "</td>";
      }
      grid += "</tr>";
    }
    grid += "</tbody></table>";
    var ov = document.createElement("div");
    ov.id = "sheet-view-overlay";
    ov.className = "assign-review-overlay";
    ov.innerHTML =
      '<div class="assign-review-card">' +
      '<div class="rev-handle"><span></span></div>' +
      "<h3>Uploaded sheet</h3>" +
      '<p class="rev-date">' + esc((sheet.kind || "sheet") + (sheet.date ? " · " + sheet.date : "") + (sheet.file ? " · " + sheet.file : "")) + " · pinch to zoom</p>" +
      '<div class="rev-sheet-zoom" id="view-zoom"><div class="rev-sheet-inner" id="view-inner">' + grid + "</div></div>" +
      '<button type="button" class="rev-done" id="view-sheet-done">Close</button></div>';
    document.body.appendChild(ov);
    var viewport = document.getElementById("view-zoom");
    var inner = document.getElementById("view-inner");
    var scale = 1, tx = 8, ty = 8;
    function apply() { inner.style.transform = "translate(" + tx + "px," + ty + "px) scale(" + scale + ")"; }
    inner.style.transformOrigin = "0 0";
    var pointers = {};
    var startDist = 0, startScale = 1, lastPan = null;
    function dist(a, b) { var dx = a.x - b.x, dy = a.y - b.y; return Math.sqrt(dx * dx + dy * dy); }
    function pts() { return Object.keys(pointers).map(function (k) { return pointers[k]; }); }
    viewport.addEventListener("pointerdown", function (e) {
      e.preventDefault();
      try { viewport.setPointerCapture(e.pointerId); } catch (err) {}
      pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
      var p = pts();
      if (p.length >= 2) { startDist = dist(p[0], p[1]); startScale = scale; lastPan = null; }
      else lastPan = { x: e.clientX, y: e.clientY };
    });
    viewport.addEventListener("pointermove", function (e) {
      if (!pointers[e.pointerId]) return;
      e.preventDefault();
      pointers[e.pointerId] = { x: e.clientX, y: e.clientY };
      var p = pts();
      if (p.length >= 2 && startDist > 8) {
        var ns = Math.min(4, Math.max(0.45, startScale * (dist(p[0], p[1]) / startDist)));
        var mid = { x: (p[0].x + p[1].x) / 2, y: (p[0].y + p[1].y) / 2 };
        var vr = viewport.getBoundingClientRect();
        var px = mid.x - vr.left, py = mid.y - vr.top;
        var k = ns / scale;
        tx = px - k * (px - tx); ty = py - k * (py - ty); scale = ns; apply();
      } else if (p.length === 1 && lastPan) {
        tx += e.clientX - lastPan.x; ty += e.clientY - lastPan.y;
        lastPan = { x: e.clientX, y: e.clientY }; apply();
      }
    });
    function up(e) { delete pointers[e.pointerId]; lastPan = null; }
    viewport.addEventListener("pointerup", up);
    viewport.addEventListener("pointercancel", up);
    setTimeout(function () {
      var vr = viewport.getBoundingClientRect();
      var w = inner.scrollWidth || inner.offsetWidth || 1;
      scale = Math.min(1.05, Math.max(0.5, (vr.width - 16) / w));
      tx = 8; ty = 8; apply();
    }, 30);
    function close() { ov.remove(); }
    document.getElementById("view-sheet-done").onclick = close;
    ov.addEventListener("click", function (e) { if (e.target === ov) close(); });
  }

  function showAssignSummary(res, file) {
    var existing = document.getElementById("assign-upload-overlay");
    if (existing) existing.remove();
    var ov = document.createElement("div");
    ov.id = "assign-upload-overlay";
    ov.className = "assign-upload-overlay";
    if (res && res.kind === "split") {
      ov.innerHTML =
        '<div class="assign-upload-card">' +
        "<h3>Roster merged</h3>" +
        "<p class='assign-meta'>" + (file ? file.name : "Name roster") + "</p>" +
        "<ul class='assign-stats'>" +
        "<li>" + ((res.people || []).length) + " names on the roster</li>" +
        "<li>" + (res.added || 0) + " added to the deck (not already placed)</li>" +
        "<li>" + ((res.calls || []).length) + " call slots</li>" +
        "</ul>" +
        "<p class='assign-meta'>Rooms stay as they are. Upload the assignment grid first if the board is still empty.</p>" +
        '<div class="assign-actions"><button type="button" class="assign-apply">Done</button></div></div>';
      document.body.appendChild(ov);
      function closeSplit(ev) {
        if (ev) { try { ev.preventDefault(); ev.stopPropagation(); } catch (e) {} }
        ov.remove();
      }
      ov.addEventListener("click", function (e) {
        if (e.target === ov || (e.target.closest && e.target.closest(".assign-apply"))) closeSplit(e);
      }, true);
      var doneSplit = ov.querySelector(".assign-apply");
      doneSplit.onclick = closeSplit;
      return;
    }
    var today = typeof todayStr === "function" ? todayStr() : "";
    var dateNote = res.date && today && dayKey(res.date) !== dayKey(today)
      ? '<div class="assign-warn">Loaded ' + prettyDay(res.date) + ". Today is " + prettyDay(today) + ".</div>"
      : "";
    var um = (res.unmatched || []).slice(0, 6).map(function (u) {
      return "<li>" + (u.room || "") + " — " + (u.staff || "") + "</li>";
    }).join("");
    var hideN = 0;
    var closedSet = {};
    (res.closed || []).forEach(function (c) {
      if (c && c.cat && c.room) closedSet[c.cat + "|" + c.room] = 1;
    });
    if (typeof CATEGORIES !== "undefined") {
      var named = {};
      (res.rooms || []).forEach(function (r) {
        if (r && r.cat && r.room && r.name && !r.closed) named[r.cat + "|" + r.room] = 1;
      });
      CATEGORIES.forEach(function (c) {
        (c.rooms || []).forEach(function (room) {
          if (!named[c.id + "|" + room] && !closedSet[c.id + "|" + room]) hideN++;
        });
      });
    }
    var ntRun = res.runners && res.runners.nt;
    var stRun = res.runners && res.runners.st;
    var runLine = (ntRun || stRun)
      ? "<li>Runners " + (ntRun ? "NT " + ntRun : "") + (ntRun && stRun ? " · " : "") + (stRun ? "ST " + stRun : "") + "</li>"
      : "";
    ov.innerHTML =
      '<div class="assign-upload-card">' +
      "<h3>Sheet applied</h3>" +
      "<p class='assign-meta'>" + (res.kind || "") + (res.date ? " · " + res.date : "") + (file ? " · " + file.name : "") + "</p>" +
      dateNote +
      "<ul class='assign-stats'>" +
      runLine +
      "<li>" + (res.rooms || []).filter(function (r) { return r && r.name && !r.closed; }).length + " rooms with a name</li>" +
      "<li>" + (res.closed || []).length + " marked CLOSED</li>" +
      (hideN ? "<li>" + hideN + " not on sheet (hidden)</li>" : "") +
      "<li>" + (res.late || []).length + " late (M+S)</li>" +
      "<li>" + (res.dinner || []).length + " dinner (Q+W+E)</li>" +
      "<li>" + ((res.onDeck || []).length) + " on deck (not in a room)</li>" +
      ((res.lateStays || []).length ? "<li>" + res.lateStays.length + " late stay</li>" : "") +
      ((res.calls || []).length ? "<li>" + res.calls.length + " call</li>" : "") +
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

  function prettyDay(k) {
    var p = String(dayKey(k) || "").split("-");
    if (p.length < 3 || !p[0]) return "";
    return parseInt(p[1], 10) + "/" + parseInt(p[2], 10) + "/" + p[0];
  }

  function sheetLoadProblem(res) {
    var today = "";
    try { today = dayKey(todayStr()); } catch (e) {}
    var kind = res && res.kind;
    if (kind !== "weekday" && kind !== "weekend") {
      if (kind === "split") {
        return {
          title: "This isn't a daily assignment sheet",
          body: "This is the name roster, not the weekday or weekend assignment grid. It will not place people into rooms.",
          allow: true,
          proceed: "Merge roster anyway"
        };
      }
      return {
        title: "This isn't a daily assignment sheet",
        body: "This file isn't a weekday or weekend assignment sheet, so it was not loaded.",
        allow: false,
        proceed: ""
      };
    }
    var sheetDay = dayKey(res.date);
    var dates = [];
    (res.dates || []).forEach(function (d) {
      var dk = dayKey(d);
      if (dk && dates.indexOf(dk) < 0) dates.push(dk);
    });
    if (!sheetDay) {
      return {
        title: "No date on this sheet",
        body: "This looks like an assignment sheet, but it has no date. Today is " + prettyDay(today) + ".",
        allow: true,
        proceed: "Load anyway"
      };
    }
    if (today && sheetDay !== today) {
      var span = dates.length > 1 ? dates.map(prettyDay).join(" and ") : prettyDay(sheetDay);
      var lead = dates.length > 1 ? ("Weekend file for " + span + ". ") : "";
      return {
        title: "This sheet is for a different day",
        body: lead + "Loading it would put " + prettyDay(sheetDay) + " on the board. Today is " + prettyDay(today) + ".",
        allow: true,
        proceed: "Load anyway"
      };
    }
    return null;
  }

  function confirmSheetLoad(problem, file) {
    return new Promise(function (resolve) {
      var existing = document.getElementById("assign-upload-overlay");
      if (existing) existing.remove();
      var ov = document.createElement("div");
      ov.id = "assign-upload-overlay";
      ov.className = "assign-upload-overlay";
      var proceed = problem.allow
        ? '<button type="button" class="assign-cancel" data-act="yes">' + problem.proceed + "</button>"
        : "";
      ov.innerHTML =
        '<div class="assign-upload-card">' +
        "<h3>" + problem.title + "</h3>" +
        '<div class="assign-warn">' + problem.body + "</div>" +
        "<p class='assign-meta'>" + (file && file.name ? file.name : "") + "</p>" +
        '<div class="assign-actions">' +
        '<button type="button" class="assign-apply" data-act="no">Don\'t load</button>' +
        proceed +
        "</div></div>";
      document.body.appendChild(ov);
      var settled = false;
      function finish(yes, ev) {
        if (settled) return;
        settled = true;
        if (ev) { try { ev.preventDefault(); ev.stopPropagation(); } catch (e) {} }
        ov.remove();
        resolve(!!yes);
      }
      ov.addEventListener("click", function (e) {
        var btn = e.target.closest && e.target.closest("[data-act]");
        if (btn) finish(btn.getAttribute("data-act") === "yes", e);
        else if (e.target === ov) finish(false, e);
      }, true);
    });
  }

  function finishAssignmentLoad(res, file) {
    applyAssignmentResult(res);
    if (res.kind === "split") {
      showAssignSummary(res, file);
      return;
    }
    if (res.kind !== "weekday" && res.kind !== "weekend") return;
    res.unsure = auditBoardVsSheet(res);
    if (res.unsure && res.unsure.length) showAssignReview(res, file);
    else showAssignSummary(res, file);
  }

  async function handleAssignmentFile(file) {
    if (!file) return;
    try {
      var buf = await file.arrayBuffer();
      var res = await parseAssignmentWorkbook(buf, file.name);
      var problem = sheetLoadProblem(res);
      if (problem) {
        var yes = await confirmSheetLoad(problem, file);
        if (!yes) return;
      }
      finishAssignmentLoad(res, file);
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
        "body.assigning .room-btn,body.assigning .weekend-room-btn{box-shadow:inset 0 0 0 1.5px rgba(122,78,45,.35);}" +
        ".staff-move{position:relative;z-index:2;padding:1px 3px;border-radius:8px;}" +
        "body.assigning .staff-move{pointer-events:none;}" +
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
        ".breakq-item.given{opacity:1;background:#E8DDD0;border-style:dashed;}" +
        ".breakq-item.given .staff-name{text-decoration:line-through;}" +
        ".breakq-had-label{font-size:10px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#9A6A38;margin:2px 0 4px;}" +
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
        ".relief-need-grid{display:grid;grid-template-columns:1fr;gap:6px;margin-top:6px;}" +
        ".relief-need{display:flex;align-items:center;flex-wrap:wrap;gap:4px 6px;min-width:0;width:100%;overflow:visible;background:#FDF6EC;border:1px solid rgba(160,98,42,.2);border-radius:8px;padding:8px 10px;font:inherit;color:#1E0E04;cursor:pointer;-webkit-appearance:none;appearance:none;text-align:left;}" +
        ".relief-need.planned{border-color:#7A4E2D;background:#F5E6D0;}" +
        ".relief-need .relief-room{flex:0 0 auto;font-size:11px;font-weight:800;white-space:nowrap;}" +
        ".relief-need .staff-name{min-width:0;overflow:visible;text-overflow:unset;white-space:normal;font-size:12px;font-weight:700;}" +
        ".relief-need .shift-pill{flex-shrink:0;}" +
        ".relief-need .relief-arrow{flex-shrink:0;font-size:11px;}" +
        ".relief-need .staff-name.missing{color:#9A6A38;font-weight:500;}" +
        ".relief-need-place{flex-shrink:0;font-size:9px;font-weight:800;padding:2px 5px;border-radius:5px;background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;margin-left:1px;}" +
        ".relief-wave-toggle{width:100%;margin:8px 0 0;padding:10px 12px;min-height:44px;border-radius:10px;border:1px solid rgba(160,98,42,.22);background:#FDF6EC;color:#7A4E2D;font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;text-align:left;cursor:pointer;}" +
        ".ondeck-out-toggle{min-height:44px;width:100%;margin:8px 0 4px;padding:10px 12px;border-radius:10px;border:1px solid rgba(160,98,42,.22);background:#FDF6EC;color:#7A4E2D;font-size:12px;font-weight:800;text-align:left;cursor:pointer;}" +
        ".ondeck-chip.later,.roster-pill.later{opacity:.82;}" +
        ".desk-undo-btn{width:100%;margin:0 0 8px;padding:10px 12px;border-radius:12px;border:1.5px solid #7A4E2D;background:#F5E6D0;color:#5A2E0A;font-size:13px;font-weight:800;cursor:pointer;}" +
        "#undo-close-bar{position:fixed;left:12px;right:12px;bottom:calc(18px + env(safe-area-inset-bottom,0px));z-index:700;padding:12px 14px;border:none;border-radius:14px;background:#2C1A0E;color:#FDF6EC;font-size:13px;font-weight:800;box-shadow:0 8px 24px rgba(30,14,4,.28);cursor:pointer;}" +
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
        ".deact-sheet-overlay{position:fixed;inset:0;background:rgba(30,14,4,.55);z-index:880;display:flex;align-items:flex-end;justify-content:center;padding:16px;}" +
        ".deact-sheet{background:#fff;border-radius:18px 18px 12px 12px;padding:16px 16px 18px;width:100%;max-width:420px;color:#1E0E04;box-shadow:0 16px 48px rgba(30,14,4,.22);}" +
        ".deact-kicker{font-size:11px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#C8781A;margin:0 0 6px;}" +
        ".deact-who{font-size:16px;font-weight:800;margin:0 0 8px;display:flex;align-items:center;gap:6px;flex-wrap:wrap;}" +
        ".deact-ask{font-size:13px;color:#7A4E2D;margin:0 0 12px;}" +
        ".deact-choice{width:100%;display:flex;flex-direction:column;align-items:flex-start;gap:2px;text-align:left;padding:12px 14px;margin:0 0 8px;border-radius:12px;border:1.5px solid rgba(160,98,42,.22);background:#FDF6EC;color:#1E0E04;cursor:pointer;font:inherit;}" +
        ".deact-choice-title{font-size:14px;font-weight:800;color:#1E0E04;}" +
        ".deact-choice-sub{font-size:12px;font-weight:600;color:#7A4E2D;}" +
        ".deact-cancel{width:100%;padding:12px;border:none;border-radius:12px;background:#F3EDE6;color:#7A4E2D;font-size:14px;font-weight:800;cursor:pointer;font:inherit;}" +
        ".sites-sheet-card{background:#fff;border-radius:16px;padding:16px 14px 14px;width:100%;max-width:380px;max-height:86vh;overflow-y:auto;-webkit-overflow-scrolling:touch;color:#1E0E04;box-shadow:0 16px 48px rgba(30,14,4,.22);}" +
        ".sites-choice{width:100%;display:flex;flex-direction:column;align-items:flex-start;gap:2px;text-align:left;padding:12px 14px;margin:0 0 8px;border-radius:12px;border:1.5px solid rgba(160,98,42,.22);background:#FDF6EC;color:#1E0E04;cursor:pointer;font:inherit;}" +
        ".sites-choice strong{font-size:15px;color:#1E0E04;}" +
        ".sites-choice span{font-size:12px;color:#7A4E2D;}" +
        ".relief-suggest{background:linear-gradient(135deg,#FFF0D8,#FFE4A0)!important;border-color:#C8781A!important;}" +
        ".relief-suggest strong{font-size:16px;}" +
        ".breaker-jobs{display:flex;flex-direction:column;gap:6px;}" +
        ".job-sec{font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#9A6A38;margin:8px 0 2px;}" +
        ".job-row{width:100%;display:flex;align-items:center;justify-content:space-between;gap:8px;text-align:left;padding:11px 12px;min-height:44px;border-radius:10px;border:1.5px solid rgba(160,98,42,.22);background:#fff;color:#1E0E04;cursor:pointer;font:inherit;}" +
        ".job-row .job-main{font-size:14px;font-weight:800;}" +
        ".job-row .job-sub{font-size:11px;font-weight:700;color:#7A4E2D;flex-shrink:0;}" +
        ".job-row.job-info{cursor:default;background:#FDF6EC;}" +
        ".job-row.job-info .job-main{text-decoration:line-through;font-weight:600;color:#7A4E2D;}" +
        ".job-row.job-late{background:#FFF0D8;border-color:#C8781A;}" +
        ".job-row.job-late .job-sub{color:#A05A10;}" +
        ".job-row.job-dinner{background:#7A4E2D;border-color:#5A3418;color:#FDF6EC;}" +
        ".job-row.job-dinner .job-main{color:#FDF6EC;}" +
        ".job-row.job-dinner .job-sub{color:#F5E0B0;}" +
        ".job-row.job-urgent{background:#FFF0F0;border-color:#C0392B;}" +
        ".job-row.job-urgent .job-sub{color:#C0392B;}" +
        ".job-empty{font-size:12px;color:#9A6A38;line-height:1.45;padding:4px 2px 2px;}" +
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
        ".room-btn.kind-late.done{background:#E8D4A8;border-color:#7A4E2D;opacity:.7;box-shadow:inset 0 0 0 2px rgba(122,78,45,.45);}" +
        ".room-btn.kind-late.done .staff-name{text-decoration:line-through;}" +
        ".room-btn.kind-dinner.done{background:#4A2E14;border-color:#2C1A0E;color:#FDF6EC;opacity:.55;box-shadow:none;}" +
        ".room-btn.kind-dinner.done .staff-name{text-decoration:line-through;color:#FDF6EC;}" +
        ".room-btn.kind-late.done::after,.room-btn.kind-dinner.done::after{content:' had it';font-size:9px;font-weight:900;letter-spacing:.02em;}" +
        ".staff-view-toggle{display:flex;gap:6px;margin:0 0 8px;}" +
        ".staff-view-btn{flex:1;padding:9px 10px;min-height:40px;border-radius:10px;border:1.5px solid rgba(160,98,42,.25);background:#FEF6EC;color:#7A4E2D;font-size:13px;font-weight:800;cursor:pointer;font-family:inherit;}" +
        ".staff-view-btn.on{background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;border-color:#7A4E2D;}" +
        "body.staff-view-jobs #board{display:none;}" +
        ".staff-jobs-slot{margin-top:8px;}" +
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
        ".sites-fn-tab{max-width:none;}" +
        ".assign-review-overlay{position:fixed;inset:0;background:rgba(30,14,4,.55);z-index:9200;display:flex;align-items:flex-end;justify-content:center;padding:10px;}" +
        ".assign-review-card{background:#fff;border-radius:18px 18px 12px 12px;padding:10px 14px 14px;width:100%;max-width:520px;max-height:92vh;display:flex;flex-direction:column;color:#1E0E04;box-shadow:0 16px 48px rgba(30,14,4,.22);}" +
        ".assign-review-card h3{margin:0;font-family:Georgia,'Times New Roman',serif;font-size:20px;}" +
        ".rev-handle{display:flex;justify-content:center;padding:2px 0 8px;}" +
        ".rev-handle span{width:40px;height:4px;border-radius:2px;background:rgba(122,78,45,.3);}" +
        ".rev-lead{font-size:12px;color:#7A4E2D;margin:2px 0 10px;}" +
        ".rev-split{display:flex;flex-direction:column;gap:10px;min-height:0;flex:1;overflow:hidden;}" +
        ".rev-sheet{flex:0 0 auto;max-height:32vh;overflow-y:auto;-webkit-overflow-scrolling:touch;border:1px solid rgba(160,98,42,.2);border-radius:12px;padding:8px;background:#FDF6EC;}" +
        ".rev-date{font-size:11px;color:#9A6A38;margin:-6px 0 8px;}" +
        ".rev-sheet-zoom{flex:0 0 auto;height:36vh;overflow:hidden;touch-action:none;border:1px solid rgba(160,98,42,.2);border-radius:12px;background:#FDF6EC;position:relative;}" +
        ".rev-sheet-inner{transform-origin:0 0;will-change:transform;display:inline-block;padding:6px;}" +
        ".rev-spread{border-collapse:collapse;font-size:9px;line-height:1.25;white-space:nowrap;}" +
        ".rev-spread td{border:1px solid rgba(160,98,42,.16);padding:3px 5px;max-width:118px;overflow:hidden;text-overflow:ellipsis;color:#1E0E04;}" +
        ".rev-spread td.on{background:#F5E6D0;box-shadow:inset 0 0 0 1.5px #7A4E2D;font-weight:800;}" +
        ".rev-q{flex:1;min-height:0;overflow-y:auto;}" +
        ".rev-sec{font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#9A6A38;margin:8px 0 4px;}" +
        ".rev-row{display:flex;gap:8px;padding:3px 4px;border-radius:6px;font-size:11px;line-height:1.3;}" +
        ".rev-row.on{background:#F5E6D0;box-shadow:inset 0 0 0 1px #7A4E2D;}" +
        ".rev-room{flex:0 0 38%;font-weight:800;color:#7A4E2D;}" +
        ".rev-staff{flex:1;color:#1E0E04;}" +
        ".rev-ask{font-size:13px;color:#7A4E2D;margin:0 0 6px;}" +
        ".rev-who{font-size:16px;font-weight:800;margin:0 0 6px;display:flex;align-items:center;gap:6px;}" +
        ".rev-why{font-size:11px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#C8781A;margin:0 0 6px;}" +
        ".rev-now{font-size:12px;color:#7A4E2D;margin:0 0 10px;}" +
        ".rev-chips{display:flex;flex-wrap:wrap;gap:6px;}" +
        ".rev-chip{padding:10px 12px;border-radius:10px;border:1.5px solid rgba(160,98,42,.25);background:#FDF6EC;color:#1E0E04;font-size:13px;font-weight:800;cursor:pointer;font:inherit;}" +
        ".rev-chip.deck{background:#fff;color:#7A4E2D;}" +
        ".rev-chip.create{background:#1E0E04;color:#FDF6EC;border-color:#1E0E04;}" +
        ".rev-create{margin-top:10px;padding:10px;border-radius:12px;border:1.5px solid rgba(160,98,42,.28);background:#FDF6EC;}" +
        ".rev-create-label{font-size:10px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;color:#9A6A38;margin:0 0 6px;}" +
        ".rev-create-input{width:100%;box-sizing:border-box;margin:0 0 10px;padding:11px 12px;border-radius:10px;border:1.5px solid rgba(160,98,42,.3);background:#fff;color:#1E0E04;font:inherit;font-size:15px;font-weight:800;}" +
        ".rev-create-cats{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 10px;}" +
        ".rev-cat{padding:8px 10px;border-radius:999px;border:1.5px solid rgba(160,98,42,.25);background:#fff;color:#7A4E2D;font-size:12px;font-weight:800;cursor:pointer;font:inherit;}" +
        ".rev-cat.on{background:#7A4E2D;color:#fff;border-color:#7A4E2D;}" +
        ".rev-create-go{width:100%;padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#C8781A,#E8A020);color:#fff;font-size:14px;font-weight:800;cursor:pointer;font:inherit;}" +
        ".rev-done{width:100%;margin-top:12px;padding:12px;border:none;border-radius:12px;background:linear-gradient(135deg,#7A4E2D,#9A6A38);color:#fff;font-size:14px;font-weight:800;cursor:pointer;}" +
        ".rev-empty{font-size:12px;color:#9A6A38;}";
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

  function renderBreakerJobList() {
    var panel = document.getElementById("breaker-panel");
    if (!panel) return;
    applyStaffView();
    var view = getStaffView();
    var html = '<div class="breaker-panel-inner breaker-jobs">';
    html += staffViewToggleHtml();
    html += '<div><div class="my-site-name">Breaker</div><div class="my-site-label">' +
      (view === "jobs" ? "Jobs for this window · switch to Board anytime" : "Whole board · switch to Jobs for a due list") +
      "</div></div>";
    if (view === "jobs") html += staffJobsHtml();
    else html += '<div class="job-empty">Use the board below. Jobs lists who is due.</div>';
    html += "</div>";
    panel.innerHTML = html;
  }

  g.roomStaff = g.roomStaff || {};
  g.onDeck = g.onDeck || [];
  g.reliefPlan = g.reliefPlan || {};
  g.lateStays = g.lateStays || [];
  g.calls = g.calls || [];
  g.assignmentMeta = g.assignmentMeta || null;
  g.parseAssignmentWorkbook = parseAssignmentWorkbook;
  g.applyAssignmentResult = applyAssignmentResult;
  g.handleAssignmentFile = handleAssignmentFile;
  g.auditBoardVsSheet = auditBoardVsSheet;
  g.normalizeRoom = normalizeRoom;
  g.rebuildRoomIndex = function () { ROOM_INDEX = null; buildRoomIndex(); };
  g.parseStaff = parseStaff;
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
  g.todayRoster = todayRoster;
  g.findStaffByName = findStaffByName;
  g.renderBreakerJobList = renderBreakerJobList;
  g.setStaffView = setStaffView;
  g.getStaffView = getStaffView;
  g.applyStaffView = applyStaffView;
  g.paintCrnaJobs = paintCrnaJobs;
  g.staffViewToggleHtml = staffViewToggleHtml;
  g.longestIdleDeck = longestIdleDeck;
  g.chipName = chipName;
  g.lastName = lastName;
  g.nameKey = nameKey;
  g.sendOccupantToDeck = sendOccupantToDeck;
  g.closeRoomIfEmpty = closeRoomIfEmpty;
  g.plantHeldInto = plantHeldInto;
  g.setHeldRoom = setHeldRoom;
  g.finishPlant = finishPlant;
  g.askOccupiedPlant = askOccupiedPlant;
  g.clearHeld = clearHeld;
  g.heldFrom = heldFrom;
  g.occupantOf = occupantOf;
  g.roomIsActive = roomIsActive;

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
