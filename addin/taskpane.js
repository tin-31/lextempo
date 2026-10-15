/* LexTempo Word add-in — rà trích dẫn căn cứ pháp lý ngay trong tài liệu.
   Chạy OFFLINE bằng chỉ mục bãi bỏ nhúng (window.LEXDATA). Tuỳ chọn: gọi API /scan.
   §0: chỉ báo "đã bãi bỏ" khi khớp CHÍNH XÁC phạm vi; ngoài phủ = "chưa xác minh". */
(function () {
  "use strict";
  var D = window.LEXDATA || { repeals: [], names: {}, covered: [], library: [] };
  var $ = function (id) { return document.getElementById(id); };

  function norm(s) {
    return String(s || "").toLowerCase().normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/[^a-z0-9]+/g, " ").trim();
  }
  // dấu ngắt câu → " # " (tên không khớp vượt ranh giới câu) — như web
  function nv(s) {
    return " " + String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d")
      .replace(/[;.:!?“”"()]+/g, " # ").replace(/[^a-z0-9#]+/g, " ").trim() + " ";
  }
  var CBHEAD = /CÔNG BÁO\/Số[^\n]*?\d{1,2}-\d{1,2}-\d{4}(?:\s+\d{1,4}(?=\s))?/g;   // dòng đầu trang Công báo
  var COVERED = {}; D.covered.forEach(function (l) { COVERED[l] = 1; });
  // Tên luật ĐÃ BIẾT: chỉ mục bãi bỏ + thư viện + tên thu từ corpus (id null = biết tên, chưa có số);
  // xếp DÀI→NGẮN (sort ổn định: tên chỉ mục đứng trước khi trùng).
  var NAMES = [];
  Object.keys(D.names || {}).forEach(function (k) { NAMES.push({ na: nv(k), id: D.names[k] }); });
  (D.lib_names || []).forEach(function (p) { NAMES.push({ na: nv(p[0]), id: p[1] || null }); });
  // luật CÙNG TÊN nhiều phiên bản: mặc định = lựa chọn của bộ tách máy chủ (D.name_default), giống web
  var NDEF = D.name_default || {};
  NAMES.sort(function (a, b) { return b.na.length - a.na.length || (NDEF[b.na.trim()] === b.id) - (NDEF[a.na.trim()] === a.id); });

  // ── phân giải luật đích (cùng thuật toán với web/app.html — xem chú thích ở đó) ──────────────
  var NUMQ = / (\d{1,3}) (\d{4}) (qh\d{1,2}) /g, SO = /^so (\d{1,3}) (\d{4}) (qh\d{1,2}) /,
      YR = /^(?:ngay \d{1,2} thang \d{1,2} )?(?:nam )?((?:19|20)\d{2}) /,
      NONLAW = /^ (?:cua )?(?:nghi dinh|thong tu|quyet dinh|nghi quyet|phap lenh|chi thi|van ban|hien phap|luat nay|bo luat nay) /,
      QHRES = /^ (?:cua )?nghi quyet so \d{1,3} \d{4} qh\d{1,2} /,         // Nghị quyết của QUỐC HỘI: có số → giữ
      ENUM = /^(?: (?:va|hoac|den|cac|diem [a-z]\d?|khoan \d+[a-z]?|dieu \d+[a-z]?|\d+[a-z]?))+(?= )/,
      ASAMEND = /^ (?:duoc (?:sua doi|bo sung|thay the)(?: bo sung)? )?nhu sau /,                                                 // "… Điều 9 như sau:" = điều của VB bị sửa
      AMOPEN = / (?:sua doi|bo sung|bai bo) (?:(?! \d{1,3} \d{4} [a-z]).)*$/;
  function qid(m) { return m[1] + "/" + m[2] + "/" + m[3].toUpperCase(); }
  // số hiệu / năm KHAI cho tên luật → tên trơn về sau kế thừa; khối căn cứ (trước "Điều 1.") ưu tiên
  var NA_IDS = {}; NAMES.forEach(function (L) { if (L.id) (NA_IDS[L.na] = NA_IDS[L.na] || []).push(L.id); });
  var YRD = /^(?:ngay \d{1,2} thang \d{1,2} )?nam ((?:19|20)\d{2}) /;
  // "sửa đổi/bổ sung/bãi bỏ" mở mà chưa đóng bằng số hiệu (từ "Căn cứ" gần nhất) → số là của LUẬT SỬA ĐỔI
  function declared(text) {
    var d = {}, h = text.search(/(^|[\n\r])[ \t]*Điều 1\./), S = nv(text), H = h >= 0 ? nv(text.slice(0, h)).length : -1;
    NAMES.forEach(function (L) {
      for (var k = S.indexOf(L.na); k >= 0; k = S.indexOf(L.na, k + 1)) {
        var rest = S.slice(k + L.na.length), n = rest.match(SO), y = rest.match(YRD); if (!n && !y) continue;
        var c = S.lastIndexOf(" can cu ", k);
        if (AMOPEN.test(S.slice(c >= 0 ? c : Math.max(0, k - 300), k + 1))) continue;
        var cc = H >= 0 ? k < H : (c >= 0 && k - c < 300), e = d[L.na] = d[L.na] || {};
        if (cc && n) e.num = e.num || qid(n); else if (cc && y) e.year = e.year || y[1]; else if (n) e.body = e.body || qid(n);
      }
    });
    return d;
  }
  function pick(L, e) {                      // phiên bản theo khai báo; null = biết tên, chưa rõ số hiệu
    if (!e) return L.id;
    if (e.num) return e.num;
    if (e.year) {
      var ids = NA_IDS[L.na] || [];
      for (var i = 0; i < ids.length; i++) if (ids[i].indexOf("/" + e.year + "/") >= 0) return ids[i];
      return e.body && e.body.indexOf("/" + e.year + "/") >= 0 ? e.body : null;
    }
    return L.id;                             // số ở thân chỉ dùng khi khớp năm căn cứ
  }
  function resolveLaw(win, last, decl) {
    var s = nv(win), hit = false, best = null, bk = last ? -1 : Infinity, blk = {}, m;
    function better(k) { return last ? k > bk : k < bk; }
    NUMQ.lastIndex = 0;
    while ((m = NUMQ.exec(s))) { if (better(m.index)) { hit = true; best = qid(m); bk = m.index; } NUMQ.lastIndex = m.index + 1; }
    for (var j = 0; j < NAMES.length; j++) {
      var L = NAMES[j];
      for (var k = s.indexOf(L.na); k >= 0; k = s.indexOf(L.na, k + 1)) {
        if (!better(k) || (blk[k] || 0) > L.na.length) continue;
        var rest = s.slice(k + L.na.length), sn = rest.match(SO), y = rest.match(YR);
        if (sn) { hit = true; best = qid(sn); bk = k; continue; }
        if (y && (!L.id || L.id.indexOf("/" + y[1] + "/") < 0)) { blk[k] = Math.max(blk[k] || 0, L.na.length); continue; }
        hit = true; best = y ? L.id : pick(L, decl && decl[L.na]); bk = k;
      }
    }
    return { id: best, hit: hit };
  }

  // ── khớp (điều/khoản/điểm) với chỉ mục bãi bỏ (bao hàm phạm vi) ───────────
  function check(lawId, art, cl, pt) {
    var exact = null, partial = null;
    for (var i = 0; i < D.repeals.length; i++) {
      var e = D.repeals[i];
      if (e.law !== lawId || e.art !== String(art)) continue;
      if (e.cl === null) { exact = e; break; }                 // cả Điều bị bãi
      if (cl && e.cl === String(cl)) {
        if (e.pt === null) { exact = e; break; }                // cả khoản bị bãi
        if (pt && e.pt === String(pt)) { exact = e; break; }    // đúng điểm
      }
      partial = partial || e;   // Điều có bãi bỏ nhưng trích dẫn không trùng đúng phần
    }
    if (exact) return { tier: "dead", e: exact };
    if (partial && !cl) return { tier: "partial", e: partial }; // trích cả Điều mà điều có phần bị bãi
    return null;
  }

  // ── bóc mọi trích dẫn "[(điểm x) (khoản y)] Điều z … <luật>" ─────────────
  // tail DỪNG trước trích dẫn kế ("điểm a"/"khoản 2"/"Điều 5" — có số/chữ) để không nuốt căn cứ sau nó;
  // không dừng ở "một số điều của Luật…"
  // số Điều có hậu tố chữ ("Điều 34a") như web; đuôi qua được xuống dòng (dán từ PDF), dừng ở . ;
  var CITE = /(?:điểm\s+([a-zđ])\s+)?(?:khoản\s+(\d+)\s+)?Điều\s+(\d+[a-zđ]?)((?:(?!điểm\s+[a-zđ]\s|khoản\s+\d|Điều\s+\d)[^.;]){0,120})/gi;
  function maskRemoved(t) {                 // trích dẫn TRONG cụm từ đang bị thay thế/bãi bỏ → che (giữ vị trí)
    return t.replace(/((?:[Tt]hay(?:\s+thế)?|[Bb]ãi\s+bỏ)\s+(?:các\s+)?cụm\s+từ\s*“)([^”]{1,1500})(”)/g,
      function (_, a, b, c) { return a + b.replace(/[^\n]/g, " ") + c; });
  }
  function extract(text) {
    text = maskRemoved(String(text).normalize("NFC").replace(CBHEAD, " "));   // NFD ("Unicode tổ hợp") trước đây ra 0 trích dẫn
    var out = [], m, decl = declared(text);
    CITE.lastIndex = 0;
    while ((m = CITE.exec(text)) !== null) {
      var pt = m[1] || null, cl = m[2] || null, art = m[3], tail = m[4] || "";
      // tiêu đề "Điều N." đầu dòng/đoạn (kể cả “Điều N. …” trong điều sửa đổi) là điều của CHÍNH văn bản /
      // văn bản bị sửa ("." nằm NGOÀI tail nên xét ký tự ngay sau "Điều N")
      if (!pt && !cl && text.charAt(m.index + m[0].length - tail.length) === "." &&
          /(^|[\n\r])[ \t]*[“"]?$/.test(text.slice(Math.max(0, m.index - 80), m.index))) continue;
      // bỏ phần liệt kê ", 48 và 49" rồi xét: ngay sau là văn bản KHÁC ("Nghị định này/số …") hoặc "như sau:"
      // → không phải luật; còn lại: đoạn SAU trước, rỗng thì 40 ký tự TRƯỚC (không vượt qua tiêu đề "Điều N.")
      var an = nv(tail).replace(ENUM, ""), r = { id: null, hit: true };
      if (!(NONLAW.test(an) && !QHRES.test(an)) && !ASAMEND.test(an)) {
        r = resolveLaw(tail, false, decl);
        if (!r.hit) r = resolveLaw(text.slice(Math.max(0, m.index - 40), m.index)
          .replace(/[\s\S]*[\n\r][ \t]*[“"]?Điều \d+[a-zđ]?\.(?=\s)/, ""), true, decl);
      }
      var law = r.id;
      var quote = m[0].replace(/\s+/g, " ").trim();
      var loc = (pt ? "điểm " + pt + " " : "") + (cl ? "khoản " + cl + " " : "") + "Điều " + art;
      out.push({ quote: quote, loc: loc, art: art, cl: cl, pt: pt, law: law });
    }
    return out;
  }

  // văn bản bị THAY TOÀN BỘ (supersession.json, mỗi cặp có câu nguyên văn)
  var TRAPS = {}; (D.traps || []).forEach(function (t) { TRAPS[t.old] = t; });
  var TODAY = new Date().toISOString().slice(0, 10);
  function classify(c) {
    var t = c.law && TRAPS[c.law];
    if (t && t.on <= TODAY)
      return { tier: "dead", msg: "LUẬT ĐÃ BỊ THAY TOÀN BỘ bởi " + t.new + " (từ " + t.on + ")" +
        (t.exc ? ", trừ " + t.exc : "") + " — đối chiếu văn bản mới" };
    var r = classify0(c);
    if (t) r.msg += " · sẽ bị " + t.new + " thay toàn bộ từ " + t.on;
    return r;
  }
  function classify0(c) {
    if (!c.law) return { tier: "unk", msg: "không xác định được luật đích (thiếu số hiệu/tên rõ)" };
    if (!COVERED[c.law]) return { tier: "unk", msg: "luật " + c.law + " NGOÀI phủ — chưa xác minh (không có nghĩa còn hiệu lực)" };
    var r = check(c.law, c.art, c.cl, c.pt);
    if (r && r.tier === "dead") return { tier: "dead", msg: "ĐÃ BÃI BỎ bởi " + r.e.by + " (từ " + r.e.on + ")", e: r.e };
    if (r && r.tier === "partial") return { tier: "partial", msg: "một phần Điều này đã bị bãi (khoản " + r.e.cl + " bởi " + r.e.by + ") — trích cả Điều, cần kiểm lại", e: r.e };
    return { tier: "ok", msg: "chưa thấy bãi bỏ trong chỉ mục (" + c.law + ")" };
  }

  // ── tô sáng trong Word ───────────────────────────────────────────────────
  function paintColor(t) { return t === "dead" ? "#FFCDD2" : t === "partial" ? "#FFF3C4" : "#C8E6C9"; }
  var _painted = [];   // các quote đã tô (để xoá lại)
  function highlight(items) {
    var flagged = items.filter(function (it) { return it.tier === "dead" || it.tier === "partial"; });
    return Word.run(function (ctx) {
      flagged.forEach(function (it) {
        if (it.quote.length < 4) { it._res = null; return; }
        it._res = ctx.document.body.search(it.quote, { matchCase: false, matchWholeWord: false });
        it._res.load("items");
      });
      return ctx.sync().then(function () {
        _painted = [];
        flagged.forEach(function (it) {
          if (!it._res) return;
          _painted.push(it.quote);
          it._res.items.forEach(function (rg) {
            try { rg.font.highlightColor = paintColor(it.tier); }
            catch (e) { try { rg.font.color = "#B3261E"; } catch (e2) {} }
          });
        });
        return ctx.sync();
      });
    });
  }
  function clearPaint() {
    var quotes = _painted.slice();
    return Word.run(function (ctx) {
      var reslist = quotes.map(function (q) {
        var r = ctx.document.body.search(q, { matchCase: false, matchWholeWord: false });
        r.load("items"); return r;
      });
      return ctx.sync().then(function () {
        reslist.forEach(function (r) {
          r.items.forEach(function (rg) { try { rg.font.highlightColor = "NoColor"; } catch (e) {} });
        });
        return ctx.sync();
      });
    }).then(function () { _painted = []; }).catch(function () {});
  }

  // ── render bảng kết quả ──────────────────────────────────────────────────
  function render(items) {
    var out = $("out"), sum = $("sum"), empty = $("empty");
    empty.hidden = true;
    var dead = items.filter(function (x) { return x.tier === "dead"; });
    var part = items.filter(function (x) { return x.tier === "partial"; });
    var unk = items.filter(function (x) { return x.tier === "unk"; });
    var ok = items.filter(function (x) { return x.tier === "ok"; });
    sum.hidden = false;
    sum.innerHTML =
      '<span class="pill dead">' + dead.length + ' đã bãi bỏ</span>' +
      (part.length ? '<span class="pill" style="color:#8a6d00">' + part.length + ' cần kiểm</span>' : "") +
      '<span class="pill ok">' + ok.length + ' chưa thấy bãi</span>' +
      '<span class="pill">' + unk.length + ' ngoài phủ</span>';
    var order = { dead: 0, partial: 1, unk: 2, ok: 3 };
    items.sort(function (a, b) { return order[a.tier] - order[b.tier]; });
    out.innerHTML = items.map(function (x) {
      var cls = x.tier === "dead" ? "dead" : (x.tier === "partial" || x.tier === "unk") ? "unk" : "";
      return '<div class="find ' + cls + '">' +
        '<div class="loc">' + esc(x.loc) + (x.law ? ' · ' + esc(x.law) : "") + '</div>' +
        '<span class="q">“' + esc(x.quote) + '”</span>' +
        '<div class="verdict">' + esc(x.msg) + '</div></div>';
    }).join("");
  }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  // ── API tuỳ chọn ─────────────────────────────────────────────────────────
  function viaApi(url, text) {
    return fetch(url, { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: text, as_of: new Date().toISOString().slice(0, 10) }) })
      .then(function (r) { return r.json(); });
  }

  // ── chạy ─────────────────────────────────────────────────────────────────
  function run() {
    $("scan").disabled = true; $("scan").textContent = "Đang rà…";
    Word.run(function (ctx) {
      var body = ctx.document.body; body.load("text");
      return ctx.sync().then(function () { return body.text; });
    }).then(function (text) {
      var items = extract(text).map(function (c) {
        var v = classify(c); c.tier = v.tier; c.msg = v.msg; return c;
      });
      render(items);
      return highlight(items);
    }).catch(function (e) {
      $("out").innerHTML = '<div class="find unk"><div class="verdict">Lỗi rà: ' + esc(e.message || e) + '</div></div>';
    }).then(function () {
      $("scan").disabled = false; $("scan").textContent = "Rà văn bản đang soạn";
    });
  }

  Office.onReady(function (info) {
    if (!info || info.host !== Office.HostType.Word) {
      $("cov").textContent = "Mở trong Microsoft Word để dùng.";
      return;
    }
    $("cov").textContent = "Phủ " + D.covered.length + " luật rà sâu · toàn văn " + D.library.length + " luật";
    $("nrep") && ($("nrep").textContent = D.repeals.length);
    $("ncov") && ($("ncov").textContent = D.covered.length);
    $("ncov2") && ($("ncov2").textContent = D.covered.length);
    $("eng") && ($("eng").textContent = D.engine);
    $("scan").onclick = run;
    $("clear").onclick = function () { clearPaint(); };
  });
})();
