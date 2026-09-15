/* LexTempo Word add-in — rà trích dẫn căn cứ pháp lý ngay trong tài liệu.
   Chạy OFFLINE bằng chỉ mục bãi bỏ nhúng (window.LEXDATA). Tuỳ chọn: gọi API /scan.
   §0: chỉ báo "đã bãi bỏ" khi khớp CHÍNH XÁC phạm vi; ngoài phủ = "chưa xác minh". */
(function () {
  "use strict";
  var D = window.LEXDATA || { repeals: [], names: {}, covered: [], library: [] };
  var $ = function (id) { return document.getElementById(id); };

  function norm(s) {
    return String(s || "").toLowerCase().normalize("NFD")
      .replace(/[̀-ͯ]/g, "").replace(/đ/g, "d").replace(/\s+/g, " ").trim();
  }
  var COVERED = {}; D.covered.forEach(function (l) { COVERED[l] = 1; });
  var NAMES = {}; Object.keys(D.names || {}).forEach(function (k) { NAMES[norm(k)] = D.names[k]; });

  // ── phân giải luật đích từ đoạn văn quanh trích dẫn ──────────────────────
  function resolveLaw(ctx) {
    var m = ctx.match(/\b(\d{1,3}\/\d{4}\/QH\d{2})\b/);
    if (m) return m[1];
    var n = norm(ctx), best = null, bestLen = 0;
    for (var key in NAMES) {
      if (n.indexOf(key) >= 0 && key.length > bestLen) { best = NAMES[key]; bestLen = key.length; }
    }
    return best;
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
  // tail DỪNG trước trích dẫn kế (điểm/khoản/Điều) để không nuốt căn cứ sau nó
  var CITE = /(?:điểm\s+([a-zđ])\s+)?(?:khoản\s+(\d+)\s+)?Điều\s+(\d+)((?:(?!điểm\s|khoản\s|Điều\s)[^.;\n]){0,120})/gi;
  function extract(text) {
    var out = [], m;
    CITE.lastIndex = 0;
    while ((m = CITE.exec(text)) !== null) {
      var pt = m[1] || null, cl = m[2] || null, art = m[3], tail = m[4] || "";
      var law = resolveLaw(tail) || resolveLaw(text.slice(Math.max(0, m.index - 40), m.index));
      var quote = m[0].replace(/\s+/g, " ").trim();
      var loc = (pt ? "điểm " + pt + " " : "") + (cl ? "khoản " + cl + " " : "") + "Điều " + art;
      out.push({ quote: quote, loc: loc, art: art, cl: cl, pt: pt, law: law });
    }
    return out;
  }

  function classify(c) {
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
