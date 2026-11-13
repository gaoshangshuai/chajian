// ==UserScript==
// @name         TikTok Studio 数据导出助手 (播放/赞/评论/分享)
// @namespace    tts.exporter
// @version      1.4.0
// @description  在 TikTok Studio「作品」页或任意账号主页（tiktok.com/@用户名）自动滚动抓取作品数据，可下拉选择月份（抓到该月自动停止），导出 Excel/CSV：标题、链接、发布时间、赞、评论、分享、播放
// @match        https://www.tiktok.com/tiktokstudio/*
// @match        https://www.tiktok.com/tiktokstudio/content*
// @match        https://www.tiktok.com/@*
// @match        https://www.tiktok.com/*
// @run-at       document-start
// @grant        none
// ==/UserScript==

/*
 * 用法一（推荐，Tampermonkey）：
 *   安装 Tampermonkey 扩展 → 新建脚本 → 粘贴本文件全部内容保存 → 打开
 *   https://www.tiktok.com/tiktokstudio/content → 右上角出现悬浮面板
 *
 * 用法二（免安装，控制台）：
 *   打开 TikTok Studio 作品页 → F12 → Console → 粘贴本文件全部内容回车
 *   （控制台方式首屏前几条可能缺失，脚本会自动从页面 DOM 兜底补齐；
 *     建议粘贴后先刷新一次页面——刷新前脚本不会保存，请用 Tampermonkey 方式）
 *
 * 说明：精确数字来自拦截页面自身的 item_list 接口响应（非界面上的 224K 缩写）。
 */

(function () {
  'use strict';
  if (window.__TTS_EXPORTER__) return;
  window.__TTS_EXPORTER__ = true;

  var state = {
    items: new Map(),   // id -> {id,title,time,likes,comments,shares,views,seq,src:'api'|'dom'}
    hasMore: null,      // null=未知 true/false
    running: false,
    seq: 0             // 抓取顺序计数器
  };

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function log(msg) { console.log('[TTS导出] ' + msg); }

  // ---------------- 1. 网络拦截 ----------------
  function num(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'string' && v !== '' && !isNaN(+v)) return +v;
    return null;
  }

  // 时间归一化：任何格式统一转成"秒"级 Unix 时间戳，避免毫秒/秒混用导致比较失效
  function normTime(t) {
    if (t === null || t === undefined || t === '') return null;
    if (typeof t === 'number') return t < 1e12 ? t : Math.floor(t / 1000);
    if (typeof t === 'string') {
      var s = t.trim();
      if (/^\d{10}$/.test(s)) return +s;
      if (/^\d{13}$/.test(s)) return Math.floor(+s / 1000);
      if (/^\d+$/.test(s)) return +s;
      var d = new Date(s);
      if (!isNaN(d.getTime())) return Math.floor(d.getTime() / 1000);
      var m = s.match(/(\d{4})[年\-\/.](\d{1,2})[月\-\/.](\d{1,2})/);
      if (m) return Math.floor(new Date(+m[1], +m[2] - 1, +m[3]).getTime() / 1000);
      return null;
    }
    return null;
  }
  function pick(o, keys) {
    if (!o || typeof o !== 'object') return undefined;
    for (var i = 0; i < keys.length; i++) if (o[keys[i]] !== undefined && o[keys[i]] !== null) return o[keys[i]];
    return undefined;
  }

  function ingest(json, url) {
    var found = 0;
    var seen = new Set();
    (function walk(node, depth) {
      if (!node || typeof node !== 'object' || depth > 14 || seen.size > 5000) return;
      if (Array.isArray(node)) {
        for (var i = 0; i < node.length; i++) walk(node[i], depth + 1);
        return;
      }
      // 记录 has_more / hasMore
      if (typeof node.has_more === 'boolean') state.hasMore = node.has_more;
      if (typeof node.hasMore === 'boolean') state.hasMore = node.hasMore;

      var explicit = 'item_id' in node || 'video_id' in node || 'aweme_id' in node;
      var rawId = pick(node, ['item_id', 'video_id', 'aweme_id', 'id']);
      var idStr = (typeof rawId === 'string' || typeof rawId === 'number') ? String(rawId) : '';
      if (idStr && /^\d+$/.test(idStr) && idStr.length >= (explicit ? 10 : 15) && !seen.has(idStr)) {
        var st = (node.video_statistics || node.video_stats || node.statistics || node.stats || node.statsV2 || node);
        var views = num(pick(st, ['play_count', 'view_count', 'playCount', 'viewCount']));
        var likes = num(pick(st, ['digg_count', 'like_count', 'diggCount', 'likeCount']));
        var comments = num(pick(st, ['comment_count', 'commentCount']));
        var shares = num(pick(st, ['share_count', 'shareCount']));
        var ct = pick(node, ['create_time', 'create_time_unix', 'createTime', 'post_time']);
        // 必须像"作品"：有播放量，且有互动数据或发布时间（排除音乐等同样带数字 id 的对象）
        var isVideoItem = (views !== null) && (likes !== null || comments !== null || shares !== null || ct !== undefined);
        if (isVideoItem) {
          // 主页模式下：校验作者，非本主页发布的作品一律丢弃
          if (isProfilePage()) {
            var targetU = (getUsername() || '').toLowerCase();
            var au = node.author || node.authorInfo || null;
            var auName = au ? (au.uniqueId || au.unique_id || au.uniqueIdStr || '') : '';
            if (auName && targetU && String(auName).toLowerCase() !== targetU) return;
          }
          seen.add(idStr);
          upsert(idStr, {
            id: idStr,
            title: String(pick(node, ['title', 'item_title', 'video_title', 'desc']) || '').trim(),
            time: normTime(ct),
            views: views, likes: likes, comments: comments, shares: shares,
            src: 'api'
          });
          found++;
        }
      }
      for (var k in node) {
        if (Object.prototype.hasOwnProperty.call(node, k)) walk(node[k], depth + 1);
      }
    })(json, 0);
    if (found) log('已从接口捕获作品，累计 ' + state.items.size + ' 条');
  }

  // 合并写入：新数据只补充缺失字段，不会用残缺记录覆盖完整记录
  // seq 记录抓取顺序（即页面里的展示顺序，时间倒序），用于月份边界判断
  function upsert(id, rec) {
    var ex = state.items.get(id);
    if (!ex) {
      rec.seq = ++state.seq;
      state.items.set(id, rec);
      return;
    }
    ['title', 'time', 'views', 'likes', 'comments', 'shares'].forEach(function (k) {
      if ((ex[k] === null || ex[k] === undefined || ex[k] === '') && rec[k] !== null && rec[k] !== undefined && rec[k] !== '') ex[k] = rec[k];
    });
    if (rec.src === 'api') ex.src = 'api';
  }

  // 只接受"作品列表"接口（主页 post/item_list、Studio creator/item_list），
  // 明确排除搜索/推荐/相关视频等会混入别人作品的接口
  var RESP_RE = /item_list|item\/list|creator/i;
  var RESP_DENY_RE = /search|related|recommend|suggest|find|feed|comment|challenge|music/i;

  function isProfilePage() { return /^\/@/.test(location.pathname); }

  function shouldIngestUrl(url) {
    var u = String(url || '');
    if (!RESP_RE.test(u)) return false;
    if (RESP_DENY_RE.test(u)) return false;
    return true;
  }

  // fetch hook
  var origFetch = window.fetch;
  if (origFetch) {
    window.fetch = function () {
      var args = arguments;
      var p = origFetch.apply(this, args);
      try {
        var url = '';
        if (typeof args[0] === 'string') url = args[0];
        else if (args[0] && args[0].url) url = args[0].url;
        if (url && shouldIngestUrl(url)) {
          p.then(function (res) {
            try {
              res.clone().json().then(function (j) { ingest(j, url); }).catch(function () {});
            } catch (e) {}
          }).catch(function () {});
        }
      } catch (e) {}
      return p;
    };
  }

  // XHR hook
  try {
    var origOpen = XMLHttpRequest.prototype.open;
    var origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (m, u) {
      try { this.__ttsUrl = String(u || ''); } catch (e) {}
      return origOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function () {
      var self = this;
      this.addEventListener('load', function () {
        try {
          if (self.__ttsUrl && shouldIngestUrl(self.__ttsUrl) && self.responseText) {
            var j = JSON.parse(self.responseText);
            if (j) ingest(j, self.__ttsUrl);
          }
        } catch (e) {}
      });
      return origSend.apply(this, arguments);
    };
  } catch (e) {}

  // 页面初始数据兜底（document-start 时可能还没有，稍后再试）
  function tryUniversalData() {
    try {
      var el = document.getElementById('__UNIVERSAL_DATA_FOR_REHYDRATION__') ||
               document.querySelector('script#__UNIVERSAL_DATA_FOR_REHYDRATION__');
      if (el && el.textContent) ingest(JSON.parse(el.textContent));
    } catch (e) {}
  }
  setTimeout(tryUniversalData, 3000);
  setTimeout(tryUniversalData, 10000);

  // ---------------- 2. 用户名 ----------------
  function findUniqueId(node, depth) {
    if (!node || typeof node !== 'object' || depth > 10) return null;
    if (Array.isArray(node)) {
      for (var i = 0; i < node.length; i++) {
        var r = findUniqueId(node[i], depth + 1);
        if (r) return r;
      }
      return null;
    }
    var u = node.uniqueId || node.unique_id || node.uniqueIdStr;
    if (typeof u === 'string' && /^[\w.\-]{1,30}$/.test(u)) return u;
    for (var k in node) {
      if (Object.prototype.hasOwnProperty.call(node, k)) {
        var r2 = findUniqueId(node[k], depth + 1);
        if (r2) return r2;
      }
    }
    return null;
  }
  function getUsername() {
    // 主页模式下直接取地址栏用户名
    var pm = location.pathname.match(/^\/@([^\/?&#]+)/);
    if (pm) return pm[1];
    try {
      var el = document.getElementById('__UNIVERSAL_DATA_FOR_REHYDRATION__');
      if (el) { var u = findUniqueId(JSON.parse(el.textContent), 0); if (u) return u; }
    } catch (e) {}
    var as = document.querySelectorAll('a[href^="/@"]');
    for (var i = 0; i < as.length; i++) {
      var m = (as[i].getAttribute('href') || '').match(/^\/@([^\/?&#]+)/);
      if (m) return m[1];
    }
    return null;
  }

  // ---------------- 3. DOM 兜底采集 ----------------
  function parseAbbr(s) {
    var m = String(s).trim().match(/^([\d.,]+)\s*([KMB万kmb])?$/i);
    if (!m) return null;
    var v = parseFloat(m[1].replace(/,/g, ''));
    if (isNaN(v)) return null;
    var unit = (m[2] || '').toUpperCase();
    if (unit === 'K') v *= 1e3;
    else if (unit === 'M') v *= 1e6;
    else if (unit === 'B') v *= 1e9;
    else if (m[2] === '万') v *= 1e4;
    return Math.round(v);
  }

  function harvestDOM() {
    var as = document.querySelectorAll('a[href*="/video/"]');
    for (var i = 0; i < as.length; i++) {
      var href = as[i].getAttribute('href') || '';
      var m = href.match(/\/video\/(\d{10,25})/);
      if (!m) continue;
      // 主页模式下：只采集本主页名下的视频链接（排除弹窗推荐/相关视频等别人的作品）
      if (isProfilePage()) {
        var hu = (href.match(/^\/@([^\/?&#]+)/) || [])[1];
        var target = (getUsername() || '').toLowerCase();
        if (!hu || (target && hu.toLowerCase() !== target)) continue;
      }
      var id = m[1];
      var exist = state.items.get(id);
      if (exist && exist.src === 'api') continue;

      // 向上找包含数字统计的行容器
      var row = as[i], nums = [];
      for (var up = 0; up < 8 && row.parentElement; up++) {
        row = row.parentElement;
        if (row.getElementsByTagName('a').length >= 1 && row.textContent.length > 30) {
          var all = row.textContent.match(/\d[\d.,]*\s*[KMBkmb万]?/g) || [];
          nums = all.map(parseAbbr).filter(function (x) { return x !== null; });
          if (nums.length >= 2) break;
        }
      }
      var d = new Date(0);
      var rec = {
        id: id,
        title: (as[i].textContent || '').trim(),
        time: null,
        views: null, likes: null, comments: null, shares: null,
        src: 'dom'
      };
      // 界面列顺序一般是：播放、赞、评论、分享
      if (nums.length >= 4) { rec.views = nums[0]; rec.likes = nums[1]; rec.comments = nums[2]; rec.shares = nums[3]; }
      else if (nums.length === 3) { rec.views = nums[0]; rec.likes = nums[1]; rec.comments = nums[2]; }
      else if (nums.length === 2) { rec.views = nums[0]; rec.likes = nums[1]; }
      // 时间
      var tm = row.textContent.match(/(\d{4})[年\-\/.](\d{1,2})[月\-\/.](\d{1,2})/);
      if (tm) rec.time = normTime(tm[1] + '-' + tm[2] + '-' + tm[3]);
      // API 数据优先（走统一合并逻辑）
      upsert(id, rec);
    }
  }

  // ---------------- 4. 自动滚动加载 ----------------
  function findScrollContainer() {
    var best = null, bestH = 0;
    var divs = document.querySelectorAll('div');
    for (var i = 0; i < divs.length; i++) {
      var d = divs[i];
      var oy = '';
      try { oy = getComputedStyle(d).overflowY; } catch (e) {}
      if (!/^(auto|scroll)$/.test(oy)) continue;
      if (d.scrollHeight - d.clientHeight < 200) continue;
      if (d.scrollHeight > bestH) { best = d; bestH = d.scrollHeight; }
    }
    return best;
  }

  function clickLoadMore() {
    var btns = document.querySelectorAll('button, [role="button"]');
    for (var i = 0; i < btns.length; i++) {
      var t = (btns[i].textContent || '').trim();
      if (/加载更多|load more|查看更多/i.test(t) && btns[i].offsetParent !== null) {
        try { btns[i].click(); return true; } catch (e) {}
      }
    }
    return false;
  }

  async function crawl() {
    state.running = true;
    var stable = 0, last = -1;
    for (var i = 0; i < 600 && state.running; i++) {
      harvestDOM();
      clickLoadMore();
      var el = findScrollContainer();
      var before = state.items.size;
      if (el) {
        el.scrollTop = Math.min(el.scrollTop + Math.max(el.clientHeight * 0.9, 400), el.scrollHeight);
        if (el.scrollHeight - el.scrollTop - el.clientHeight < 5) el.scrollTop = el.scrollHeight;
      }
      window.scrollTo(0, document.body.scrollHeight);
      await sleep(900);
      if (state.items.size === before && before === last) stable++; else stable = 0;
      last = state.items.size;
      // 选定月份：以发布日期为准。按抓取顺序（页面展示顺序，时间倒序）取"最新载入的3条"，
      // 若这3条全部早于所选月份，说明该月已完整抓完，立即停止。
      // （不用"任意3条"判断，避免主页置顶的旧视频误触发）
      var selEl = document.getElementById('tts-month');
      var ymSel = selEl ? (selEl.value || '') : '';
      if (ymSel) {
        var startTs = monthStartTs(ymSel);
        var arr = [];
        state.items.forEach(function (r) { arr.push(r); });
        arr.sort(function (a, b) { return (b.seq || 0) - (a.seq || 0); });
        var tail = arr.slice(0, 3);
        var oldTail = 0, inMonth = 0;
        for (var ti = 0; ti < tail.length; ti++) {
          if (tail[ti].time && tail[ti].time < startTs) oldTail++;
        }
        state.items.forEach(function (r) {
          if (r.time && monthKeyOf(r) === ymSel) inMonth++;
        });
        refreshMonthOptions();
        if (oldTail >= 3 && tail.length >= 3) {
          log(ymSel + ' 数据抓全（' + inMonth + ' 条，以发布日期为准），提前停止');
          break;
        }
      } else {
        refreshMonthOptions();
        log('抓取中… 已捕获 ' + state.items.size + ' 条' + (state.hasMore === false ? '（服务端已无更多）' : ''));
      }
      if (state.hasMore === false && stable >= 3) break;
      if (stable >= 10) break;
    }
    harvestDOM();
    refreshMonthOptions();
    state.running = false;
    var selEl2 = document.getElementById('tts-month');
    var ymSel2 = selEl2 ? (selEl2.value || '') : '';
    if (ymSel2) {
      var c = 0;
      state.items.forEach(function (r) { if (monthKeyOf(r) === ymSel2) c++; });
      log('抓取完成：' + ymSel2 + ' 共 ' + c + ' 条');
    } else {
      log('抓取完成，共 ' + state.items.size + ' 条作品');
    }
    // 唯一的界面提示：抓取完成
    var stEl = document.getElementById('tts-status');
    if (stEl) {
      stEl.textContent = '已成功下载，请点击"导出Excel"';
      stEl.style.display = 'block';
    }
    return state.items.size;
  }

  // ---------------- 5. 月份筛选 ----------------
  function parseMonthInput(s) {
    s = (s || '').trim();
    if (!s) return null; // 全部
    var m = s.match(/^(\d{4})\s*[-\/年.]\s*(\d{1,2})月?$/);
    if (m) return m[1] + '-' + ('0' + +m[2]).slice(-2);
    var m2 = s.match(/^(\d{1,2})月$/);
    if (m2) {
      var now = new Date();
      return now.getFullYear() + '-' + ('0' + +m2[1]).slice(-2);
    }
    return 'INVALID';
  }

  function fmtTime(t) {
    if (t === null || t === undefined || t === '') return '';
    var d = null;
    if (typeof t === 'number') d = new Date(t < 1e12 ? t * 1000 : t);
    else if (typeof t === 'string') {
      if (/^\d{10}$/.test(t)) d = new Date(+t * 1000);
      else if (/^\d{13}$/.test(t)) d = new Date(+t);
      else d = new Date(t);
    }
    if (!d || isNaN(d.getTime())) return '';
    var p = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  function monthKeyOf(rec) {
    var s = fmtTime(rec.time);
    return s ? s.slice(0, 7) : '';
  }

  // 刷新月份下拉框：近 24 个月 + 已捕获数据的条数统计
  function refreshMonthOptions() {
    var sel = document.getElementById('tts-month');
    if (!sel) return;
    var cur = sel.value;
    var counts = {};
    state.items.forEach(function (r) {
      var ym = monthKeyOf(r);
      if (ym) counts[ym] = (counts[ym] || 0) + 1;
    });
    var now = new Date();
    var html = '<option value="">全部作品</option>';
    for (var k = 0; k < 24; k++) {
      var d = new Date(now.getFullYear(), now.getMonth() - k, 1);
      var ym = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
      var label = ym.replace('-', '年') + '月';
      if (counts[ym]) label += '（' + counts[ym] + '条）';
      else if (state.items.size) label += '（0条）';
      html += '<option value="' + ym + '">' + label + '</option>';
    }
    sel.innerHTML = html;
    sel.value = (cur && sel.querySelector('option[value="' + cur + '"]')) ? cur : '';
  }

  function monthStartTs(ym) {
    var p = ym.split('-');
    return new Date(+p[0], +p[1] - 1, 1).getTime() / 1000;
  }

  function buildRows() {
    var month = parseMonthInput(document.getElementById('tts-month') ? document.getElementById('tts-month').value : '');
    if (month === 'INVALID') { alert('月份格式不正确，请输入如 2026-10，留空表示导出全部'); return null; }
    var username = getUsername();
    var list = [];
    state.items.forEach(function (rec) {
      var ym = monthKeyOf(rec);
      if (month && ym !== month) return;
      list.push(rec);
    });
    list.sort(function (a, b) { return (b.time || 0) - (a.time || 0); });
    var header = ['Video title', 'Video link', 'Post time', 'Total views', 'Total likes', 'Total comments', 'Total shares'];
    var rows = [header];
    list.forEach(function (r) {
      rows.push([
        r.title || '',
        'https://www.tiktok.com/@' + (username || 'user') + '/video/' + r.id,
        fmtTime(r.time),
        r.views === null ? '' : r.views,
        r.likes === null ? '' : r.likes,
        r.comments === null ? '' : r.comments,
        r.shares === null ? '' : r.shares
      ]);
    });
    return rows;
  }

  // ---------------- 6. 导出 ----------------
  function loadScript(src) {
    return new Promise(function (res, rej) {
      var s = document.createElement('script');
      s.src = src;
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }
  async function ensureXLSX() {
    if (window.XLSX) return;
    try { await loadScript('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'); } catch (e) {}
    if (!window.XLSX) {
      try { await loadScript('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'); } catch (e) {}
    }
    if (!window.XLSX) throw new Error('导出组件加载失败（网络拦截），请稍后再试一次');
  }

  function stamp() {
    var d = new Date(), p = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes());
  }

  async function exportExcel() {
    var rows = buildRows();
    if (!rows) return;
    if (rows.length <= 1) { alert('没有符合条件的数据。请先点击「开始抓取」，或确认月份输入。'); return; }
    try {
      await ensureXLSX();
      var ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!cols'] = [{ wch: 40 }, { wch: 46 }, { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 12 }];
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'TikTok数据');
      var month = (document.getElementById('tts-month').value || '').trim();
      XLSX.writeFile(wb, 'TikTok作品数据_' + (month ? month + '_' : '全部_') + stamp() + '.xlsx');
      log('已导出 Excel：' + (rows.length - 1) + ' 条');
    } catch (e) {
      alert(e.message);
    }
  }

  // ---------------- 7. 悬浮面板 ----------------
  function buildPanel() {
    var box = document.createElement('div');
    box.id = 'tts-exporter-panel';
    box.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647;background:#fff;color:#111;'
      + 'border:1px solid #ddd;border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.18);'
      + 'padding:12px;width:250px;font:12px/1.6 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif;';
    var isProfile = /^\/@/.test(location.pathname);
    var title = isProfile
      ? '📊 TikTok 主页数据导出（@' + (getUsername() || '...') + '）'
      : '📊 TikTok Studio 数据导出';
    box.innerHTML =
      '<div id="tts-title" style="font-weight:700;font-size:13px;margin-bottom:8px;">' + title + '</div>'
      + '<div style="display:flex;gap:6px;margin-bottom:8px;">'
      + '  <select id="tts-month" style="flex:1;min-width:0;padding:5px 8px;border:1px solid #ccc;border-radius:6px;font-size:12px;background:#fff;">'
      + '    <option value="">全部作品</option>'
      + '  </select>'
      + '</div>'
      + '<div style="display:flex;gap:6px;">'
      + '  <button id="tts-start" style="flex:1;padding:7px 0;border:0;border-radius:6px;background:#fe2c55;color:#fff;cursor:pointer;font-weight:600;">开始抓取</button>'
      + '  <button id="tts-xlsx" style="flex:1;padding:7px 0;border:0;border-radius:6px;background:#07c160;color:#fff;cursor:pointer;font-weight:600;">导出 Excel</button>'
      + '</div>'
      + '<div id="tts-status" style="margin-top:10px;font-size:12px;font-weight:600;color:#07c160;display:none;"></div>';

    document.body.appendChild(box);

    function showDone() {
      var st = document.getElementById('tts-status');
      if (st) {
        st.textContent = '已成功下载，请点击"导出Excel"';
        st.style.display = 'block';
      }
    }
    window.__ttsShowDone = showDone;

    box.querySelector('#tts-start').onclick = function () {
      if (state.running) return;
      var st = document.getElementById('tts-status');
      if (st) st.style.display = 'none';
      crawl();
    };
    box.querySelector('#tts-xlsx').onclick = exportExcel;
  }

  function init() {
    if (document.body) {
      buildPanel();
      harvestDOM();
      refreshMonthOptions();
      // SPA 内切换账号/页面时清空数据
      var lastPath = location.pathname;
      setInterval(function () {
        if (location.pathname !== lastPath) {
          lastPath = location.pathname;
          state.items.clear();
          state.hasMore = null;
          refreshMonthOptions();
          if (/^\/@/.test(lastPath)) {
            log('已切换到 @' + (getUsername() || '') + ' 的主页，数据已清空，点击「开始抓取」');
            var t = document.getElementById('tts-title');
            if (t) t.textContent = '📊 TikTok 主页数据导出（@' + (getUsername() || '...') + '）';
          }
        }
      }, 1000);
    }
    else setTimeout(init, 500);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
