/*!
 * team-progress.js  高经团队周进度核算(纯 JS,依赖 xlsx.bundle.js)
 * 流程(与业务手工流程一致):
 *  1) 上传"有效订单池"Excel(审核后导出的有效订单集合,强基/升学可多个合并)
 *  2) 上传"高经业绩"底表(含 高经 列 与 用户id 列)
 *  3) 用订单池的"账号ID"匹配底表的"用户id",取出该行的高经/基地/基地负责人/经理
 *  4) 数透: 高经团队 × 订单金额(求和)
 */
(function (global) {
  'use strict';

  var POOL_FILES = [];   // [{name, rows}]   有效订单池(可多个)
  var MAP = null;        // {name, cols, byId, byGh, rows, gjCount}
  var lastResult = null;

  // ---------- 工具 ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function str(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function toFloat(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isFinite(n) ? n : 0;
  }
  function round2(n) { return Math.round(n * 100) / 100; }
  function fmtMoney(n) {
    return toFloat(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function dtVal(d) {
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
      'T' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }
  function inRange(dt, startStr, endStr) {
    var s = startStr ? new Date(startStr) : null;
    var e = endStr ? new Date(endStr) : null;
    if (!s && !e) return true;
    if (!dt) return false;
    if (s && dt < s) return false;
    if (e && dt > e) return false;
    return true;
  }
  function log(msg, cls) {
    var el = document.getElementById('log');
    if (!el) return;
    var line = document.createElement('div');
    line.className = cls || '';
    line.textContent = '[' + new Date().toLocaleTimeString() + '] ' + msg;
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
  }

  // 是否为表头行(有效订单池导出文件表头会重复出现,需要跳过)
  function isHeaderRow(r) {
    var c0 = str(r && r[0]);
    return c0 === '账号ID' || c0 === '订单ID';
  }
  function isOrderPoolRows(rows) {
    for (var i = 0; i < Math.min(3, rows.length); i++) {
      var m = (rows[i] || []).map(function (x) { return str(x); }).join(',');
      if (m.indexOf('账号ID') >= 0 || m.indexOf('订单ID') >= 0) return true;
    }
    return false;
  }

  // ---------- 读取 ----------
  async function readFile(file) {
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(buf, { cellDates: true });
    return wb.SheetNames.map(function (sn) {
      return { sn: sn, rows: XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' }) };
    });
  }

  // ---------- 底表解析(自动识别列) ----------
  // 识别规则:表头行取"高经"列;用户id列优先表头命中(用户id/账号/urs/c),否则按列值 urs- 命中率
  function detectTeamCols(rows) {
    var hIdx = -1, header = [];
    for (var i = 0; i < Math.min(5, rows.length); i++) {
      var m = (rows[i] || []).map(function (x) { return str(x); });
      if (m.join(',').indexOf('高经') >= 0) { hIdx = i; header = m; break; }
    }
    if (hIdx < 0) return null;                       // 没有"高经"列 → 不是高经底表
    var nCol = 0;
    rows.slice(hIdx, hIdx + 50).forEach(function (r) { nCol = Math.max(nCol, (r || []).length); });

    function colOf(re) {
      for (var k = 0; k < header.length; k++) if (re.test(header[k])) return k;
      return -1;
    }
    var cGj = colOf(/高经/);
    var cBase = colOf(/^基地$/);
    var cHead = colOf(/基地负责人|负责人/);
    var cMgr = colOf(/经理/);
    var cSales = colOf(/销售姓名|姓名/);
    var cGh = colOf(/工号/);
    var cUid = colOf(/用户id|用户ID|账号id|账号ID|^c$|urs/i);
    if (cUid < 0) {                                  // 按列值命中 urs- 比例推断
      var best = -1, bestRate = 0;
      for (var c = 0; c < nCol; c++) {
        var nonEmpty = 0, hit = 0;
        for (var i2 = hIdx + 1; i2 < Math.min(rows.length, hIdx + 60); i2++) {
          var v = str((rows[i2] || [])[c]);
          if (!v) continue;
          nonEmpty++;
          if (/^urs-/.test(v)) hit++;
        }
        if (nonEmpty >= 3 && hit / nonEmpty > bestRate) { bestRate = hit / nonEmpty; best = c; }
      }
      if (best >= 0 && bestRate >= 0.5) cUid = best;
    }
    return { headerRow: hIdx, gj: cGj, uid: cUid, base: cBase, head: cHead, mgr: cMgr, sales: cSales, gh: cGh, nCol: nCol };
  }

  function buildTeamMap(rows, cols, fileName) {
    var byId = {}, gjSet = {}, nData = 0, conflict = 0;
    for (var i = cols.headerRow + 1; i < rows.length; i++) {
      var r = rows[i] || [];
      var gj = str(r[cols.gj]);
      if (!gj) continue;
      nData++;
      gjSet[gj] = 1;
      var uid = cols.uid >= 0 ? str(r[cols.uid]) : '';
      if (!uid) continue;
      var info = {
        gj: gj,
        base: cols.base >= 0 ? str(r[cols.base]) : '',
        head: cols.head >= 0 ? str(r[cols.head]) : '',
        mgr: cols.mgr >= 0 ? str(r[cols.mgr]) : ''
      };
      if (byId[uid] && byId[uid].gj !== gj) conflict++;
      if (!byId[uid]) byId[uid] = info;
    }
    return { name: fileName, cols: cols, byId: byId,
             nData: nData, conflict: conflict,
             uidCount: Object.keys(byId).length, gjCount: Object.keys(gjSet).length };
  }

  // ---------- 订单池归一行(纯函数: 传入 [{name, rows}] ) ----------
  function collectPoolRows(files) {
    var out = [];
    (files || []).forEach(function (f) {
      f.rows.forEach(function (r) {
        if (!r || isHeaderRow(r)) return;
        var id = str(r[0]);
        if (!id) return;
        out.push({ src: f.name, row: r, uid: id, amount: toFloat(r[4]),
                   orderId: str(r[3]), itemId: str(r[10]), phone: str(r[2]),
                   channel: str(r[20]), gonghao: str(r[25]), name: str(r[1]),
                   dt: AuditCore.parseDt(r[7]) || AuditCore.parseDt(r[5]),
                   perfTime: str(r[7]), payTime: str(r[5]) });
      });
    });
    return out;
  }

  function dedupe(rows, mode) {
    if (mode === 'none' || !mode) return { rows: rows, removed: 0 };
    var seen = {}, out = [], removed = 0;
    rows.forEach(function (o) {
      var k;
      if (mode === 'orderItem') k = o.orderId + '|' + o.itemId;
      else if (mode === 'order') k = o.orderId;
      else if (mode === 'phone') k = o.phone;
      else k = '';
      if (!k) { out.push(o); return; }
      if (seen[k]) { removed++; return; }
      seen[k] = 1;
      out.push(o);
    });
    return { rows: out, removed: removed };
  }

  // ---------- 核心:匹配 + 数透(纯函数,便于自测) ----------
  function compute(poolObjs, map, params) {
    var p = params || {};
    // 1) 时间过滤
    var inTime = poolObjs.filter(function (o) {
      if (p.start || p.end) return inRange(o.dt, p.start || '', p.end || '');
      return true;
    });
    // 2) 去重
    var dd = dedupe(inTime, p.dedupeMode);
    var kept = dd.rows;

    // 3) 匹配高经
    var gjAgg = {}, unmatched = [], matched = 0, custAll = {};
    kept.forEach(function (o) {
      var info = map.byId[o.uid];
      if (!info) { unmatched.push(o); return; }
      matched++;
      custAll[o.uid] = 1;
      o.gj = info.gj; o.base = info.base; o.head = info.head; o.mgr = info.mgr;
      if (!gjAgg[info.gj]) gjAgg[info.gj] = { gj: info.gj, amount: 0, count: 0, custs: {}, bases: {}, heads: {}, mgrs: {}, orders: [] };
      var g = gjAgg[info.gj];
      g.amount += o.amount;
      g.count++;
      g.custs[o.uid] = 1;
      if (info.base) g.bases[info.base] = 1;
      if (info.head) g.heads[info.head] = 1;
      if (info.mgr) g.mgrs[info.mgr] = 1;
      g.orders.push(o);
    });

    // 4) 汇总行(按金额倒序)
    var rows = Object.keys(gjAgg).map(function (k) {
      var g = gjAgg[k];
      return {
        gj: g.gj,
        base: Object.keys(g.bases).join('/'),
        head: Object.keys(g.heads).join('/'),
        mgr: Object.keys(g.mgrs).join('/'),
        amount: round2(g.amount),
        count: g.count,
        custCount: Object.keys(g.custs).length,
        orders: g.orders
      };
    }).sort(function (a, b) { return b.amount - a.amount; });

    var totals = { amount: 0, count: 0, custCount: Object.keys(custAll).length, gjCount: rows.length };
    rows.forEach(function (r) { totals.amount += r.amount; totals.count += r.count; });
    totals.amount = round2(totals.amount);

    var unmatchedAmt = round2(unmatched.reduce(function (s, o) { return s + o.amount; }, 0));
    return {
      params: p, rows: rows, totals: totals,
      unmatched: unmatched, unmatchedAmt: unmatchedAmt,
      poolTotal: poolObjs.length, timeFiltered: inTime.length, deduped: kept.length, dedupRemoved: dd.removed,
      matched: matched
    };
  }

  // ---------- 上传交互 ----------
  function renderPoolFiles() {
    var box = document.getElementById('poolFiles');
    box.innerHTML = '';
    POOL_FILES.forEach(function (f, i) {
      var d = document.createElement('div');
      d.className = 'f';
      d.innerHTML = '<span>✅ ' + esc(f.name) + '</span><button data-i="' + i + '">✕ 移除</button>';
      box.appendChild(d);
    });
    Array.prototype.forEach.call(box.querySelectorAll('button'), function (b) {
      b.onclick = function () { POOL_FILES.splice(+b.getAttribute('data-i'), 1); renderPoolFiles(); };
    });
    document.getElementById('poolBox').className = 'up-box' + (POOL_FILES.length ? ' ok' : '');
  }

  function renderMapInfo() {
    var box = document.getElementById('mapFiles'), det = document.getElementById('mapDetected');
    box.innerHTML = '';
    if (!MAP) { det.style.display = 'none'; document.getElementById('mapBox').className = 'up-box'; return; }
    box.innerHTML = '<div class="f"><span>✅ ' + esc(MAP.name) + '</span><button id="mapDel">✕ 移除</button></div>';
    document.getElementById('mapDel').onclick = function () { MAP = null; renderMapInfo(); };
    var c = MAP.cols;
    det.style.display = 'block';
    det.innerHTML =
      '识别结果:高经列=<b>' + esc(c.headerRow >= 0 ? (c.header[c.gj] || '第' + c.gj + '列') : '') + '</b>' +
      ' · 用户id列=<b>' + esc(c.uid >= 0 ? (c.header[c.uid] || '第' + c.uid + '列') : '未识别') + '</b><br>' +
      '数据行 ' + MAP.nData + ' · 唯一用户id ' + MAP.uidCount + ' · 高经 ' + MAP.gjCount + ' 个' +
      (MAP.conflict ? ' · <span style="color:var(--red)">同id不同高经冲突 ' + MAP.conflict + '</span>' : ' · 无冲突');
    document.getElementById('mapBox').className = 'up-box ok';
  }

  async function handlePoolFiles(fileList) {
    var arr = Array.prototype.slice.call(fileList);
    for (var i = 0; i < arr.length; i++) {
      var f = arr[i];
      if (!/\.xlsx?$/i.test(f.name)) continue;
      try {
        var sheets = await readFile(f);
        var rows = null;
        sheets.forEach(function (sh) { if (!rows && isOrderPoolRows(sh.rows)) rows = sh.rows; });
        if (!rows) { log('未识别为订单池: ' + f.name + ' — 未找到含"账号ID/订单ID"表头的表', 'err'); continue; }
        POOL_FILES.push({ name: f.name, rows: rows });
        log('已载入有效订单池: ' + f.name + ' (' + rows.filter(function (r, i) { return r && !isHeaderRow(r) && str(r[0]); }).length + ' 行)', 'ok');
      } catch (e) { log('读取失败: ' + f.name + ' — ' + e.message, 'err'); }
    }
    renderPoolFiles();
  }

  async function handleMapFile(file) {
    if (!/\.xlsx?$/i.test(file.name)) return;
    try {
      var sheets = await readFile(file);
      var found = null;
      sheets.forEach(function (sh) {
        if (found) return;
        var cols = detectTeamCols(sh.rows);
        if (cols && cols.uid >= 0) found = { rows: sh.rows, cols: cols };
      });
      if (!found) {
        log('未识别为高经底表: ' + file.name + ' — 需要同时有"高经"列和可识别的用户id列', 'err');
        return;
      }
      MAP = buildTeamMap(found.rows, found.cols, file.name);
      renderMapInfo();
      log('已载入高经底表: ' + file.name + ' — 高经列[' + str(found.rows[found.cols.headerRow][found.cols.gj]) +
        '] 用户id列[' + str(found.rows[found.cols.headerRow][found.cols.uid]) + '],唯一用户id ' + MAP.uidCount +
        ' 个,高经 ' + MAP.gjCount + ' 个' + (MAP.conflict ? ',同id冲突 ' + MAP.conflict : ''), 'ok');
    } catch (e) { log('读取失败: ' + file.name + ' — ' + e.message, 'err'); }
  }

  // ---------- 运行 ----------
  function run() {
    if (!POOL_FILES.length) { alert('请先上传有效订单池 Excel'); return; }
    if (!MAP) { alert('请上传带高经团队的底表(高经业绩.xlsx)'); return; }
    log('开始匹配并数透…');
    var params = {
      period: document.getElementById('period').value.trim(),
      start: document.getElementById('orderStart').value,
      end: document.getElementById('orderEnd').value,
      dedupeMode: document.getElementById('dedupeMode').value
    };
    var pool = collectPoolRows(POOL_FILES);
    lastResult = compute(pool, MAP, params);
    var r = lastResult;
    log('完成:订单池 ' + r.poolTotal + ' 单' +
      (params.start || params.end ? ' → 时间范围内 ' + r.timeFiltered + ' 单' : '') +
      (r.dedupRemoved ? ' → 去重剔除 ' + r.dedupRemoved + ' 行,剩 ' + r.deduped + ' 单' : '') +
      ';匹配到高经 ' + r.matched + ' 单(' + Math.round(r.matched / (r.deduped || 1) * 100) + '%),' +
      '未匹配 ' + r.unmatched.length + ' 单(金额 ' + fmtMoney(r.unmatchedAmt) + ');' +
      '数透 ' + r.rows.length + ' 个高经,订单金额合计 ' + fmtMoney(r.totals.amount),
      r.unmatched.length ? '' : 'ok');
    renderStats(r); renderTable(r, '');
    document.getElementById('searchTeam').value = '';
    document.getElementById('results').style.display = 'block';
  }

  function renderStats(r) {
    var el = document.getElementById('stats');
    function card(t, n, c, click) {
      return '<div class="card' + (click ? ' clickable' : '') + '"' + (click ? ' data-card="unmatched"' : '') +
        '><div class="card-n" style="color:' + c + '">' + n + '</div><div class="card-t">' + t + '</div></div>';
    }
    el.innerHTML =
      card('高经团队数', r.totals.gjCount, 'var(--blue)') +
      card('订单金额合计(元)', fmtMoney(r.totals.amount), 'var(--gold)') +
      card('匹配订单数', r.totals.count, 'var(--green)') +
      card('出单用户数', r.totals.custCount, 'var(--green)') +
      card('未匹配订单', r.unmatched.length, r.unmatched.length ? 'var(--red)' : 'var(--gray)', true);
    el.onclick = function (e) {
      var c = e.target.closest('[data-card]');
      if (c && c.getAttribute('data-card') === 'unmatched') openUnmatched();
    };
  }

  function renderTable(r, kw) {
    var el = document.getElementById('teamTable');
    var kwL = (kw || '').toLowerCase();
    var max = r.rows.length ? r.rows[0].amount : 1;
    var head = '<table class="grid"><tr>' +
      ['排名', '高经团队', '基地', '基地负责人', '经理', '订单金额(元)', '占比', '订单数', '出单用户数', '进度']
        .map(function (x) { return '<th>' + x + '</th>'; }).join('') + '</tr>';
    var body = '', shown = 0;
    r.rows.forEach(function (t, i) {
      if (kwL) {
        var hay = (t.gj + ' ' + t.base + ' ' + t.head + ' ' + t.mgr).toLowerCase();
        if (hay.indexOf(kwL) < 0) return;
      }
      shown++;
      var share = r.totals.amount ? (t.amount / r.totals.amount * 100) : 0;
      var w = max ? Math.min(100, t.amount / max * 100) : 0;
      body += '<tr class="click" data-i="' + i + '">' +
        '<td>' + (i + 1) + '</td>' +
        '<td><b>' + esc(t.gj) + '</b></td>' +
        '<td>' + esc(t.base || '—') + '</td>' +
        '<td>' + esc(t.head || '—') + '</td>' +
        '<td>' + esc(t.mgr || '—') + '</td>' +
        '<td>' + fmtMoney(t.amount) + '</td>' +
        '<td>' + share.toFixed(1) + '%</td>' +
        '<td>' + t.count + '</td>' +
        '<td>' + t.custCount + '</td>' +
        '<td><div class="bar-wrap"><div class="bar-fill" style="width:' + w + '%"></div></div></td>' +
        '</tr>';
    });
    if (!shown) body = '<tr><td colspan="10" class="muted" style="padding:16px">没有匹配的结果</td></tr>';
    el.innerHTML = head + body + '</table>';
    el.onclick = function (e) {
      var tr = e.target.closest('tr.click');
      if (tr) openDetail(+tr.getAttribute('data-i'));
    };
  }

  // ---------- 弹窗 ----------
  function openModal(id) { document.getElementById(id).style.display = 'flex'; }
  function closeModal(id) { document.getElementById(id).style.display = 'none'; }

  function openDetail(i) {
    var r = lastResult;
    if (!r || !r.rows[i]) return;
    var t = r.rows[i];
    document.getElementById('detailModalTitle').textContent =
      t.gj + ' · 金额 ' + fmtMoney(t.amount) + ' / ' + t.count + ' 单 / ' + t.custCount + ' 位用户' +
      (t.base ? ' · ' + t.base : '');
    var rows = '<table class="grid"><tr><th>账号ID(用户id)</th><th>真实姓名</th><th>联系方式</th><th>订单ID</th>' +
      '<th>商品名称</th><th>订单金额(元)</th><th>业绩归属时间</th><th>归属工号</th></tr>';
    t.orders.sort(function (a, b) { return (b.dt ? b.dt.getTime() : 0) - (a.dt ? a.dt.getTime() : 0); })
      .forEach(function (o) {
        rows += '<tr><td style="text-align:left">' + esc(o.uid) + '</td><td>' + esc(o.name) + '</td><td>' + esc(o.phone) + '</td><td>' + esc(o.orderId) + '</td>' +
          '<td style="text-align:left">' + esc(str(o.row[11])) + '</td><td>' + fmtMoney(o.amount) + '</td><td>' + esc(o.perfTime || o.payTime) + '</td><td>' + esc(o.gonghao) + '</td></tr>';
      });
    rows += '</table>';
    document.getElementById('detailModalBody').innerHTML = rows;
    openModal('detailModal');
  }

  function openUnmatched() {
    var r = lastResult;
    if (!r) return;
    var rows = '<table class="grid"><tr><th>账号ID(用户id)</th><th>真实姓名</th><th>订单ID</th><th>商品名称</th><th>订单金额(元)</th><th>业绩归属时间</th><th>归属工号</th><th>来源文件</th></tr>';
    r.unmatched.slice().sort(function (a, b) { return b.amount - a.amount; }).forEach(function (o) {
      rows += '<tr><td style="text-align:left">' + esc(o.uid) + '</td><td>' + esc(o.name) + '</td><td>' + esc(o.orderId) + '</td>' +
        '<td style="text-align:left">' + esc(str(o.row[11])) + '</td><td>' + fmtMoney(o.amount) + '</td><td>' + esc(o.perfTime || o.payTime) + '</td>' +
        '<td>' + esc(o.gonghao) + '</td><td>' + esc(o.src) + '</td></tr>';
    });
    rows += '</table>';
    if (!r.unmatched.length) rows = '<p class="muted">全部订单都匹配到了高经。</p>';
    document.getElementById('unmatchedModalBody').innerHTML = rows;
    openModal('unmatchedModal');
  }

  // ---------- 导出 ----------
  function exportXlsx(r) {
    var wb = XLSX.utils.book_new();
    // 1) 高经团队数透
    var aoa = [['排名', '高经团队', '基地', '基地负责人', '经理', '订单金额(元)', '占比', '订单数', '出单用户数']];
    r.rows.forEach(function (t, i) {
      aoa.push([i + 1, t.gj, t.base, t.head, t.mgr, t.amount,
        (r.totals.amount ? (t.amount / r.totals.amount * 100).toFixed(2) + '%' : ''), t.count, t.custCount]);
    });
    aoa.push(['合计', r.rows.length + ' 个高经', '', '', '', r.totals.amount, '100.00%', r.totals.count, r.totals.custCount]);
    if (r.unmatched.length) aoa.push(['未匹配', '', '', '', '', r.unmatchedAmt, '', r.unmatched.length, '']);
    var ws1 = XLSX.utils.aoa_to_sheet(aoa);
    ws1['!cols'] = [{ wch: 6 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 9 }, { wch: 9 }, { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, ws1, '高经团队数透');
    // 2) 订单明细(原始订单池行 + 匹配出的高经等列)
    var poolHeader = null;
    POOL_FILES.forEach(function (f) {
      if (poolHeader) return;
      f.rows.forEach(function (row) { if (!poolHeader && isHeaderRow(row)) poolHeader = row; });
    });
    var baseCols = poolHeader ? poolHeader.slice() : [];
    var nBase = baseCols.length;
    var detAoa = [baseCols.concat(['高经', '基地', '基地负责人', '经理'])];
    r.rows.forEach(function (t) {
      t.orders.forEach(function (o) {
        var line = o.row.slice(0, nBase);
        while (line.length < nBase) line.push('');
        detAoa.push(line.concat([o.gj || '', o.base || '', o.head || '', o.mgr || '']));
      });
    });
    var ws2 = XLSX.utils.aoa_to_sheet(detAoa);
    XLSX.utils.book_append_sheet(wb, ws2, '订单明细(含高经)');
    // 3) 未匹配订单
    if (r.unmatched.length) {
      var uAoa = [baseCols.concat(['未匹配原因'])];
      r.unmatched.forEach(function (o) {
        var line = o.row.slice(0, nBase);
        while (line.length < nBase) line.push('');
        uAoa.push(line.concat(['底表中无该用户id']));
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(uAoa), '未匹配订单');
    }
    var fname = (r.params.period ? r.params.period + ' ' : '') + '高经团队周进度核算.xlsx';
    XLSX.writeFile(wb, fname);
  }

  // ---------- 初始化 ----------
  function initPage() {
    var pd = document.getElementById('poolDrop'), pi = document.getElementById('poolInput');
    pd.onclick = function () { pi.click(); };
    pd.ondragover = function (e) { e.preventDefault(); pd.classList.add('drag'); };
    pd.ondragleave = function () { pd.classList.remove('drag'); };
    pd.ondrop = function (e) { e.preventDefault(); pd.classList.remove('drag'); handlePoolFiles(e.dataTransfer.files); };
    pi.onchange = function () { handlePoolFiles(pi.files); pi.value = ''; };

    var md = document.getElementById('mapDrop'), mi = document.getElementById('mapInput');
    md.onclick = function () { mi.click(); };
    md.ondragover = function (e) { e.preventDefault(); md.classList.add('drag'); };
    md.ondragleave = function () { md.classList.remove('drag'); };
    md.ondrop = function (e) { e.preventDefault(); md.classList.remove('drag'); if (e.dataTransfer.files[0]) handleMapFile(e.dataTransfer.files[0]); };
    mi.onchange = function () { if (mi.files[0]) handleMapFile(mi.files[0]); mi.value = ''; };

    document.getElementById('btnRun').onclick = run;
    document.getElementById('btnDownload').onclick = function () { if (lastResult) exportXlsx(lastResult); };
    document.getElementById('searchTeam').oninput = function () { if (lastResult) renderTable(lastResult, this.value); };
    document.getElementById('detailModalClose').onclick = function () { closeModal('detailModal'); };
    document.getElementById('unmatchedModalClose').onclick = function () { closeModal('unmatchedModal'); };
    Array.prototype.forEach.call(document.querySelectorAll('.modal-mask'), function (m) {
      m.addEventListener('click', function (e) { if (e.target === m) m.style.display = 'none'; });
    });
    renderPoolFiles();
    renderMapInfo();
  }

  if (document.getElementById('poolDrop')) initPage();

  // 暴露纯函数,供自测与复用
  global.TeamProgress = {
    compute: compute,
    detectTeamCols: detectTeamCols,
    buildTeamMap: buildTeamMap,
    collectPoolRows: collectPoolRows,
    dedupe: dedupe,
    isOrderPoolRows: isOrderPoolRows,
    isHeaderRow: isHeaderRow,
    toFloat: toFloat,
    round2: round2,
    fmtMoney: fmtMoney
  };
})(typeof window !== 'undefined' ? window : this);