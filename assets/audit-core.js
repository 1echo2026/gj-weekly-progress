/*!
 * audit-core.js  中奖名单审核核心逻辑(纯 JS,无第三方依赖)
 * 与 Python 版 scripts/audit_winners.py + filter_orders.py 口径一致:
 *  1) 匹配池: 全部渠道且订单金额 >= 999(渠道不再剔除,渠道差异在资格判定阶段归类异常)
 *     有效订单(判池/工号累计/订单池导出): 仅渠道前缀 grow_xcg_zhuanjs_xiaoshou
 *  2) 头像匹配中奖名单 ↔ 问卷;加密手机号(前3位****后4位)匹配订单
 *  3) 一个手机号只保留提交时间最早的一次抽奖
 *  4) 按工号累计有效出单逐单判池:
 *     - 单笔订单金额 >= 2万:该订单直接归巅峰
 *     - 多手机号累计:跨2万的那笔及之后出单归巅峰,未满2万时归进阶
 *  5) 中奖手机号关联订单归哪个池,本次抽奖就必须在哪个池
 *  6) 渠道异常归类: 匹配到的订单全部非销售直推渠道时,渠道前缀为
 *     grow_xcg_zhuanjs_fudao 的按"非销售直推"处理,其它渠道按"非转介绍"处理
 */
(function (global) {
  'use strict';

  var CN_NUM = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8};
  var NUM_CN = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八'};
  var formSeq = 0;  // 提交记录的稳定编号(进阶问卷在前),跨重新审核保持一致,供前端绿色通道/修改定位

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function toFloat(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isFinite(n) ? n : 0;
  }

  function parseDt(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    if (!s) return null;
    var m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    if (m) {
      return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
    }
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  function fmtDt(d) {
    if (!d) return '';
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
      ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
  }

  function fmtMoney(n) {
    return toFloat(n).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  }

  function cleanPhone(raw) {
    if (raw === null || raw === undefined) return {phone: '', ok: false};
    var m = String(raw).match(/1\d{10}/);
    return m ? {phone: m[0], ok: true} : {phone: String(raw).trim(), ok: false};
  }

  function maskPhone(p) {
    return p.length >= 7 ? p.slice(0, 3) + '****' + p.slice(-4) : p;
  }

  function parsePrizeText(text) {
    if (!text) return {level: null, name: ''};
    var s = String(text).trim();
    // 兼容 "一等奖：冰箱" "一等奖 冰箱" "1等奖·冰箱" "一等奖冰箱" 等格式
    var m = s.match(/^([一二三四五六七八九十\d]+)等奖[：:、\s·]*(.*)$/);
    if (!m) return {level: null, name: s};
    var lv = /^\d+$/.test(m[1]) ? parseInt(m[1], 10) : (CN_NUM[m[1]] || 0);
    var name = (m[2] || '').trim();
    return {level: lv, name: name || s};
  }

  // ---------- 读取结构(输入均为二维数组) ----------
  function loadDistribution(rows) {
    var map = {};
    var curLevel = null;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i] || [];
      var c0 = str(row[0]);
      var m = c0.match(/^([一二三四五六七八九十]+)等奖[：:]/);
      if (m && c0.indexOf('×') >= 0) { curLevel = CN_NUM[m[1]]; continue; }
      if (c0 === '奖项名称') continue;
      if (row[1] && row[2] && str(row[2]).indexOf('http') === 0) {
        map[str(row[2])] = {
          level: curLevel,
          prize_name: c0,
          nickname: str(row[1]),
          redeem: str(row[3])
        };
      }
    }
    return map;
  }

  function loadQuestionnaire(rows, pool) {
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i] || [];
      if (!row[6]) continue;
      var pp = parsePrizeText(row[2]);
      var cp = cleanPhone(row[6]);
      var dt = parseDt(row[3]);
      out.push({
        pool: pool,
        nickname: str(row[0]),
        avatar: str(row[1]),
        prize_level: pp.level,
        prize_name: pp.name,
        submit_time_raw: dt ? fmtDt(dt) : str(row[3]),
        submit_dt: dt,
        base: str(row[4]),
        real_name: str(row[5]),
        phone: cp.phone,
        phone_format_ok: cp.ok,
        form_i: formSeq++,
        raw_row: row
      });
    }
    return out;
  }

  // ---------- 异常类型分类 ----------
  // 依据无效原因/重复/备注,归类为可套用话术模板的异常类型
  function classifyAbnormal(s) {
    if (s.is_duplicate) return '重复抽奖';
    var r = s.invalid_reasons || '';
    if (r.indexOf('问卷提交时间') === 0 && r.indexOf('不在本期活动时间范围内') > 0) return '提交时间超范围';
    if (r.indexOf('出单时间不符合当前抽奖日期') >= 0) return '出单时间不符合当前抽奖日期';
    if (r.indexOf('手机号格式不正确') >= 0) return '手机号格式错误';
    if (r.indexOf('手机号在订单池中查找不到') >= 0) return '手机号未查到';
    if (r.indexOf('非转介绍') >= 0) return '非转介绍';
    if (r.indexOf('非销售直推') >= 0) return '非销售直推';
    if (r.indexOf('未在有效订单池查找到') >= 0) return '订单未达有效条件';
    if (r.indexOf('问卷头像与中奖名单无法匹配') >= 0) return '头像无法匹配';
    if (r.indexOf('奖池抽奖') >= 0 && (r.indexOf('应属') >= 0 || r.indexOf('应直接') >= 0 || r.indexOf('应在') >= 0)) {
      return s.expected_pool === '巅峰' ? '抽错奖池（应属巅峰）' : '抽错奖池（应属进阶）';
    }
    if (s.abnormal_status) return '含退款/换课订单';
    return '其他异常';
  }

  // ---------- 主流程 ----------
  function audit(input) {
    var params = input.params || {};
    formSeq = 0;
    var PREFIX = params.channelPrefix || 'grow_xcg_zhuanjs_xiaoshou';
    var MIN_AMT = toFloat(params.minAmount != null ? params.minAmount : 999);
    var THRESHOLD = toFloat(params.threshold != null ? params.threshold : 20000);
    // 可配置的时间范围(为空表示不限制)
    var PERIOD = str(params.period);
    var QISHU = str(params.qishu);
    var FORM_START = parseDt(params.formStart), FORM_END = parseDt(params.formEnd);      // 问卷提交时间范围
    var ORDER_START = parseDt(params.orderStart), ORDER_END = parseDt(params.orderEnd);  // 订单支付时间范围

    function inRange(dt, start, end) {
      if (!dt) return true;
      if (start && dt < start) return false;
      if (end && dt > end) return false;
      return true;
    }

    // 1) 过滤有效订单(含支付时间范围)
    function filterPool(rows) {
      var title = rows[0] || [];
      var header = rows[1] || [];
      // 动态查找表头行: 强基文件行0即表头(账号ID/真实姓名...),升学文件行1为表头
      var start = 2;
      for (var k = 0; k < Math.min(rows.length, 3); k++) {
        var rr = rows[k] || [];
        if (str(rr[0]) === '账号ID' || str(rr[1]) === '真实姓名') {
          header = rr;
          title = k > 0 ? (rows[k - 1] || []) : [];   // 表头在第0行时没有标题行,留空避免导出重复表头
          start = k + 1;
          break;
        }
      }
      // 匹配池保留全部渠道(渠道差异在资格判定阶段归类为异常)
      var kept = [], rejAmt = 0, rejTime = 0;
      for (var i = start; i < rows.length; i++) {
        var r = rows[i];
        if (!r || !r[0]) continue;
        var amt = toFloat(r[4]);
        if (amt < MIN_AMT) { rejAmt++; continue; }
        if (!inRange(parseDt(r[5]), ORDER_START, ORDER_END)) { rejTime++; continue; }
        kept.push(r);
      }
      return {title: title, header: header, kept: kept, rejAmount: rejAmt, rejTime: rejTime, total: kept.length + rejAmt + rejTime};
    }

    var fQj = filterPool(input.rawOrders.qiangji || []);
    var fSx = filterPool(input.rawOrders.shengxue || []);

    function toRec(r, source) {
      return {
        source: source,
        account: str(r[0]),
        cust_name: str(r[1]),
        masked: str(r[2]),
        order_id: str(r[3]),
        amount: toFloat(r[4]),
        pay_time: str(r[5]),
        dt: parseDt(r[5]),
        perf_time: str(r[7]),
        perf_dt: parseDt(r[7]),
        status: str(r[6]),
        product_id: str(r[10]),
        sales_name: str(r[24]),
        gonghao: str(r[25]),
        channel: str(r[20])
      };
    }

    // 匹配池: 全部渠道(≥999),用于匹配中奖名单;渠道差异在资格判定阶段归类为异常
    var allRows = [];
    fQj.kept.forEach(function (r) { allRows.push(toRec(r, '强基')); });
    fSx.kept.forEach(function (r) { allRows.push(toRec(r, '升学')); });
    // 有效订单: 仅销售直推渠道(grow_xcg_zhuanjs_xiaoshou),用于判池与工号累计
    function isXiaoshou(r) { return r.channel.indexOf(PREFIX) === 0; }
    var validRows = allRows.filter(isXiaoshou);

    // (订单ID,商品ID) 去重,防两池重复导出
    var seenKey = {}, deduped = [];
    allRows.forEach(function (r) {
      var k = r.order_id + '|' + r.product_id;
      if (seenKey[k]) return;
      seenKey[k] = 1;
      deduped.push(r);
    });
    allRows = deduped;
    validRows = allRows.filter(isXiaoshou);

    // 2) 中奖名单 + 问卷
    var winners = {
      '进阶': loadDistribution(input.distribution['进阶'] || []),
      '巅峰': loadDistribution(input.distribution['巅峰'] || [])
    };
    var submissions = []
      .concat(loadQuestionnaire(input.questionnaire['进阶'] || [], '进阶'))
      .concat(loadQuestionnaire(input.questionnaire['巅峰'] || [], '巅峰'));

    // 3) 工号时间线 + 逐单判池
    var ghRowCnt = {}, ghNameVotes = {}, orderAgg = {};
    function inc(obj, k, n) { obj[k] = (obj[k] || 0) + (n || 1); }
    validRows.forEach(function (r) {
      if (!r.gonghao) return;
      inc(ghRowCnt, r.gonghao);
      ghNameVotes[r.gonghao] = ghNameVotes[r.gonghao] || {};
      inc(ghNameVotes[r.gonghao], r.sales_name);
      if (orderAgg[r.order_id]) {
        orderAgg[r.order_id].amount += r.amount;
      } else {
        orderAgg[r.order_id] = {
          amount: r.amount, dt: r.dt, masked: r.masked, gonghao: r.gonghao,
          pool: '进阶', cum_after: 0, direct_peak: false
        };
      }
    });

    var ghOrders = {};
    Object.keys(orderAgg).forEach(function (oid) {
      var gh = orderAgg[oid].gonghao;
      (ghOrders[gh] = ghOrders[gh] || {})[oid] = [orderAgg[oid].dt, orderAgg[oid].amount];
    });

    var ghStats = {};
    Object.keys(ghOrders).forEach(function (gh) {
      var entries = Object.keys(ghOrders[gh]).map(function (oid) {
        return [oid, ghOrders[gh][oid][0], ghOrders[gh][oid][1]];
      });
      entries.sort(function (a, b) { return (a[1] ? a[1].getTime() : 8.64e15) - (b[1] ? b[1].getTime() : 8.64e15); });
      var total = 0, maxSingle = 0, cum = 0, peakSince = null, peakOrders = 0;
      entries.forEach(function (e) {
        var oid = e[0], dt = e[1], amt = e[2];
        total += amt;
        if (amt > maxSingle) maxSingle = amt;
        cum += amt;
        var isPeak = amt >= THRESHOLD || cum >= THRESHOLD;
        var o = orderAgg[oid];
        o.pool = isPeak ? '巅峰' : '进阶';
        o.cum_after = Math.round(cum * 100) / 100;
        o.direct_peak = amt >= THRESHOLD;
        if (isPeak && !peakSince) peakSince = dt;
        if (isPeak) peakOrders++;
      });
      var votes = ghNameVotes[gh], bestName = '', bestN = -1;
      Object.keys(votes).forEach(function (n) { if (votes[n] > bestN) { bestN = votes[n]; bestName = n; } });
      ghStats[gh] = {
        sales_name: bestName,
        total: Math.round(total * 100) / 100,
        max_single: Math.round(maxSingle * 100) / 100,
        row_count: ghRowCnt[gh] || 0,
        peak_since: peakSince,
        peak_orders: peakOrders
      };
    });

    // 4) 掩码 -> 匹配池订单行(全部渠道,≥999)
    var validIndex = {};
    allRows.forEach(function (r) {
      if (!r.masked) return;
      (validIndex[r.masked] = validIndex[r.masked] || []).push(r);
    });

    // 5) 匹配 + 期望奖池
    submissions.forEach(function (s) {
      // 未提供"发放情况"(中奖名单)表时,跳过头像匹配校验(问卷表单已含基地/姓名/出单手机号,不再依赖发放表)
      var win = winners[s.pool][s.avatar];
      s.avatar_matched = Object.keys(winners[s.pool]).length ? !!win : true;
      s.win_nickname = win ? win.nickname : '';
      s.win_redeem = win ? win.redeem : '';
      s.match = null;
      s.total_amount = 0; s.order_count = 0; s.max_single_order = 0;
      s.peak_since = null; s.expected_pool = ''; s.peak_reason = '';
      s.account = ''; s.sales_name = ''; s.gonghao = ''; s.abnormal_status = ''; s.channel = '';
      if (!s.phone_format_ok) return;
      var rows = validIndex[maskPhone(s.phone)] || [];
      var acctSet = {};
      rows.forEach(function (r) { acctSet[r.account] = 1; });
      if (!rows.length || Object.keys(acctSet).length > 1) return;
      s.match = 'ok';
      // 渠道归类: 匹配到的订单全部非销售直推渠道时,记为渠道异常(非销售直推/非转介绍)
      var xsRows = rows.filter(isXiaoshou);
      if (!xsRows.length) {
        s.channel_bad = {channel: rows[rows.length - 1].channel};
        s.channel = s.channel_bad.channel;
        var tmax = null;
        rows.forEach(function (r) { if (r.dt && (!tmax || r.dt > tmax)) tmax = r.dt; });
        s.order_time = tmax ? fmtDt(tmax) : '';
        return;
      }
      rows = xsRows;
      s.account = rows[0].account;
      // 业绩归属渠道(匹配到的有效订单渠道,去重合并)
      var chSet = {};
      rows.forEach(function (r) { if (r.channel) chSet[r.channel] = 1; });
      s.channel = Object.keys(chSet).join('、');
      var nv = {}, gv = {};
      rows.forEach(function (r) { inc(nv, r.sales_name); inc(gv, r.gonghao); });
      s.sales_name = Object.keys(nv).sort(function (a, b) { return nv[b] - nv[a]; })[0] || '';
      s.gonghao = Object.keys(gv).sort(function (a, b) { return gv[b] - gv[a]; })[0] || '';
      var st = ghStats[s.gonghao];
      if (st) {
        s.total_amount = st.total;
        s.order_count = st.row_count;
        s.max_single_order = st.max_single;
        s.peak_since = st.peak_since;
        s.peak_orders = st.peak_orders || 0;
      }
      s.order_time = '';
      rows.forEach(function (r) {
        // 业绩归属时间取订单池"业绩归属时间"列(c7);无该列时回退支付时间(c5)
        var t = r.perf_dt || r.dt;
        if (t && (!s.order_time || t > s.order_time)) s.order_time = t;
      });
      s.order_time = s.order_time ? fmtDt(s.order_time) : '';
      var oidMap = {}, relInfos = [];
      rows.forEach(function (r) { if (r.order_id) oidMap[r.order_id] = 1; });
      Object.keys(oidMap).forEach(function (oid) { if (orderAgg[oid]) relInfos.push(orderAgg[oid]); });
      if (relInfos.length) {
        var hasPeak = relInfos.some(function (o) { return o.pool === '巅峰'; });
        s.expected_pool = hasPeak ? '巅峰' : '进阶';
        if (hasPeak) {
          var direct = relInfos.filter(function (o) { return o.direct_peak; });
          if (direct.length) {
            var o = direct.sort(function (a, b) { return b.amount - a.amount; })[0];
            s.peak_reason = '该出单手机号存在单笔订单金额' + fmtMoney(o.amount) + '元≥2万元' +
              '（支付于' + fmtDt(o.dt) + '），应直接在巅峰奖池抽奖';
          } else {
            var o2 = relInfos.filter(function (o) { return o.pool === '巅峰'; })[0];
            s.peak_reason = '该手机号订单支付于' + fmtDt(o2.dt) + '时，工号累计有效出单金额已达' +
              fmtMoney(o2.cum_after) + '元（≥2万），累计达2万之后的出单应在巅峰奖池抽奖';
          }
        } else {
          var last = relInfos.sort(function (a, b) { return b.cum_after - a.cum_after; })[0];
          s.peak_reason = '该手机号订单支付于' + fmtDt(last.dt) + '时，工号累计有效出单金额仅' +
            fmtMoney(last.cum_after) + '元（<2万），应在进阶奖池抽奖';
        }
      }
      var ab = {};
      rows.forEach(function (r) { if (r.status === '完全退款' || r.status === '换课原订单') ab[r.status] = 1; });
      s.abnormal_status = Object.keys(ab).join('、');
    });

    // 6) 一手机号一单去重(保留提交时间最早的一条)
    var groups = {};
    submissions.forEach(function (s, i) { if (s.phone_format_ok) (groups[s.phone] = groups[s.phone] || []).push(i); });
    var duplicates = [];
    Object.keys(groups).forEach(function (phone) {
      var idxs = groups[phone];
      if (idxs.length <= 1) return;
      idxs.sort(function (a, b) {
        var da = submissions[a].submit_dt, db = submissions[b].submit_dt;
        return (da ? da.getTime() : 8.64e15) - (db ? db.getTime() : 8.64e15);
      });
      // 同一手机号在最早提交时刻有多条记录时(例如进阶全量表单与巅峰表单包含
      // 完全相同的提交),优先保留巅峰奖池那条 —— 巅峰表单是该奖池的精确来源,
      // 避免巅峰提交被进阶副本按数组顺序误判为重复而整池丢失。
      var earliest = idxs[0];
      var firstT = submissions[earliest].submit_dt ? submissions[earliest].submit_dt.getTime() : 8.64e15;
      idxs.forEach(function (i) {
        var t = submissions[i].submit_dt ? submissions[i].submit_dt.getTime() : 8.64e15;
        if (t === firstT && submissions[i].pool === '巅峰' && submissions[earliest].pool !== '巅峰') earliest = i;
      });
      idxs.forEach(function (i) {
        if (i === earliest) return;
        submissions[i].is_duplicate = true;
        duplicates.push(submissions[i]);
      });
    });
    submissions.forEach(function (s) { s.is_duplicate = !!s.is_duplicate; });

    // 7) 资格判定
    var validList = [], invalidList = [];
    submissions.forEach(function (s) {
      var reasons = [], remark = '';
      if (!inRange(s.submit_dt, FORM_START, FORM_END)) {
        reasons.push('问卷提交时间' + (s.submit_time_raw || '未知') + '不在本期活动时间范围内');
      }
      if (!s.phone_format_ok) {
        reasons.push('出单手机号格式不正确，在订单池中查找不到，疑似输入错误');
      } else if (s.channel_bad) {
        var cbCh = s.channel_bad.channel;
        if (cbCh.indexOf('grow_xcg_zhuanjs_fudao') === 0) {
          reasons.push('未在有效订单池查找到（订单归属渠道非销售直推渠道：' + cbCh + '，属于辅导leads接单）');
        } else {
          reasons.push('未在有效订单池查找到（订单归属渠道非转介绍渠道：' + cbCh + '，不属于转介绍出单）');
        }
      } else if (!s.match) {
        var mp = maskPhone(s.phone);
        var rawRows = [];
        (input.rawOrders.qiangji || []).concat(input.rawOrders.shengxue || []).forEach(function (r) {
          if (r && str(r[2]) === mp) rawRows.push(r);
        });
        if (rawRows.length) {
          // 业绩归属时间(订单支付时间)是否在本期活动时间范围内;未配置时间范围时不限制
          var inWin = rawRows.some(function (r) { return inRange(parseDt(r[5]), ORDER_START, ORDER_END); });
          if (!inWin && (ORDER_START || ORDER_END)) {
            var t0 = rawRows.map(function (r) { return parseDt(r[5]); })
                            .filter(function (d) { return !!d; }).sort(function (a, b) { return a - b; })[0];
            s.order_time = t0 ? fmtDt(t0) : '';
            reasons.push('出单时间不符合当前抽奖日期（订单池匹配到的业绩归属时间' +
              (s.order_time || str(rawRows[0][5]) || '未知') + '不在本期活动时间范围内）');
          } else {
            // 手机号在原始订单池能找到但未进入匹配池(未达≥999):要么金额不足,要么匹配到多个账号
            var hasQualified = rawRows.some(function (r) {
              return toFloat(r[4]) >= MIN_AMT && inRange(parseDt(r[5]), ORDER_START, ORDER_END);
            });
            if (hasQualified) {
              reasons.push('未在有效订单池查找到（手机号在订单池匹配到多个账号的达标订单，无法唯一归属）');
            } else {
              reasons.push('未在有效订单池查找到（订单金额不足' + Math.round(MIN_AMT) + '元，未达活动有效条件）');
            }
          }
        } else {
          reasons.push('手机号在订单池中查找不到，疑似输入错误');
        }
      }
      if (!s.avatar_matched) reasons.push('问卷头像与中奖名单无法匹配');
      if (s.match && !s.channel_bad && s.expected_pool && s.pool !== s.expected_pool) {
        reasons.push(s.peak_reason + '，但本次在' + s.pool + '奖池抽奖');
      }
      if (s.abnormal_status) remark += '含' + s.abnormal_status + '订单，请核查';
      s.invalid_reasons = reasons.join('；');
      s.remark = remark;
      s.abnormal_type = classifyAbnormal(s);
      if (s.is_duplicate) return;
      (reasons.length ? invalidList : validList).push(s);
    });

    // 8) 姓名↔工号对照
    var nameGh = {}, ghPhones = {};
    validRows.forEach(function (r) {
      if (r.sales_name && r.gonghao) (nameGh[r.sales_name] = nameGh[r.sales_name] || {})[r.gonghao] = 1;
    });
    validList.concat(invalidList).forEach(function (s) {
      if (s.match === 'ok' && s.gonghao) (ghPhones[s.gonghao] = ghPhones[s.gonghao] || {})[s.phone] = 1;
    });
    var dupNames = {};
    Object.keys(nameGh).forEach(function (n) { if (Object.keys(nameGh[n]).length > 1) dupNames[n] = 1; });

    var nameGhRows = [];
    Object.keys(nameGh).sort().forEach(function (name) {
      Object.keys(nameGh[name]).sort().forEach(function (gh) {
        var st = ghStats[gh] || {};
        nameGhRows.push({
          name: name,
          gonghao: gh,
          phones: Object.keys(ghPhones[gh] || {}).sort(),
          gh_count: Object.keys(nameGh[name]).length,
          total: st.total || 0,
          peak_since: st.peak_since || null,
          is_dup: !!dupNames[name]
        });
      });
    });

    // 9) 图片格式分组(仅有效、去重后;保持问卷顺序)
    function tierGroups(pool, maxLevel) {
      var buckets = {};
      for (var lv = 1; lv <= maxLevel; lv++) buckets[lv] = [];
      var ungraded = [];
      validList.forEach(function (s) {
        if (s.pool !== pool) return;
        if (s.prize_level && buckets[s.prize_level]) buckets[s.prize_level].push(s);
        else ungraded.push(s);
      });
      function toRow(s) {
        var b = s.base;
        if (b.lastIndexOf('基地') === b.length - 2) b = b.slice(0, -2);
        return {base: b, name: s.real_name, gonghao: s.gonghao, phone: s.phone,
                nickname: s.nickname, redeem: s.win_redeem,
                order_time: s.order_time || '', total_amount: s.total_amount || 0};
      }
      var out = [];
      for (var l = 1; l <= maxLevel; l++) {
        if (!buckets[l] || !buckets[l].length) continue;
        out.push({
          level: l,
          level_text: NUM_CN[l] + '等奖',
          prize: buckets[l][0].prize_name,
          rows: buckets[l].map(toRow)
        });
      }
      // 问卷奖品文本无法解析出等级的行:不丢弃,按奖品名称聚合展示
      if (ungraded.length) {
        var byName = {}, order = [];
        ungraded.forEach(function (s) {
          var k = s.prize_name || '未填写奖品';
          if (!byName[k]) { byName[k] = []; order.push(k); }
          byName[k].push(s);
        });
        order.forEach(function (k) {
          out.push({level: null, level_text: '奖品', prize: k, rows: byName[k].map(toRow)});
        });
      }
      return out;
    }

    // 10) 异常沟通名单(无效+重复+含异常订单备注;供套用话术模板)
    var communication = [];
    function commBase(s) {
      return {
        pool: s.pool, nickname: s.nickname, real_name: s.real_name, base: s.base,
        phone: s.phone, prize_name: s.prize_name, prize_level: s.prize_level,
        submit_time: s.submit_time_raw, abnormal_type: s.abnormal_type,
        reason: s.invalid_reasons || '', remark: s.remark, is_duplicate: !!s.is_duplicate,
        expected_pool: s.expected_pool, period: PERIOD, qishu: QISHU,
        order_time: s.order_time || '', peak_orders: s.peak_orders || 0,
        form_i: s.form_i
      };
    }
    invalidList.forEach(function (s) { communication.push(commBase(s)); });
    duplicates.forEach(function (s) { communication.push(commBase(s)); });
    validList.forEach(function (s) { if (s.abnormal_status) communication.push(commBase(s)); });

    // 有效订单池导出仅保留销售直推渠道行
    function xiaoKept(rows) { return rows.filter(function (r) { return str(r[20]).indexOf(PREFIX) === 0; }); }
    var keptQj = xiaoKept(fQj.kept), keptSx = xiaoKept(fSx.kept);

    return {
      params: {channelPrefix: PREFIX, minAmount: MIN_AMT, threshold: THRESHOLD,
               period: PERIOD, qishu: QISHU, formStart: FORM_START, formEnd: FORM_END,
               orderStart: ORDER_START, orderEnd: ORDER_END},
      validPools: {
        '强基': {title: fQj.title, header: fQj.header, kept: keptQj,
                 stats: {total: fQj.total, rejAmount: fQj.rejAmount, rejTime: fQj.rejTime, kept: keptQj.length}},
        '升学': {title: fSx.title, header: fSx.header, kept: keptSx,
                 stats: {total: fSx.total, rejAmount: fSx.rejAmount, rejTime: fSx.rejTime, kept: keptSx.length}}
      },
      stats: {
        submitted: submissions.length,
        valid: validList.length,
        validJJ: validList.filter(function (s) { return s.pool === '进阶'; }).length,
        validDF: validList.filter(function (s) { return s.pool === '巅峰'; }).length,
        invalid: invalidList.length,
        duplicates: duplicates.length,
        ghCount: Object.keys(ghStats).length,
        ghPeak: Object.keys(ghStats).filter(function (g) { return ghStats[g].peak_since; }).length
      },
      validList: validList,
      invalidList: invalidList,
      duplicates: duplicates,
      communication: communication,
      nameGhRows: nameGhRows,
      ghStats: ghStats,
      groups: {'巅峰': tierGroups('巅峰', 4), '进阶': tierGroups('进阶', 8)}
    };
  }

  var api = {
    audit: audit,
    parseDt: parseDt,
    fmtDt: fmtDt,
    fmtMoney: fmtMoney,
    cleanPhone: cleanPhone,
    maskPhone: maskPhone,
    parsePrizeText: parsePrizeText
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.AuditCore = api;
})(typeof window !== 'undefined' ? window : this);
