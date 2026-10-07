// ============================================================
//  SGL 日報 自動送信スクリプト
//  毎日 18:00 JST に自動実行
//  Google スプレッドシートの Apps Script に貼り付けて使用
// ============================================================

const SUPABASE_URL = 'https://pqoexlikdviqqgdnzuud.supabase.co';
const SUPABASE_KEY = 'sb_publishable_9KR8h4be2Q5-4ux6g55hSg_Jkv9UYBM';
const EMAIL_TO    = 'data_sgl@ecodesic.co.jp';
const SHEET_NAME  = '日報';

// ────────────────────────────────────────────────────────────
// メイン（毎日18時に実行）
// ────────────────────────────────────────────────────────────
function dailyReport() {
  const today       = getJSTDateStr(new Date());
  const sow34date   = offsetDate(today, -34);

  const mc  = parseNote(fetchOne('morning_check', today));
  const ec  = parseNote(fetchOne('evening_check', today));
  const hc33 = parseNote(fetchOne('harvest_count', today + '_33'));
  const hc34 = parseNote(fetchOne('harvest_count', sow34date + '_harvest'));
  const liquidTotal = fetchLiquidAmount();

  const ekihiLabel = getEkihiLabel(today);
  const memos      = fetchMemos();
  const alerts     = fetchAlerts();
  const sheetUrl   = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const tana1Url   = 'https://raw.githubusercontent.com/sotetsugreenlab-sys/sumashaqi/main/screenshots/tana-1.png';
  const tana2Url   = 'https://raw.githubusercontent.com/sotetsugreenlab-sys/sumashaqi/main/screenshots/tana-2.png';
  appendToSheet(today, mc, ec, hc33, hc34, liquidTotal, ekihiLabel);
  sendReportEmail(today, mc, ec, hc33, hc34, liquidTotal, ekihiLabel, memos, alerts, sheetUrl, tana1Url, tana2Url);
}

// ────────────────────────────────────────────────────────────
// スプレッドシート書き込み
// ────────────────────────────────────────────────────────────
function fixSheetHeader() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('シート「' + SHEET_NAME + '」が見つかりません'); return; }
  const row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = row.indexOf('貯液(L)');
  if (col === -1) { Logger.log('更新不要（すでに修正済みか、列が見つかりません）'); return; }
  sheet.getRange(1, col + 1).setValue('貯液残量(L)');
  Logger.log('ヘッダーを「貯液残量(L)」に更新しました（列 ' + (col + 1) + '）');
}

function appendToSheet(today, mc, ec, hc33, hc34, liquidTotal, ekihiLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '日付',
      '朝:AC温度', '朝:AC風速',
      '帰:AC温度', '帰:AC風速',
      '前借播種日', '前借数(33日)',
      'メイン播種日', '最小重量(g)', '平均重量(g)', '70g以上', '75g以上', '翌袋', 'シェラトン数', 'シェラトン重量(g)', 'チップバーン',
      '貯液残量(L)', '液肥'
    ];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold').setBackground('#3D7A50').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const sow33 = offsetDate(today, -33);
  const sow34 = offsetDate(today, -34);
  sheet.appendRow([
    today,
    c(mc.ac_temp) + '℃', acMode(mc, 'ac'),
    c(ec.ec_ac_temp) + '℃', acMode(ec, 'ec_ac'),
    sow33, c(hc33.maegari),
    sow34,
    c(hc34.min_weight), c(hc34.avg_weight),
    c(hc34.w70), c(hc34.w75), c(hc34.yoku_bukuro),
    c(hc34.sheraton_count), c(hc34.sheraton_weight), c(hc34.chipburn),
    liquidTotal,
    ekihiLabel || '-'
  ]);
}

// ────────────────────────────────────────────────────────────
// メール送信
// ────────────────────────────────────────────────────────────
function sendReportEmail(today, mc, ec, hc33, hc34, liquidTotal, ekihiLabel, memos, alerts, sheetUrl, tana1Url, tana2Url) {
  const sow33 = offsetDate(today, -33);
  const sow34 = offsetDate(today, -34);
  const row = (label, val) =>
    `<tr>
      <td style="padding:7px 14px;color:#555;border-bottom:1px solid #f0f0f0;white-space:nowrap;">${label}</td>
      <td style="padding:7px 14px;font-weight:600;border-bottom:1px solid #f0f0f0;">${val}</td>
    </tr>`;

  const section = (color, emoji, title, rows) =>
    `<div style="margin-bottom:20px;">
      <div style="font-size:14px;font-weight:700;color:${color};margin-bottom:6px;">${emoji} ${title}</div>
      <table style="width:100%;border-collapse:collapse;font-size:14px;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #e8e8e8;">
        ${rows}
      </table>
    </div>`;

  const html = `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"></head><body>
<div style="font-family:-apple-system,BlinkMacSystemFont,sans-serif;max-width:560px;margin:0 auto;">
  <div style="background:#3D7A50;color:#fff;padding:18px 20px;border-radius:12px 12px 0 0;">
    <div style="font-size:12px;opacity:0.75;letter-spacing:0.05em;">SGL スタッフアプリ</div>
    <div style="font-size:22px;font-weight:700;margin-top:4px;">日報 ${today}</div>
  </div>

  <div style="background:#f7f7f5;padding:20px;">

    ${section('#3D7A50', '&#127749;', '朝チェック', [
      row('エアコン温度', c(mc.ac_temp) + ' ℃'),
      row('エアコン風速', acMode(mc, 'ac'))
    ].join(''))}

    ${section('#185FA5', '&#127750;', '帰り際チェック', [
      row('エアコン温度', c(ec.ec_ac_temp) + ' ℃'),
      row('エアコン風速', acMode(ec, 'ec_ac'))
    ].join(''))}

    ${section('#8B5A00', '&#9986;', `前借数（33日目刈取 / 播種 ${sow33}）`, [
      row('前借数', c(hc33.maegari))
    ].join(''))}

    ${section('#8B5A00', '&#128230;', `メイン刈取（34日目刈取 / 播種 ${sow34}）`, [
      row('最小重量', c(hc34.min_weight) + ' g'),
      row('平均重量', c(hc34.avg_weight) + ' g'),
      row('70g 以上', c(hc34.w70)),
      row('75g 以上', c(hc34.w75)),
      row('翌袋', c(hc34.yoku_bukuro)),
      row('シェラトン 数', c(hc34.sheraton_count)),
      row('シェラトン 重量', c(hc34.sheraton_weight) + ' g'),
      row('チップバーン', c(hc34.chipburn))
    ].join(''))}

    ${section('#185FA5', '&#128167;', '貯液', [
      row('貯液量（現在残量）', liquidTotal + ' L')
    ].join(''))}

    <div style="margin-bottom:20px;">
      <div style="font-size:14px;font-weight:700;color:#6B7280;margin-bottom:6px;">&#129692; 棚管理</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <img src="${tana1Url}" style="max-width:100%;border-radius:8px;border:1px solid #e8e8e8;" alt="棚1">
        <img src="${tana2Url}" style="max-width:100%;border-radius:8px;border:1px solid #e8e8e8;" alt="棚2">
      </div>
    </div>

    ${buildNotesSection(ekihiLabel, memos, alerts, today)}

  </div>
  <div style="background:#e8e8e8;padding:10px 20px;border-radius:0 0 12px 12px;font-size:12px;color:#555;text-align:center;">
    <a href="${sheetUrl}" style="color:#3D7A50;font-weight:700;text-decoration:none;">&#128202; スプレッドシートで確認</a>
    <span style="color:#bbb;margin:0 8px;">|</span>
    自動送信 ${today} 18:00 JST — SGL スタッフアプリ
  </div>
</div></body></html>`;

  let plainText =
    `【SGL日報】${today}\n\n` +
    `朝 AC: ${c(mc.ac_temp)}℃ ${acMode(mc,'ac')}\n` +
    `帰 AC: ${c(ec.ec_ac_temp)}℃ ${acMode(ec,'ec_ac')}\n` +
    `前借り: ${c(hc33.maegari)}\n` +
    `平均重量: ${c(hc34.avg_weight)}g\n` +
    `貯液: ${liquidTotal}L\n`;
  if (ekihiLabel) {
    const next = getNextEkihiEvent(today);
    plainText += `液肥: ${ekihiLabel}`;
    if (next) plainText += `  (次回${next.label}: ${next.date})`;
    plainText += '\n';
  }
  const alertDatesSet = new Set(alerts.map(a => a.sow_date));
  const activeDates = [...new Set([...memos.map(m => m.sow_date), ...alerts.map(a => a.sow_date)])]
    .sort().filter(d => offsetDate(d, 34) >= today);
  if (activeDates.length) {
    plainText += `\n--- メモ・アラート一覧 ---\n`;
    activeDates.forEach(date => {
      const isAlert = alertDatesSet.has(date);
      const memo = memos.find(m => m.sow_date === date);
      plainText += `${isAlert ? '⚠️' : '📝'} 播種:${date} 刈取:${offsetDate(date,34)}${memo ? '  '+memo.note : ''}\n`;
    });
  }

  GmailApp.sendEmail(EMAIL_TO, `【SGL日報】${today}`, plainText, { htmlBody: html });
}

// ────────────────────────────────────────────────────────────
// Supabase ヘルパー
// ────────────────────────────────────────────────────────────
function fetchOne(crop, sowDate) {
  const url = SUPABASE_URL + '/rest/v1/lot_notes'
    + '?crop=eq.' + encodeURIComponent(crop)
    + '&sow_date=eq.' + encodeURIComponent(sowDate)
    + '&select=note&limit=1';
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return null;
    const data = JSON.parse(res.getContentText());
    return data.length > 0 ? data[0].note : null;
  } catch(e) { return null; }
}

function fetchLiquidAmount() {
  const url = SUPABASE_URL + '/rest/v1/liquid_log?select=delta';
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return 0;
    const data = JSON.parse(res.getContentText());
    const total = data.reduce((s, r) => s + (r.delta || 0), 0);
    return Math.round(total * 10) / 10;
  } catch(e) { return 0; }
}

// ────────────────────────────────────────────────────────────
// ユーティリティ
// ────────────────────────────────────────────────────────────
function parseNote(raw) {
  if (!raw) return {};
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

function getJSTDateStr(date) {
  const jst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  return jst.toISOString().slice(0, 10);
}

function offsetDate(dateStr, days) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function acMode(state, prefix) {
  const modes = [['_kyu','急'], ['_tsu','強'], ['_chu','中'], ['_jaku','弱']];
  const found = modes.filter(([sfx]) => state[prefix + sfx]).map(([,l]) => l);
  return found.length ? found.join('/') : '未設定';
}

function c(v) { return (v != null && v !== '') ? v : '-'; }
function b(v) { return v ? '✓' : '－'; }

// ────────────────────────────────────────────────────────────
// メモ・アラート・液肥 ヘルパー
// ────────────────────────────────────────────────────────────
function fetchMemos() {
  const url = SUPABASE_URL + '/rest/v1/lot_notes'
    + '?crop=eq.sumashaqi'
    + '&select=sow_date,note'
    + '&order=sow_date.asc';
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return [];
    const data = JSON.parse(res.getContentText());
    return data.filter(r => r.note && r.note.trim() !== '');
  } catch(e) { return []; }
}

function fetchAlerts() {
  const url = SUPABASE_URL + '/rest/v1/lot_notes'
    + '?crop=eq.sumashaqi_alert'
    + '&select=sow_date'
    + '&order=sow_date.asc';
  try {
    const res = UrlFetchApp.fetch(url, {
      headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return [];
    return JSON.parse(res.getContentText());
  } catch(e) { return []; }
}

const EKIHI_BASE = '2026-03-02';
const EKIHI_EVENTS = [{ label: '入替', offset: 0 }, { label: '追肥', offset: 4 }, { label: '追肥', offset: 9 }];

function getEkihiLabel(today) {
  const base = new Date(EKIHI_BASE + 'T00:00:00Z');
  const d    = new Date(today + 'T00:00:00Z');
  const daysSince = Math.round((d - base) / 86400000);
  if (daysSince < 0) return null;
  const cycleNum = Math.floor(daysSince / 14);
  const hit = EKIHI_EVENTS.find(e => cycleNum * 14 + e.offset === daysSince);
  return hit ? hit.label : null;
}

function getNextEkihiEvent(today) {
  const base = new Date(EKIHI_BASE + 'T00:00:00Z');
  const d    = new Date(today + 'T00:00:00Z');
  const daysSince = Math.round((d - base) / 86400000);
  if (daysSince < 0) return { label: '入替', date: EKIHI_BASE };
  const cycleNum = Math.floor(daysSince / 14);
  for (let c = cycleNum; c <= cycleNum + 1; c++) {
    for (const evt of EKIHI_EVENTS) {
      const evtDay = c * 14 + evt.offset;
      if (evtDay > daysSince) return { label: evt.label, date: offsetDate(EKIHI_BASE, evtDay) };
    }
  }
  return null;
}

function buildNotesSection(ekihiLabel, memos, alerts, today) {
  // 液肥: today is event day only
  const hasEkihi = !!ekihiLabel;

  // メモ・アラート: filter out lots whose harvest date (sow+34) has already passed
  const alertDates = new Set(alerts.map(a => a.sow_date));
  const allDates = [...new Set([...memos.map(m => m.sow_date), ...alerts.map(a => a.sow_date)])].sort();
  const activeDates = allDates.filter(d => offsetDate(d, 34) >= today);

  if (!hasEkihi && activeDates.length === 0) return '';

  let inner = '';

  // 液肥カード（当日イベントの日のみ）
  if (hasEkihi) {
    const next = getNextEkihiEvent(today);
    inner += `<div style="background:#EEF6FF;border-radius:8px;padding:10px 14px;border:1px solid #c5d8f0;margin-bottom:${activeDates.length > 0 ? '10px' : '0'};">
      <div style="font-size:13px;font-weight:700;color:#185FA5;">&#128167; 本日の液肥: ${ekihiLabel}</div>
      ${next ? `<div style="font-size:12px;color:#555;margin-top:4px;">次回${next.label}: ${next.date}</div>` : ''}
    </div>`;
  }

  // メモ・アラートテーブル（播種日・刈取日・内容）
  if (activeDates.length > 0) {
    let rows = '';
    activeDates.forEach(date => {
      const isAlert = alertDates.has(date);
      const memo    = memos.find(m => m.sow_date === date);
      const harvestDate = offsetDate(date, 34);
      const label   = isAlert ? '&#9888;' : '&#128221;';
      rows += `<tr>
        <td style="padding:7px 10px;white-space:nowrap;color:#555;border-bottom:1px solid #f0f0f0;">${label}&nbsp;${date}</td>
        <td style="padding:7px 10px;white-space:nowrap;color:#555;border-bottom:1px solid #f0f0f0;">${harvestDate}</td>
        <td style="padding:7px 10px;border-bottom:1px solid #f0f0f0;">${memo ? memo.note : ''}</td>
      </tr>`;
    });
    inner += `<table style="width:100%;border-collapse:collapse;font-size:13px;background:#fff;border-radius:8px;overflow:hidden;border:1px solid #e8e8e8;">
      <thead>
        <tr style="background:#f0f0ee;">
          <th style="padding:6px 10px;text-align:left;font-size:11px;color:#888;font-weight:600;white-space:nowrap;">播種日</th>
          <th style="padding:6px 10px;text-align:left;font-size:11px;color:#888;font-weight:600;white-space:nowrap;">刈取日</th>
          <th style="padding:6px 10px;text-align:left;font-size:11px;color:#888;font-weight:600;">内容</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>`;
  }

  return `<div style="margin-bottom:20px;">
    <div style="font-size:14px;font-weight:700;color:#1a1a1a;margin-bottom:8px;">&#128203; メモ・アラート・液肥</div>
    ${inner}
  </div>`;
}
