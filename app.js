// ============================================================
// 股票損益追蹤系統 主程式
// ============================================================

const STORAGE_KEY = 'stockTracker.v1';

// ---------- 全域狀態 ----------
const State = {
  data: null,           // 從 localStorage 載入的資料
  currentAccountId: null,
  charts: {},           // Chart.js 實例
};

// ---------- 資料結構 ----------
function emptyAccount(name, broker = '元大') {
  return {
    id: 'acc_' + Date.now() + '_' + Math.random().toString(36).slice(2,7),
    name,
    broker,
    createdAt: new Date().toISOString(),
    unrealized: [],          // 最新一筆未實現損益
    unrealizedSnapshotDate: null,
    snapshots: [],           // [{date, items, totalMarket, totalCost, totalPL}]
    trades: [],              // 投資明細
    realized: [],            // 已實現（含 interest, shortFee, adjust, note, actualPL）
    adjustments: {}          // {realizedKey: {adjust, note}} 即使重新匯入也保留
  };
}

function emptyData() {
  return {
    version: 1,
    accounts: [],
    currentAccountId: null
  };
}

// ---------- 儲存 ----------
function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.warn('localStorage parse error', e); }
  return emptyData();
}

function save() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(State.data));
  } catch (e) {
    toast('儲存失敗：' + e.message, 'err');
  }
}

// ---------- 工具 ----------
function fmt(n, opts = {}) {
  if (n == null || n === '') return '—';
  if (typeof n !== 'number') return String(n);
  const { decimals = 0, sign = false } = opts;
  const fixed = n.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
  return sign && n > 0 ? '+' + fixed : fixed;
}

function fmtPct(n, decimals = 2) {
  if (n == null || isNaN(n)) return '—';
  return (n > 0 ? '+' : '') + n.toFixed(decimals) + '%';
}

function plClass(n) {
  if (n > 0) return 'pos';
  if (n < 0) return 'neg';
  return '';
}

function todayStr() {
  return Parsers.formatDate(new Date());
}

function toast(msg, kind = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'show ' + kind;
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => t.className = '', 2500);
}

// ---------- 簡易 modal ----------
function showModal({ title, html, onConfirm, confirmText = '確定', cancelText = '取消' }) {
  const root = document.getElementById('modalRoot');
  root.innerHTML = `
    <div class="modal-overlay">
      <div class="modal">
        <h3>${title}</h3>
        <div class="modal-body">${html}</div>
        <div class="modal-actions">
          <button class="btn-mini" data-act="cancel">${cancelText}</button>
          <button class="btn-mini primary" data-act="ok">${confirmText}</button>
        </div>
      </div>
    </div>
  `;
  return new Promise(resolve => {
    root.querySelector('[data-act="cancel"]').onclick = () => { root.innerHTML=''; resolve(null); };
    root.querySelector('[data-act="ok"]').onclick = () => {
      const result = onConfirm ? onConfirm(root.querySelector('.modal-body')) : true;
      root.innerHTML = '';
      resolve(result);
    };
  });
}

function promptText(title, defaultValue = '') {
  return showModal({
    title,
    html: `<input type="text" id="__prompt" value="${defaultValue}" style="width:100%">`,
    onConfirm: (body) => body.querySelector('#__prompt').value.trim()
  });
}

function confirmDialog(title, message) {
  return showModal({
    title,
    html: `<p>${message}</p>`,
    onConfirm: () => true
  });
}

// ============================================================
// 帳戶管理
// ============================================================

function getCurrentAccount() {
  if (!State.data || !State.currentAccountId) return null;
  return State.data.accounts.find(a => a.id === State.currentAccountId);
}

function refreshAccountSelector() {
  const sel = document.getElementById('accountSelect');
  sel.innerHTML = '';
  for (const a of State.data.accounts) {
    const opt = document.createElement('option');
    opt.value = a.id;
    opt.textContent = `${a.broker} - ${a.name}`;
    sel.appendChild(opt);
  }
  sel.value = State.currentAccountId || '';

  document.getElementById('currentAccountName').textContent =
    getCurrentAccount() ? `${getCurrentAccount().broker} - ${getCurrentAccount().name}` : '（無）';
  document.getElementById('accountTitle').textContent =
    getCurrentAccount() ? `帳戶明細：${getCurrentAccount().broker} - ${getCurrentAccount().name}` : '帳戶明細';
}

async function newAccount() {
  const name = await promptText('新增帳戶（券商與名稱）', '');
  if (!name) return;
  // 偵測券商
  let broker = '其他';
  if (/元大/.test(name)) broker = '元大';
  else if (/國泰/.test(name)) broker = '國泰';
  else if (/新光/.test(name)) broker = '新光';
  else {
    const b = await promptText('選擇券商：元大 / 國泰 / 新光 / 其他', '元大');
    if (b) broker = b;
  }
  const acc = emptyAccount(name, broker);
  State.data.accounts.push(acc);
  State.currentAccountId = acc.id;
  State.data.currentAccountId = acc.id;
  save();
  refreshAccountSelector();
  renderAll();
  toast(`已建立帳戶：${broker} - ${name}`, 'ok');
}

async function renameAccount() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const name = await promptText('新名稱', acc.name);
  if (!name) return;
  acc.name = name;
  save();
  refreshAccountSelector();
  toast('已重新命名', 'ok');
}

async function deleteAccount() {
  const acc = getCurrentAccount();
  if (!acc) return;
  const ok = await confirmDialog('刪除帳戶', `確定要刪除「${acc.broker} - ${acc.name}」？所有資料將永久消失。`);
  if (!ok) return;
  State.data.accounts = State.data.accounts.filter(a => a.id !== acc.id);
  State.currentAccountId = State.data.accounts[0]?.id || null;
  State.data.currentAccountId = State.currentAccountId;
  save();
  refreshAccountSelector();
  renderAll();
  toast('已刪除帳戶', 'ok');
}

// ============================================================
// 匯入檔案
// ============================================================

async function readFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: 'array', cellDates: true });
}

function logImport(msg, kind = '') {
  const el = document.getElementById('importLog');
  const line = document.createElement('div');
  if (kind) line.className = kind;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}

async function importUnrealized(file) {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  try {
    const wb = await readFile(file);
    const result = Parsers.parseUnrealized(wb);
    const items = result.items;
    const broker = result.broker;
    const dateInput = document.getElementById('snapshotDate').value;
    const date = dateInput ? Parsers.formatDate(dateInput) : todayStr();

    acc.unrealized = items;
    acc.unrealizedSnapshotDate = date;

    const totalMarket = items.reduce((s, x) => s + x.marketValue, 0);
    const totalCost = items.reduce((s, x) => s + x.cost, 0);
    const totalPL = items.reduce((s, x) => s + x.pl, 0);

    // 加入快照（同日覆蓋）
    acc.snapshots = acc.snapshots.filter(s => s.date !== date);
    acc.snapshots.push({ date, totalMarket, totalCost, totalPL, count: items.length });
    acc.snapshots.sort((a, b) => a.date.localeCompare(b.date));

    save();
    logImport(`✓ 未實現損益匯入成功（${brokerName(broker)}）：${items.length} 筆，總市值 ${fmt(totalMarket)}`, 'ok');
    toast(`未實現損益已匯入（${items.length} 筆）`, 'ok');
    renderAll();
  } catch (e) {
    console.error(e);
    logImport(`✗ 未實現損益匯入失敗：${e.message}`, 'err');
    toast('匯入失敗：' + e.message, 'err');
  }
}

async function importTrades(file) {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const append = document.getElementById('tradesAppend').checked;
  try {
    const wb = await readFile(file);
    const result = Parsers.parseTrades(wb);
    const items = result.items;
    const broker = result.broker;

    if (append) {
      // 用「日期+代號+買賣+數量+價金」當鍵去重
      const seen = new Set(acc.trades.map(t => `${t.date}|${t.code}|${t.action}|${t.qty}|${t.amount}|${t.price}`));
      let added = 0;
      for (const t of items) {
        const k = `${t.date}|${t.code}|${t.action}|${t.qty}|${t.amount}|${t.price}`;
        if (!seen.has(k)) { acc.trades.push(t); seen.add(k); added++; }
      }
      logImport(`✓ 投資明細追加成功（${brokerName(broker)}）：本檔 ${items.length} 筆，新增 ${added} 筆（去重 ${items.length - added} 筆）`, 'ok');
    } else {
      acc.trades = items;
      logImport(`✓ 投資明細匯入成功（${brokerName(broker)}，取代）：${items.length} 筆`, 'ok');
    }

    // 排序
    acc.trades.sort((a, b) => a.date.localeCompare(b.date));

    // 重新比對已實現
    if (acc.realized.length > 0) {
      const r = Parsers.enrichRealizedWithInterest(acc.realized, acc.trades);
      logImport(`  ↻ 重新比對已實現損益：${r.matched}/${r.total} 筆融資融券對到利息`, 'ok');
    }

    save();
    toast(`投資明細已${append?'追加':'匯入'}`, 'ok');
    renderAll();
  } catch (e) {
    console.error(e);
    logImport(`✗ 投資明細匯入失敗：${e.message}`, 'err');
    toast('匯入失敗：' + e.message, 'err');
  }
}

async function importRealized(file) {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const append = document.getElementById('realizedAppend').checked;
  try {
    const wb = await readFile(file);
    const result = Parsers.parseRealized(wb);
    const items = result.items;
    const broker = result.broker;

    let merged;
    if (append) {
      const seen = new Set(acc.realized.map(r => realizedKey(r)));
      merged = [...acc.realized];
      let added = 0;
      for (const r of items) {
        const k = realizedKey(r);
        if (!seen.has(k)) { merged.push(r); seen.add(k); added++; }
      }
      logImport(`✓ 已實現損益追加成功（${brokerName(broker)}）：本檔 ${items.length} 筆，新增 ${added} 筆`, 'ok');
    } else {
      merged = items;
      logImport(`✓ 已實現損益匯入成功（${brokerName(broker)}，取代）：${items.length} 筆`, 'ok');
    }

    // 還原以前儲存的調整金額與備註
    let restored = 0;
    for (const r of merged) {
      const k = realizedKey(r);
      const saved = acc.adjustments[k];
      if (saved) {
        r.adjust = saved.adjust || 0;
        r.note = saved.note || '';
        if (saved.adjust || saved.note) restored++;
      } else {
        r.adjust = 0;
        r.note = '';
      }
    }
    if (restored > 0) logImport(`  ↻ 還原了 ${restored} 筆先前的調整紀錄`, 'ok');

    // 比對投資明細補利息
    const er = Parsers.enrichRealizedWithInterest(merged, acc.trades);
    if (er.total > 0) {
      logImport(`  ↻ 比對投資明細：${er.matched}/${er.total} 筆融資融券交易找到利息資料`, er.matched < er.total ? 'warn' : 'ok');
      if (er.matched < er.total) {
        logImport(`  ⚠ 有 ${er.total - er.matched} 筆沒對到，請確認投資明細是否完整`, 'warn');
      }
    }

    // 排序
    merged.sort((a, b) => (b.sellDate || '').localeCompare(a.sellDate || ''));
    acc.realized = merged;

    save();
    toast(`已實現損益已${append?'追加':'匯入'}`, 'ok');
    renderAll();
  } catch (e) {
    console.error(e);
    logImport(`✗ 已實現損益匯入失敗：${e.message}`, 'err');
    toast('匯入失敗：' + e.message, 'err');
  }
}

function brokerName(b) {
  if (b === 'yuanta') return '元大';
  if (b === 'cathay') return '國泰';
  if (b === 'sks') return '新光';
  return b || '未知';
}

function realizedKey(r) {
  return `${r.code}|${r.sellDate}|${r.qty}|${r.sellPrice}|${r.buyDate}|${r.buyPrice}`;
}

// ============================================================
// 備份 / 還原
// ============================================================

function exportBackup() {
  const data = JSON.stringify(State.data, null, 2);
  const blob = new Blob([data], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const ts = new Date().toISOString().slice(0,10);
  a.download = `stock-tracker-backup-${ts}.json`;
  a.click();
  URL.revokeObjectURL(url);
  toast('備份檔已下載', 'ok');
}

async function restoreBackup(file) {
  try {
    const text = await file.text();
    const data = JSON.parse(text);
    if (!data.accounts || !Array.isArray(data.accounts)) throw new Error('檔案格式不符');
    const ok = await confirmDialog(
      '還原備份',
      `將取代目前所有資料（${State.data.accounts.length} 個帳戶）為備份檔的 ${data.accounts.length} 個帳戶。確定？`
    );
    if (!ok) return;
    State.data = data;
    State.currentAccountId = data.currentAccountId || (data.accounts[0]?.id || null);
    save();
    refreshAccountSelector();
    renderAll();
    toast('還原成功', 'ok');
  } catch (e) {
    toast('還原失敗：' + e.message, 'err');
  }
}

// ============================================================
// 匯出已實現損益（含補上的欄位）
// ============================================================
function exportRealizedExcel() {
  const acc = getCurrentAccount();
  if (!acc || !acc.realized.length) return toast('沒有已實現損益資料', 'err');

  const headers = [
    '代號','名稱','類別',
    '賣出日','賣價','買進日','買價',
    '數量','沖銷成本','手續費','交易稅',
    '利息','融券手續費','調整金額','備註',
    '盈虧','實際盈虧'
  ];
  const rows = acc.realized.map(r => [
    r.code, r.name, r.sellCategory,
    r.sellDate, r.sellPrice, r.buyDate, r.buyPrice,
    r.qty, r.cost, r.fee, r.tax,
    r.interest || 0, r.shortFee || 0, r.adjust || 0, r.note || '',
    r.pl, (r.pl + (r.adjust || 0))
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '已實現損益');
  const ts = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `已實現損益_${acc.broker}_${acc.name}_${ts}.xlsx`);
  toast('已匯出 Excel', 'ok');
}

// ============================================================
// 渲染
// ============================================================

function renderAll() {
  renderOverview();
  renderAccount();
  renderRealized();
  renderUnrealized();
  renderTrades();
}

// ---------- 帳戶聚合 ----------
function aggregateAccount(acc) {
  const totalMarket = acc.unrealized.reduce((s,x) => s+x.marketValue, 0);
  const totalCost   = acc.unrealized.reduce((s,x) => s+x.cost, 0);
  const unrealizedPL= acc.unrealized.reduce((s,x) => s+x.pl, 0);
  const totalInterest = acc.realized.reduce((s,r) => s+(r.interest||0), 0);
  const totalShortFee = acc.realized.reduce((s,r) => s+(r.shortFee||0), 0);
  const realizedPL = acc.realized.reduce((s,r) => s + (r.pl + (r.adjust||0)), 0);
  return { totalMarket, totalCost, unrealizedPL, totalInterest, totalShortFee, realizedPL };
}

function setVal(id, val, withClass = false) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = typeof val === 'number' ? fmt(val, { sign: withClass }) : val;
  if (withClass) el.className = 'value ' + plClass(val);
}

// ---------- 總覽 ----------
function renderOverview() {
  let M=0, C=0, U=0, R=0, I=0, S=0;
  const perAccount = [];
  for (const a of State.data.accounts) {
    const g = aggregateAccount(a);
    M += g.totalMarket; C += g.totalCost; U += g.unrealizedPL;
    R += g.realizedPL; I += g.totalInterest; S += g.totalShortFee;
    perAccount.push({ name: `${a.broker}-${a.name}`, ...g });
  }
  setVal('ovTotalMarket', M);
  setVal('ovTotalCost', C);
  setVal('ovUnrealizedPL', U, true);
  setVal('ovRealizedPL', R, true);
  setVal('ovTotalInterest', I);
  setVal('ovTotalShortFee', S);

  // 表
  const tb = document.querySelector('#ovAccountTable tbody');
  tb.innerHTML = perAccount.length ? perAccount.map(a => `
    <tr>
      <td>${a.name}</td>
      <td>${fmt(a.totalMarket)}</td>
      <td>${fmt(a.totalCost)}</td>
      <td class="${plClass(a.unrealizedPL)}">${fmt(a.unrealizedPL,{sign:true})}</td>
      <td class="${plClass(a.realizedPL)}">${fmt(a.realizedPL,{sign:true})}</td>
      <td>${fmt(a.totalInterest)}</td>
      <td>${fmt(a.totalShortFee)}</td>
    </tr>
  `).join('') : '<tr><td colspan="7" class="empty-state">尚無帳戶資料</td></tr>';

  // 圖
  drawAccountCharts(perAccount);
  drawTrendChart();
}

function drawAccountCharts(perAccount) {
  // 市值圓餅
  if (State.charts.market) State.charts.market.destroy();
  if (State.charts.pl) State.charts.pl.destroy();
  if (!perAccount.length || perAccount.every(a => a.totalMarket === 0)) return;

  const colors = ['#4a9eff','#3ddc84','#ffb547','#ff5f5f','#b46aff','#5fd0d0','#ff8b3d'];
  const ctx1 = document.getElementById('chartAccountMarket').getContext('2d');
  State.charts.market = new Chart(ctx1, {
    type: 'doughnut',
    data: {
      labels: perAccount.map(a => a.name),
      datasets: [{
        data: perAccount.map(a => a.totalMarket),
        backgroundColor: colors,
        borderColor: '#1a2028',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#e6edf3' } },
        title: { display: true, text: '各帳戶市值分布', color: '#e6edf3' }
      }
    }
  });

  // 損益柱狀
  const ctx2 = document.getElementById('chartAccountPL').getContext('2d');
  State.charts.pl = new Chart(ctx2, {
    type: 'bar',
    data: {
      labels: perAccount.map(a => a.name),
      datasets: [
        { label: '未實現', data: perAccount.map(a => a.unrealizedPL), backgroundColor: '#4a9eff' },
        { label: '已實現', data: perAccount.map(a => a.realizedPL), backgroundColor: '#3ddc84' }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#e6edf3' } },
        title: { display: true, text: '各帳戶損益對比', color: '#e6edf3' }
      },
      scales: {
        x: { ticks: { color: '#8b95a5' }, grid: { color: '#303a48' } },
        y: { ticks: { color: '#8b95a5' }, grid: { color: '#303a48' } }
      }
    }
  });
}

function drawTrendChart() {
  if (State.charts.trend) State.charts.trend.destroy();
  // 取所有帳戶所有快照，依日期合併
  const dateMap = new Map(); // date -> { market, cost, pl }
  for (const acc of State.data.accounts) {
    for (const s of (acc.snapshots || [])) {
      const cur = dateMap.get(s.date) || { market:0, cost:0, pl:0 };
      cur.market += s.totalMarket;
      cur.cost += s.totalCost;
      cur.pl += s.totalPL;
      dateMap.set(s.date, cur);
    }
  }
  const dates = [...dateMap.keys()].sort();
  if (!dates.length) return;
  const ctx = document.getElementById('chartTrend').getContext('2d');
  State.charts.trend = new Chart(ctx, {
    type: 'line',
    data: {
      labels: dates,
      datasets: [
        { label: '市值', data: dates.map(d => dateMap.get(d).market), borderColor: '#4a9eff', backgroundColor: 'rgba(74,158,255,0.1)', fill: true, tension: 0.3 },
        { label: '成本', data: dates.map(d => dateMap.get(d).cost), borderColor: '#8b95a5', borderDash: [5,5], fill: false, tension: 0.3 },
        { label: '未實現損益', data: dates.map(d => dateMap.get(d).pl), borderColor: '#3ddc84', fill: false, tension: 0.3 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: '#e6edf3' } } },
      scales: {
        x: { ticks: { color: '#8b95a5' }, grid: { color: '#303a48' } },
        y: { ticks: { color: '#8b95a5' }, grid: { color: '#303a48' } }
      }
    }
  });
}

// ---------- 帳戶頁 ----------
function renderAccount() {
  const acc = getCurrentAccount();
  const status = document.getElementById('acDataStatus');
  if (!acc) {
    setVal('acMarket','—'); setVal('acCost','—'); setVal('acUnrealizedPL','—');
    setVal('acRealizedPL','—'); setVal('acInterest','—'); setVal('acShortFee','—');
    status.innerHTML = '<div class="empty-state">尚未選擇帳戶</div>';
    return;
  }
  const g = aggregateAccount(acc);
  setVal('acMarket', g.totalMarket);
  setVal('acCost', g.totalCost);
  setVal('acUnrealizedPL', g.unrealizedPL, true);
  setVal('acRealizedPL', g.realizedPL, true);
  setVal('acInterest', g.totalInterest);
  setVal('acShortFee', g.totalShortFee);

  const lines = [];
  lines.push(`<p class="hint">未實現損益：<strong>${acc.unrealized.length}</strong> 檔（快照日 ${acc.unrealizedSnapshotDate || '—'}）</p>`);
  lines.push(`<p class="hint">投資明細：<strong>${acc.trades.length}</strong> 筆交易</p>`);
  lines.push(`<p class="hint">已實現損益：<strong>${acc.realized.length}</strong> 筆</p>`);
  lines.push(`<p class="hint">歷史快照：<strong>${(acc.snapshots||[]).length}</strong> 筆</p>`);
  status.innerHTML = lines.join('');
}

// ---------- 已實現損益 ----------
function renderRealized() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#realizedTable tbody');
  if (!acc || !acc.realized.length) {
    tb.innerHTML = '<tr><td colspan="17" class="empty-state">尚無已實現損益資料</td></tr>';
    setVal('rzOriginal','—'); setVal('rzAdjust','—'); setVal('rzActual','—');
    setVal('rzInterest','—'); setVal('rzShortFee','—');
    return;
  }
  const search = (document.getElementById('rzSearch').value || '').toLowerCase();
  let original=0, adjust=0, actual=0, interest=0, shortFee=0;

  const filtered = acc.realized.filter(r => {
    if (!search) return true;
    return r.code.toLowerCase().includes(search) || (r.name || '').toLowerCase().includes(search);
  });

  // 統計用所有資料（不只篩選後）
  for (const r of acc.realized) {
    original += r.pl;
    adjust += (r.adjust || 0);
    actual += r.pl + (r.adjust || 0);
    interest += (r.interest || 0);
    shortFee += (r.shortFee || 0);
  }
  setVal('rzOriginal', original, true);
  setVal('rzAdjust', adjust, true);
  setVal('rzActual', actual, true);
  setVal('rzInterest', interest);
  setVal('rzShortFee', shortFee);

  tb.innerHTML = filtered.map((r, idx) => {
    const realIdx = acc.realized.indexOf(r);
    const actualPL = r.pl + (r.adjust || 0);
    return `
      <tr data-idx="${realIdx}">
        <td>${r.code}</td>
        <td>${r.name}</td>
        <td>${r.sellCategory}</td>
        <td>${r.sellDate}</td>
        <td>${fmt(r.sellPrice, {decimals:2})}</td>
        <td>${r.buyDate || '—'}</td>
        <td>${fmt(r.buyPrice, {decimals:2})}</td>
        <td>${fmt(r.qty)}</td>
        <td>${fmt(r.cost)}</td>
        <td>${fmt(r.fee)}</td>
        <td>${fmt(r.tax)}</td>
        <td class="hl">${fmt(r.interest || 0)}</td>
        <td class="hl">${fmt(r.shortFee || 0)}</td>
        <td class="hl"><input type="number" class="editable-cell" data-field="adjust" value="${r.adjust || 0}"></td>
        <td><input type="text" class="editable-cell note" data-field="note" value="${(r.note || '').replace(/"/g,'&quot;')}" placeholder="備註…"></td>
        <td class="${plClass(r.pl)}">${fmt(r.pl, {sign:true})}</td>
        <td class="hl ${plClass(actualPL)}">${fmt(actualPL, {sign:true})}</td>
      </tr>
    `;
  }).join('');

  // 綁定編輯事件
  tb.querySelectorAll('.editable-cell').forEach(input => {
    input.addEventListener('change', (e) => {
      const tr = e.target.closest('tr');
      const idx = parseInt(tr.dataset.idx, 10);
      const field = e.target.dataset.field;
      const r = acc.realized[idx];
      if (field === 'adjust') {
        r.adjust = parseFloat(e.target.value) || 0;
      } else if (field === 'note') {
        r.note = e.target.value;
      }
      // 同步存到 adjustments，這樣重新匯入會被還原
      const k = realizedKey(r);
      acc.adjustments[k] = { adjust: r.adjust || 0, note: r.note || '' };
      save();
      renderRealized();
      renderOverview();
      renderAccount();
    });
  });
}

// ---------- 未實現損益 ----------
function renderUnrealized() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#unrealizedTable tbody');
  const info = document.getElementById('unrealizedSnapshotInfo');
  const snapBody = document.querySelector('#snapshotTable tbody');
  if (!acc || !acc.unrealized.length) {
    tb.innerHTML = '<tr><td colspan="13" class="empty-state">尚無未實現損益資料</td></tr>';
    info.textContent = '';
    setVal('urMarket','—'); setVal('urCost','—'); setVal('urPL','—'); setVal('urRate','—');
    snapBody.innerHTML = '<tr><td colspan="5" class="empty-state">無快照</td></tr>';
    return;
  }
  let M=0, C=0, P=0;
  for (const x of acc.unrealized) { M+=x.marketValue; C+=x.cost; P+=x.pl; }
  setVal('urMarket', M);
  setVal('urCost', C);
  setVal('urPL', P, true);
  document.getElementById('urPL').className = 'value ' + plClass(P);
  document.getElementById('urRate').textContent = C > 0 ? fmtPct((P / C) * 100) : '—';
  document.getElementById('urRate').className = 'value ' + plClass(P);

  info.textContent = `快照日：${acc.unrealizedSnapshotDate || '—'}　共 ${acc.unrealized.length} 檔`;

  tb.innerHTML = acc.unrealized.map(x => `
    <tr>
      <td>${x.code}</td>
      <td>${x.name}</td>
      <td>${x.category}</td>
      <td>${fmt(x.qty)}</td>
      <td>${fmt(x.price, {decimals:2})}</td>
      <td>${fmt(x.marketValue)}</td>
      <td>${fmt(x.cost)}</td>
      <td>${fmt(x.avgCost, {decimals:4})}</td>
      <td>${fmt(x.fee)}</td>
      <td>${fmt(x.tax)}</td>
      <td>${fmt(x.interest)}</td>
      <td class="${plClass(x.pl)}">${fmt(x.pl, {sign:true})}</td>
      <td class="${plClass(x.pl)}">${x.rate || (x.cost ? fmtPct(x.pl/x.cost*100) : '—')}</td>
    </tr>
  `).join('');

  // 快照表
  snapBody.innerHTML = (acc.snapshots || []).slice().reverse().map(s => `
    <tr>
      <td>${s.date}</td>
      <td>${fmt(s.totalMarket)}</td>
      <td>${fmt(s.totalCost)}</td>
      <td class="${plClass(s.totalPL)}">${fmt(s.totalPL, {sign:true})}</td>
      <td><button class="btn-mini danger" data-snap="${s.date}">刪除</button></td>
    </tr>
  `).join('') || '<tr><td colspan="5" class="empty-state">無快照</td></tr>';

  snapBody.querySelectorAll('button[data-snap]').forEach(btn => {
    btn.onclick = async () => {
      const ok = await confirmDialog('刪除快照', `確定刪除 ${btn.dataset.snap} 的快照？`);
      if (!ok) return;
      acc.snapshots = acc.snapshots.filter(s => s.date !== btn.dataset.snap);
      save(); renderAll();
    };
  });
}

// ---------- 投資明細 ----------
function renderTrades() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#tradesTable tbody');
  if (!acc || !acc.trades.length) {
    tb.innerHTML = '<tr><td colspan="16" class="empty-state">尚無投資明細</td></tr>';
    return;
  }
  const search = (document.getElementById('trSearch').value || '').toLowerCase();
  const ftype = document.getElementById('trFilterType').value;
  const faction = document.getElementById('trFilterAction').value;

  const filtered = acc.trades.filter(t => {
    if (search && !(t.code.toLowerCase().includes(search) || (t.name||'').toLowerCase().includes(search))) return false;
    if (ftype && t.category !== ftype) return false;
    if (faction && t.action !== faction) return false;
    return true;
  });

  tb.innerHTML = filtered.map(t => `
    <tr>
      <td>${t.date}</td>
      <td>${t.code}</td>
      <td>${t.name}</td>
      <td>${t.action}</td>
      <td>${t.category}</td>
      <td>${fmt(t.qty)}</td>
      <td>${fmt(t.price, {decimals:2})}</td>
      <td>${fmt(t.amount)}</td>
      <td>${fmt(t.fee)}</td>
      <td>${fmt(t.tax)}</td>
      <td>${fmt(t.marginAmount)}</td>
      <td class="${t.marginInterest ? 'hl' : ''}">${fmt(t.marginInterest)}</td>
      <td class="${t.shortFee ? 'hl' : ''}">${fmt(t.shortFee)}</td>
      <td>${fmt(t.receivable)}</td>
      <td class="${plClass(t.pl)}">${fmt(t.pl, {sign:true})}</td>
      <td>${t.settleDate || '—'}</td>
    </tr>
  `).join('') || '<tr><td colspan="16" class="empty-state">沒有符合條件的資料</td></tr>';
}

// ============================================================
// 事件綁定
// ============================================================

function bindEvents() {
  // Tab 切換
  document.querySelectorAll('.tab').forEach(btn => {
    btn.onclick = () => {
      document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    };
  });

  // 帳戶選擇
  document.getElementById('accountSelect').onchange = (e) => {
    State.currentAccountId = e.target.value;
    State.data.currentAccountId = e.target.value;
    save();
    refreshAccountSelector();
    renderAll();
  };
  document.getElementById('btnNewAccount').onclick = newAccount;
  document.getElementById('btnRenameAccount').onclick = renameAccount;
  document.getElementById('btnDeleteAccount').onclick = deleteAccount;

  // 備份
  document.getElementById('btnBackup').onclick = exportBackup;
  document.getElementById('btnRestore').onclick = () => document.getElementById('restoreFile').click();
  document.getElementById('restoreFile').onchange = (e) => {
    const f = e.target.files[0]; if (f) restoreBackup(f);
    e.target.value = '';
  };

  // 上傳
  bindUpload('upUnrealized', importUnrealized);
  bindUpload('upTrades', importTrades);
  bindUpload('upRealized', importRealized);

  // 快照日期預設今天
  document.getElementById('snapshotDate').valueAsDate = new Date();

  // 搜尋
  document.getElementById('rzSearch').oninput = renderRealized;
  document.getElementById('trSearch').oninput = renderTrades;
  document.getElementById('trFilterType').onchange = renderTrades;
  document.getElementById('trFilterAction').onchange = renderTrades;

  // 匯出已實現
  document.getElementById('btnExportRealized').onclick = exportRealizedExcel;

  // 清空帳戶
  document.getElementById('btnClearAccount').onclick = async () => {
    const acc = getCurrentAccount();
    if (!acc) return;
    const ok = await confirmDialog('清空帳戶資料', `確定清空「${acc.broker} - ${acc.name}」的所有未實現、投資明細、已實現資料？（帳戶與調整紀錄保留）`);
    if (!ok) return;
    acc.unrealized = []; acc.trades = []; acc.realized = []; acc.snapshots = [];
    save(); renderAll(); toast('已清空', 'ok');
  };
}

function bindUpload(inputId, handler) {
  const input = document.getElementById(inputId);
  const area = input.closest('.upload-area');
  input.onchange = (e) => {
    const f = e.target.files[0];
    if (f) handler(f);
    e.target.value = '';
  };
  area.ondragover = (e) => { e.preventDefault(); area.classList.add('dragover'); };
  area.ondragleave = () => area.classList.remove('dragover');
  area.ondrop = (e) => {
    e.preventDefault(); area.classList.remove('dragover');
    const f = e.dataTransfer.files[0];
    if (f) handler(f);
  };
}

// ============================================================
// 啟動
// ============================================================
async function init() {
  State.data = load();

  // 沒有帳戶就建立一個預設的
  if (!State.data.accounts.length) {
    State.data.accounts.push(emptyAccount('主帳戶', '元大'));
    State.data.currentAccountId = State.data.accounts[0].id;
    save();
  }
  State.currentAccountId = State.data.currentAccountId || State.data.accounts[0].id;

  bindEvents();
  refreshAccountSelector();
  renderAll();
}

document.addEventListener('DOMContentLoaded', init);
