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
    adjustments: {},         // {realizedKey: {adjust, note}} 即使重新匯入也保留
    loans: []                // 借款紀錄 [{id, purpose, principal, rate, startDate, dueDate, status, interestPayments:[], repayments:[]}]
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
    if (raw) {
      const data = JSON.parse(raw);
      // 資料遷移：舊版本帳戶補上 loans 欄位
      if (data.accounts) {
        for (const acc of data.accounts) {
          if (!Array.isArray(acc.loans)) acc.loans = [];
        }
      }
      return data;
    }
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
    opt.textContent = a.name; // 只顯示帳戶名（已含券商前綴）
    sel.appendChild(opt);
  }
  sel.value = State.currentAccountId || '';

  document.getElementById('currentAccountName').textContent =
    getCurrentAccount() ? getCurrentAccount().name : '（無）';
  document.getElementById('accountTitle').textContent =
    getCurrentAccount() ? `帳戶明細：${getCurrentAccount().name}` : '帳戶明細';
}

async function newAccount() {
  const name = await promptText('新增帳戶（格式：券商－名字，例如 元大－音）', '');
  if (!name) return;

  // 從名稱前綴偵測券商（用來自動辨識匯入檔案的格式）
  let broker = '其他';
  if (/元大/.test(name)) broker = '元大';
  else if (/國泰/.test(name)) broker = '國泰';
  else if (/新光/.test(name)) broker = '新光';

  const acc = emptyAccount(name, broker);
  State.data.accounts.push(acc);
  State.currentAccountId = acc.id;
  State.data.currentAccountId = acc.id;
  save();
  refreshAccountSelector();
  renderAll();
  toast(`已建立帳戶：${name}`, 'ok');
}

async function renameAccount() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const name = await promptText('新名稱（格式：券商－名字）', acc.name);
  if (!name) return;
  acc.name = name;
  // 重新偵測券商
  if (/元大/.test(name)) acc.broker = '元大';
  else if (/國泰/.test(name)) acc.broker = '國泰';
  else if (/新光/.test(name)) acc.broker = '新光';
  save();
  refreshAccountSelector();
  renderAll();
  toast('已重新命名', 'ok');
}

async function deleteAccount() {
  const acc = getCurrentAccount();
  if (!acc) return;
  const ok = await confirmDialog('刪除帳戶', `確定要刪除「${acc.name}」？所有資料將永久消失。`);
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
  XLSX.writeFile(wb, `已實現損益_${acc.name}_${ts}.xlsx`);
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
  renderLoans();
}

// ---------- 帳戶聚合 ----------
function aggregateAccount(acc) {
  const totalMarket = acc.unrealized.reduce((s,x) => s+x.marketValue, 0);
  const totalCost   = acc.unrealized.reduce((s,x) => s+x.cost, 0);
  const unrealizedPL= acc.unrealized.reduce((s,x) => s+x.pl, 0);
  const totalInterest = acc.realized.reduce((s,r) => s+(r.interest||0), 0);
  const totalShortFee = acc.realized.reduce((s,r) => s+(r.shortFee||0), 0);
  const realizedPLRaw = acc.realized.reduce((s,r) => s + (r.pl + (r.adjust||0)), 0);

  // 借款利息累計
  const loans = acc.loans || [];
  const loanInterestPaid = loans.reduce((s, l) =>
    s + (l.interestPayments || []).reduce((ss, p) => ss + (p.amount || 0), 0), 0);
  const loanPrincipal = loans.reduce((s, l) => s + (l.principal || 0), 0);
  const loanRepaid = loans.reduce((s, l) =>
    s + (l.repayments || []).reduce((ss, p) => ss + (p.amount || 0), 0), 0);
  const loanRemaining = loanPrincipal - loanRepaid;

  // 已實現損益要扣借款利息（你要求的：借款利息計入實際損益）
  const realizedPL = realizedPLRaw - loanInterestPaid;

  return {
    totalMarket, totalCost, unrealizedPL,
    totalInterest, totalShortFee,
    realizedPL,        // 已扣借款利息
    realizedPLRaw,     // 未扣借款利息（純股票交易實現損益）
    loanInterestPaid,
    loanPrincipal, loanRepaid, loanRemaining
  };
}

function setVal(id, val, withClass = false) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = typeof val === 'number' ? fmt(val, { sign: withClass }) : val;
  if (withClass) el.className = 'value ' + plClass(val);
}

// ---------- 總覽 ----------
function renderOverview() {
  let M=0, C=0, U=0, R=0, I=0, S=0, LI=0, LB=0;
  const perAccount = [];
  for (const a of State.data.accounts) {
    const g = aggregateAccount(a);
    M += g.totalMarket; C += g.totalCost; U += g.unrealizedPL;
    R += g.realizedPL; I += g.totalInterest; S += g.totalShortFee;
    LI += g.loanInterestPaid; LB += g.loanRemaining;
    perAccount.push({ name: a.name, ...g });
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

  const colors = ['#2563eb','#059669','#d97706','#dc2626','#7c3aed','#0891b2','#db2777'];
  const ctx1 = document.getElementById('chartAccountMarket').getContext('2d');
  State.charts.market = new Chart(ctx1, {
    type: 'doughnut',
    data: {
      labels: perAccount.map(a => a.name),
      datasets: [{
        data: perAccount.map(a => a.totalMarket),
        backgroundColor: colors,
        borderColor: '#ffffff',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#1f2937' } },
        title: { display: true, text: '各帳戶市值分布', color: '#1f2937' }
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
        { label: '未實現', data: perAccount.map(a => a.unrealizedPL), backgroundColor: '#2563eb' },
        { label: '已實現', data: perAccount.map(a => a.realizedPL), backgroundColor: '#059669' }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#1f2937' } },
        title: { display: true, text: '各帳戶損益對比', color: '#1f2937' }
      },
      scales: {
        x: { ticks: { color: '#6b7280' }, grid: { color: '#e8edf3' } },
        y: { ticks: { color: '#6b7280' }, grid: { color: '#e8edf3' } }
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
        { label: '市值', data: dates.map(d => dateMap.get(d).market), borderColor: '#2563eb', backgroundColor: 'rgba(37,99,235,0.1)', fill: true, tension: 0.3 },
        { label: '成本', data: dates.map(d => dateMap.get(d).cost), borderColor: '#9ca3af', borderDash: [5,5], fill: false, tension: 0.3 },
        { label: '未實現損益', data: dates.map(d => dateMap.get(d).pl), borderColor: '#059669', fill: false, tension: 0.3 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: '#1f2937' } } },
      scales: {
        x: { ticks: { color: '#6b7280' }, grid: { color: '#e8edf3' } },
        y: { ticks: { color: '#6b7280' }, grid: { color: '#e8edf3' } }
      }
    }
  });
}

// ---------- 帳戶頁 ----------
function renderAccount() {
  const acc = getCurrentAccount();
  const status = document.getElementById('acDataStatus');
  if (!acc) {
    ['acMarket','acCost','acUnrealizedPL','acRealizedPL','acInterest','acShortFee','acLoanInterest','acLoanBalance']
      .forEach(id => setVal(id, '—'));
    status.innerHTML = '<div class="empty-state">尚未選擇帳戶</div>';
    document.querySelector('#monthlyTable tbody').innerHTML =
      '<tr><td colspan="10" class="empty-state">尚未選擇帳戶</td></tr>';
    return;
  }
  const g = aggregateAccount(acc);
  setVal('acMarket', g.totalMarket);
  setVal('acCost', g.totalCost);
  setVal('acUnrealizedPL', g.unrealizedPL, true);
  setVal('acRealizedPL', g.realizedPL, true);
  setVal('acInterest', g.totalInterest);
  setVal('acShortFee', g.totalShortFee);
  setVal('acLoanInterest', g.loanInterestPaid);
  setVal('acLoanBalance', g.loanRemaining);

  const lines = [];
  lines.push(`<p class="hint">未實現損益：<strong>${acc.unrealized.length}</strong> 檔（快照日 ${acc.unrealizedSnapshotDate || '—'}）　|　投資明細：<strong>${acc.trades.length}</strong> 筆　|　已實現損益：<strong>${acc.realized.length}</strong> 筆　|　歷史快照：<strong>${(acc.snapshots||[]).length}</strong> 筆　|　借款：<strong>${(acc.loans||[]).length}</strong> 筆</p>`);
  status.innerHTML = lines.join('');

  renderMonthly();
}

// ---------- 每月損益彙總 ----------
function getMonthKey(dateStr) {
  if (!dateStr) return '';
  const m = dateStr.match(/^(\d{4})\/(\d{1,2})/);
  if (!m) return '';
  return `${m[1]}-${m[2].padStart(2,'0')}`;
}

function buildMonthlyData(acc) {
  // key: 'YYYY-MM' -> { realizedPL, interest, shortFee, adjust, loanInterest, tradeCount, tradeAmount, realizedItems[], tradeItems[] }
  const months = new Map();
  const ensure = (k) => {
    if (!months.has(k)) months.set(k, {
      key: k, realizedPL: 0, interest: 0, shortFee: 0,
      adjust: 0, loanInterest: 0,
      tradeCount: 0, tradeAmount: 0,
      realizedItems: [], tradeItems: []
    });
    return months.get(k);
  };

  // 已實現：以賣出日為準
  for (const r of (acc.realized || [])) {
    const k = getMonthKey(r.sellDate);
    if (!k) continue;
    const m = ensure(k);
    m.realizedPL += (r.pl || 0);
    m.interest += (r.interest || 0);
    m.shortFee += (r.shortFee || 0);
    m.adjust += (r.adjust || 0);
    m.realizedItems.push(r);
  }

  // 投資明細：以成交日為準
  for (const t of (acc.trades || [])) {
    const k = getMonthKey(t.date);
    if (!k) continue;
    const m = ensure(k);
    m.tradeCount++;
    m.tradeAmount += (t.amount || 0);
    m.tradeItems.push(t);
  }

  // 借款利息支付：以付款日為準
  for (const l of (acc.loans || [])) {
    for (const p of (l.interestPayments || [])) {
      const k = getMonthKey(p.date);
      if (!k) continue;
      const m = ensure(k);
      m.loanInterest += (p.amount || 0);
    }
  }

  // 計算實際損益（已實現損益 + 調整 - 借款利息）
  // 注意：盈虧已經扣過融資利息和融券手續費
  for (const m of months.values()) {
    m.actual = m.realizedPL + m.adjust - m.loanInterest;
  }

  return [...months.values()].sort((a, b) => b.key.localeCompare(a.key));
}

const _expandedMonths = new Set();

function renderMonthly() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#monthlyTable tbody');
  const yearSel = document.getElementById('monthlyYear');
  if (!acc) return;

  const data = buildMonthlyData(acc);

  // 年份下拉
  const years = [...new Set(data.map(m => m.key.slice(0,4)))].sort().reverse();
  const currentYear = yearSel.value;
  yearSel.innerHTML = '<option value="">所有年份</option>' +
    years.map(y => `<option value="${y}" ${y===currentYear?'selected':''}>${y}</option>`).join('');

  const filtered = currentYear ? data.filter(m => m.key.startsWith(currentYear)) : data;

  if (!filtered.length) {
    tb.innerHTML = '<tr><td colspan="10" class="empty-state">尚無資料（請先匯入已實現損益或投資明細）</td></tr>';
    return;
  }

  const rows = [];
  for (const m of filtered) {
    const expanded = _expandedMonths.has(m.key);
    const [year, month] = m.key.split('-');
    rows.push(`
      <tr class="month-row" data-month="${m.key}">
        <td><span class="month-toggle ${expanded?'expanded':''}">▶</span></td>
        <td><span class="year-tag">${year}</span><strong>${parseInt(month)}月</strong></td>
        <td class="${plClass(m.realizedPL)}">${fmt(m.realizedPL,{sign:true})}</td>
        <td>${fmt(m.interest)}</td>
        <td>${fmt(m.shortFee)}</td>
        <td class="${m.loanInterest?'neg':''}">${m.loanInterest ? '-'+fmt(m.loanInterest) : '—'}</td>
        <td class="${plClass(m.adjust)}">${m.adjust?fmt(m.adjust,{sign:true}):'—'}</td>
        <td class="hl ${plClass(m.actual)}">${fmt(m.actual,{sign:true})}</td>
        <td>${m.tradeCount}</td>
        <td>${fmt(m.tradeAmount)}</td>
      </tr>
    `);
    if (expanded) {
      rows.push(`<tr class="month-detail-row"><td colspan="10">${renderMonthDetail(m)}</td></tr>`);
    }
  }
  tb.innerHTML = rows.join('');

  // 綁定展開
  tb.querySelectorAll('.month-row').forEach(row => {
    row.onclick = () => {
      const k = row.dataset.month;
      if (_expandedMonths.has(k)) _expandedMonths.delete(k); else _expandedMonths.add(k);
      renderMonthly();
    };
  });
}

function renderMonthDetail(m) {
  const html = [];

  // 已實現損益細項
  if (m.realizedItems.length > 0) {
    html.push('<div class="nested">');
    html.push('<h4 style="padding:10px 14px 0">💰 已實現損益細項（' + m.realizedItems.length + ' 筆）</h4>');
    html.push('<table class="loan-subtable"><thead><tr>');
    html.push('<th>代號</th><th>名稱</th><th>類別</th><th>賣出日</th><th>買進日</th><th>數量</th><th>賣價</th><th>買價</th><th>盈虧</th><th>利息</th><th>融券費</th><th>調整</th><th>實際</th>');
    html.push('</tr></thead><tbody>');
    for (const r of m.realizedItems) {
      const actual = r.pl + (r.adjust || 0);
      html.push(`<tr>
        <td>${r.code}</td><td>${r.name||''}</td><td>${r.sellCategory||''}</td>
        <td>${r.sellDate}</td><td>${r.buyDate||'—'}</td>
        <td>${fmt(r.qty)}</td>
        <td>${fmt(r.sellPrice,{decimals:2})}</td>
        <td>${fmt(r.buyPrice,{decimals:2})}</td>
        <td class="${plClass(r.pl)}">${fmt(r.pl,{sign:true})}</td>
        <td>${fmt(r.interest||0)}</td>
        <td>${fmt(r.shortFee||0)}</td>
        <td class="${plClass(r.adjust||0)}">${(r.adjust||0)?fmt(r.adjust,{sign:true}):'—'}</td>
        <td class="${plClass(actual)}"><strong>${fmt(actual,{sign:true})}</strong></td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
  }

  // 該月投資明細（交易紀錄）
  if (m.tradeItems.length > 0) {
    html.push('<div class="nested">');
    html.push('<h4 style="padding:10px 14px 0">📋 投資明細（' + m.tradeItems.length + ' 筆）</h4>');
    html.push('<table class="loan-subtable"><thead><tr>');
    html.push('<th>日期</th><th>代號</th><th>名稱</th><th>買賣</th><th>類別</th><th>數量</th><th>單價</th><th>價金</th><th>手續費</th><th>交易稅</th><th>利息</th><th>損益</th>');
    html.push('</tr></thead><tbody>');
    for (const t of m.tradeItems) {
      html.push(`<tr>
        <td>${t.date}</td><td>${t.code}</td><td>${t.name||''}</td>
        <td>${t.action||''}</td><td>${t.category||''}</td>
        <td>${fmt(t.qty)}</td>
        <td>${fmt(t.price,{decimals:2})}</td>
        <td>${fmt(t.amount)}</td>
        <td>${fmt(t.fee)}</td>
        <td>${fmt(t.tax)}</td>
        <td>${fmt(t.marginInterest||0)}</td>
        <td class="${plClass(t.pl||0)}">${(t.pl||0)?fmt(t.pl,{sign:true}):'—'}</td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
  }

  // 該月借款利息支付
  const loanPayments = [];
  for (const l of (getCurrentAccount().loans || [])) {
    for (const p of (l.interestPayments || [])) {
      if (getMonthKey(p.date) === m.key) {
        loanPayments.push({ ...p, loanId: l.id, purpose: l.purpose });
      }
    }
  }
  if (loanPayments.length > 0) {
    html.push('<div class="nested">');
    html.push('<h4 style="padding:10px 14px 0">🏛️ 該月借款利息支付（' + loanPayments.length + ' 筆）</h4>');
    html.push('<table class="loan-subtable"><thead><tr>');
    html.push('<th>日期</th><th>借款編號</th><th>用途</th><th>金額</th><th>備註</th>');
    html.push('</tr></thead><tbody>');
    for (const p of loanPayments) {
      html.push(`<tr>
        <td>${p.date}</td><td>${p.loanId}</td><td>${p.purpose||''}</td>
        <td class="neg">-${fmt(p.amount)}</td>
        <td>${p.note||''}</td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
  }

  if (html.length === 0) html.push('<div class="empty-state">該月無細項資料</div>');
  return html.join('');
}

// ---------- 匯出每月損益 Excel ----------
function exportMonthlyExcel() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const data = buildMonthlyData(acc);
  if (!data.length) return toast('沒有資料可匯出', 'err');

  const headers = ['月份','已實現損益','融資利息','融券手續費','借款利息','調整金額','實際損益','交易筆數','總成交金額'];
  const rows = data.map(m => [
    m.key, m.realizedPL, m.interest, m.shortFee,
    m.loanInterest, m.adjust, m.actual,
    m.tradeCount, m.tradeAmount
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '每月損益');
  const ts = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `每月損益_${acc.name}_${ts}.xlsx`);
  toast('已匯出 Excel', 'ok');
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

// ============================================================
// 股票借款
// ============================================================

function loanGenId(acc) {
  const existing = (acc.loans || []).map(l => l.id);
  let n = existing.length + 1;
  while (existing.includes('LOAN-' + String(n).padStart(3,'0'))) n++;
  return 'LOAN-' + String(n).padStart(3, '0');
}

function calcEstimatedInterest(loan, asOfDate = new Date()) {
  // 預估應付利息 = 本金 × 年利率 × 天數 / 365
  // 還款部分需從還款日起停止計息（簡化：用平均剩餘本金估算）
  const rate = (loan.rate || 0) / 100;
  const start = new Date(loan.startDate);
  if (isNaN(start)) return 0;
  const end = loan.status === 'settled' && loan.repayments && loan.repayments.length > 0
    ? new Date(loan.repayments[loan.repayments.length-1].date)
    : asOfDate;
  if (isNaN(end) || end < start) return 0;

  // 用「分段計息」：從起始日開始，依還款事件變化本金
  const events = (loan.repayments || []).slice().sort((a,b) => a.date.localeCompare(b.date));
  let principal = loan.principal || 0;
  let cursor = start;
  let total = 0;

  for (const ev of events) {
    const evDate = new Date(ev.date);
    if (isNaN(evDate) || evDate <= cursor) continue;
    if (evDate > end) break;
    const days = (evDate - cursor) / 86400000;
    total += principal * rate * days / 365;
    principal -= (ev.amount || 0);
    cursor = evDate;
  }
  // 最後一段（cursor → end）
  if (principal > 0 && end > cursor) {
    const days = (end - cursor) / 86400000;
    total += principal * rate * days / 365;
  }
  return Math.round(total);
}

function renderLoans() {
  const acc = getCurrentAccount();
  const list = document.getElementById('loanList');
  const nameSpan = document.getElementById('loansAccountName');
  if (!acc) {
    nameSpan.textContent = '—';
    list.innerHTML = '<div class="empty-state">尚未選擇帳戶</div>';
    ['loanTotalPrincipal','loanTotalRepaid','loanRemaining','loanInterestPaid','loanInterestEstimated','loanActiveCount']
      .forEach(id => setVal(id, '—'));
    return;
  }
  nameSpan.textContent = acc.name;

  const loans = acc.loans || [];
  const totalPrincipal = loans.reduce((s,l) => s + (l.principal || 0), 0);
  const totalRepaid = loans.reduce((s,l) => s + (l.repayments||[]).reduce((ss,r) => ss+(r.amount||0), 0), 0);
  const interestPaid = loans.reduce((s,l) => s + (l.interestPayments||[]).reduce((ss,r) => ss+(r.amount||0), 0), 0);
  const interestEst = loans.reduce((s,l) => s + calcEstimatedInterest(l), 0);
  const activeCount = loans.filter(l => l.status !== 'settled').length;

  setVal('loanTotalPrincipal', totalPrincipal);
  setVal('loanTotalRepaid', totalRepaid);
  setVal('loanRemaining', totalPrincipal - totalRepaid);
  setVal('loanInterestPaid', interestPaid);
  setVal('loanInterestEstimated', interestEst);
  setVal('loanActiveCount', activeCount);

  if (!loans.length) {
    list.innerHTML = '<div class="empty-state">尚無借款紀錄。點上方「＋ 新增借款」開始記錄。</div>';
    return;
  }

  list.innerHTML = loans.map(l => renderLoanCard(l)).join('');
  bindLoanCardEvents();
}

function renderLoanCard(loan) {
  const principal = loan.principal || 0;
  const repaid = (loan.repayments || []).reduce((s,r) => s + (r.amount||0), 0);
  const remaining = principal - repaid;
  const interestPaid = (loan.interestPayments || []).reduce((s,r) => s + (r.amount||0), 0);
  const interestEst = calcEstimatedInterest(loan);
  const settled = loan.status === 'settled';

  let interestRows = (loan.interestPayments || []).slice().reverse().map((p, i) => `
    <tr>
      <td>${p.date}</td>
      <td>${fmt(p.amount)}</td>
      <td>${p.note || ''}</td>
      <td><button class="btn-mini danger" data-act="del-int" data-loan="${loan.id}" data-idx="${(loan.interestPayments||[]).length-1-i}">刪除</button></td>
    </tr>
  `).join('') || '<tr><td colspan="4" class="empty-state">無紀錄</td></tr>';

  let repayRows = (loan.repayments || []).slice().reverse().map((p, i) => {
    // 計算當時還款後剩餘本金
    let bal = principal;
    const sorted = (loan.repayments || []).slice().sort((a,b) => a.date.localeCompare(b.date));
    for (const r of sorted) {
      bal -= (r.amount || 0);
      if (r === p || (r.date===p.date && r.amount===p.amount && r.note===p.note)) break;
    }
    return `
    <tr>
      <td>${p.date}</td>
      <td>${fmt(p.amount)}</td>
      <td>${fmt(Math.max(0, bal))}</td>
      <td>${p.note || ''}</td>
      <td><button class="btn-mini danger" data-act="del-rep" data-loan="${loan.id}" data-idx="${(loan.repayments||[]).length-1-i}">刪除</button></td>
    </tr>
  `;}).join('') || '<tr><td colspan="5" class="empty-state">無紀錄</td></tr>';

  const diff = interestPaid - interestEst;
  const diffStr = diff === 0 ? '—' : (diff > 0 ? '多付 +' + fmt(Math.abs(diff)) : '少付 -' + fmt(Math.abs(diff)));

  return `
    <div class="loan-card">
      <div class="loan-card-header ${settled?'settled':''}">
        <span class="loan-id">${loan.id}</span>
        <span class="loan-purpose">${loan.purpose || '（無說明）'}</span>
        <span class="loan-status ${settled?'settled':''}">${settled?'已結清':'進行中'}</span>
      </div>
      <div class="loan-card-body">
        <div class="loan-fields">
          <div class="loan-field"><div class="label">借款金額</div><div class="value">${fmt(principal)}</div></div>
          <div class="loan-field"><div class="label">年利率</div><div class="value">${(loan.rate||0).toFixed(2)}%</div></div>
          <div class="loan-field"><div class="label">起始日</div><div class="value">${loan.startDate || '—'}</div></div>
          <div class="loan-field"><div class="label">到期日</div><div class="value">${loan.dueDate || '—'}</div></div>
          <div class="loan-field"><div class="label">已還款</div><div class="value">${fmt(repaid)}</div></div>
          <div class="loan-field"><div class="label">剩餘本金</div><div class="value ${remaining>0?'':'pos'}">${fmt(Math.max(0, remaining))}</div></div>
          <div class="loan-field"><div class="label">已付利息</div><div class="value">${fmt(interestPaid)}</div></div>
          <div class="loan-field"><div class="label">預估應付利息</div><div class="value">${fmt(interestEst)}</div></div>
          <div class="loan-field"><div class="label">實付 vs 預估</div><div class="value ${diff>0?'neg':(diff<0?'pos':'')}">${diffStr}</div></div>
        </div>

        <div class="loan-subsection">
          <h4>💸 利息支付紀錄</h4>
          <table class="loan-subtable">
            <thead><tr><th>日期</th><th>金額</th><th>備註</th><th></th></tr></thead>
            <tbody>${interestRows}</tbody>
          </table>
        </div>

        <div class="loan-subsection">
          <h4>💰 還款紀錄</h4>
          <table class="loan-subtable">
            <thead><tr><th>日期</th><th>還款金額</th><th>還款後本金</th><th>備註</th><th></th></tr></thead>
            <tbody>${repayRows}</tbody>
          </table>
        </div>

        <div class="loan-actions">
          <button class="btn-mini primary" data-act="add-int" data-loan="${loan.id}">＋ 新增利息支付</button>
          <button class="btn-mini primary" data-act="add-rep" data-loan="${loan.id}">＋ 新增還款</button>
          <button class="btn-mini" data-act="edit" data-loan="${loan.id}">編輯</button>
          <button class="btn-mini" data-act="toggle" data-loan="${loan.id}">${settled?'標為進行中':'標為已結清'}</button>
          <button class="btn-mini danger" data-act="delete" data-loan="${loan.id}">刪除借款</button>
        </div>
      </div>
    </div>
  `;
}

function bindLoanCardEvents() {
  document.querySelectorAll('#loanList button[data-act]').forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      const act = btn.dataset.act;
      const loanId = btn.dataset.loan;
      const acc = getCurrentAccount();
      const loan = acc.loans.find(l => l.id === loanId);
      if (!loan) return;

      if (act === 'add-int') await addInterestPayment(loan);
      else if (act === 'add-rep') await addRepayment(loan);
      else if (act === 'edit') await editLoan(loan);
      else if (act === 'toggle') {
        loan.status = loan.status === 'settled' ? 'active' : 'settled';
        save(); renderLoans(); renderAll();
        toast(loan.status === 'settled' ? '已標為已結清' : '已標為進行中', 'ok');
      }
      else if (act === 'delete') {
        const ok = await confirmDialog('刪除借款', `確定刪除「${loan.id} - ${loan.purpose||''}」？所有利息和還款紀錄會一併消失。`);
        if (ok) {
          acc.loans = acc.loans.filter(l => l.id !== loanId);
          save(); renderLoans(); renderAll();
          toast('已刪除', 'ok');
        }
      }
      else if (act === 'del-int') {
        const idx = parseInt(btn.dataset.idx, 10);
        loan.interestPayments.splice(idx, 1);
        save(); renderLoans(); renderAll();
      }
      else if (act === 'del-rep') {
        const idx = parseInt(btn.dataset.idx, 10);
        loan.repayments.splice(idx, 1);
        // 重新評估狀態
        const repaid = loan.repayments.reduce((s,r) => s+(r.amount||0), 0);
        if (repaid < loan.principal) loan.status = 'active';
        save(); renderLoans(); renderAll();
      }
    };
  });
}

async function addLoan() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const today = todayStr().replace(/\//g, '-');
  const result = await showModal({
    title: '新增借款',
    html: `
      <div style="display:grid;gap:10px">
        <label>用途/備註<input type="text" id="lf-purpose" placeholder="例：買保瑞融資自備款"></label>
        <label>借款金額<input type="number" id="lf-principal" placeholder="500000"></label>
        <label>年利率(%)<input type="number" id="lf-rate" placeholder="2.5" step="0.01"></label>
        <label>起始日<input type="date" id="lf-start" value="${today}"></label>
        <label>到期日<input type="date" id="lf-due"></label>
      </div>
    `,
    onConfirm: (body) => ({
      purpose: body.querySelector('#lf-purpose').value.trim(),
      principal: parseFloat(body.querySelector('#lf-principal').value) || 0,
      rate: parseFloat(body.querySelector('#lf-rate').value) || 0,
      startDate: body.querySelector('#lf-start').value,
      dueDate: body.querySelector('#lf-due').value
    })
  });
  if (!result || !result.principal) return;

  const loan = {
    id: loanGenId(acc),
    purpose: result.purpose,
    principal: result.principal,
    rate: result.rate,
    startDate: Parsers.formatDate(result.startDate),
    dueDate: result.dueDate ? Parsers.formatDate(result.dueDate) : '',
    status: 'active',
    interestPayments: [],
    repayments: [],
    createdAt: new Date().toISOString()
  };
  acc.loans.push(loan);
  save();
  renderLoans();
  renderAll();
  toast(`已新增借款 ${loan.id}`, 'ok');
}

async function editLoan(loan) {
  const result = await showModal({
    title: `編輯借款 ${loan.id}`,
    html: `
      <div style="display:grid;gap:10px">
        <label>用途/備註<input type="text" id="lf-purpose" value="${(loan.purpose||'').replace(/"/g,'&quot;')}"></label>
        <label>借款金額<input type="number" id="lf-principal" value="${loan.principal||0}"></label>
        <label>年利率(%)<input type="number" id="lf-rate" value="${loan.rate||0}" step="0.01"></label>
        <label>起始日<input type="date" id="lf-start" value="${(loan.startDate||'').replace(/\//g,'-')}"></label>
        <label>到期日<input type="date" id="lf-due" value="${(loan.dueDate||'').replace(/\//g,'-')}"></label>
      </div>
    `,
    onConfirm: (body) => ({
      purpose: body.querySelector('#lf-purpose').value.trim(),
      principal: parseFloat(body.querySelector('#lf-principal').value) || 0,
      rate: parseFloat(body.querySelector('#lf-rate').value) || 0,
      startDate: body.querySelector('#lf-start').value,
      dueDate: body.querySelector('#lf-due').value
    })
  });
  if (!result) return;
  loan.purpose = result.purpose;
  loan.principal = result.principal;
  loan.rate = result.rate;
  loan.startDate = Parsers.formatDate(result.startDate);
  loan.dueDate = result.dueDate ? Parsers.formatDate(result.dueDate) : '';
  save();
  renderLoans();
  renderAll();
  toast('已更新', 'ok');
}

async function addInterestPayment(loan) {
  const today = todayStr().replace(/\//g, '-');
  const result = await showModal({
    title: `新增利息支付 - ${loan.id}`,
    html: `
      <div style="display:grid;gap:10px">
        <label>日期<input type="date" id="ip-date" value="${today}"></label>
        <label>金額<input type="number" id="ip-amount" placeholder="例：1042"></label>
        <label>備註<input type="text" id="ip-note" placeholder="可不填"></label>
      </div>
    `,
    onConfirm: (body) => ({
      date: body.querySelector('#ip-date').value,
      amount: parseFloat(body.querySelector('#ip-amount').value) || 0,
      note: body.querySelector('#ip-note').value.trim()
    })
  });
  if (!result || !result.amount) return;
  if (!loan.interestPayments) loan.interestPayments = [];
  loan.interestPayments.push({
    date: Parsers.formatDate(result.date),
    amount: result.amount,
    note: result.note
  });
  loan.interestPayments.sort((a,b) => a.date.localeCompare(b.date));
  save();
  renderLoans();
  renderAll();
  toast('已新增利息支付', 'ok');
}

async function addRepayment(loan) {
  const today = todayStr().replace(/\//g, '-');
  const repaid = (loan.repayments||[]).reduce((s,r) => s+(r.amount||0), 0);
  const remaining = (loan.principal||0) - repaid;
  const result = await showModal({
    title: `新增還款 - ${loan.id}`,
    html: `
      <div style="display:grid;gap:10px">
        <label>日期<input type="date" id="rp-date" value="${today}"></label>
        <label>還款金額（剩餘 ${fmt(remaining)}）<input type="number" id="rp-amount"></label>
        <label>備註<input type="text" id="rp-note" placeholder="可不填"></label>
        <label><input type="checkbox" id="rp-settle"> 此為全額還清</label>
      </div>
    `,
    onConfirm: (body) => ({
      date: body.querySelector('#rp-date').value,
      amount: parseFloat(body.querySelector('#rp-amount').value) || 0,
      note: body.querySelector('#rp-note').value.trim(),
      settle: body.querySelector('#rp-settle').checked
    })
  });
  if (!result || !result.amount) return;
  if (!loan.repayments) loan.repayments = [];
  loan.repayments.push({
    date: Parsers.formatDate(result.date),
    amount: result.amount,
    note: result.note
  });
  loan.repayments.sort((a,b) => a.date.localeCompare(b.date));

  const newRepaid = loan.repayments.reduce((s,r) => s+(r.amount||0), 0);
  if (result.settle || newRepaid >= (loan.principal||0)) {
    loan.status = 'settled';
  }
  save();
  renderLoans();
  renderAll();
  toast('已新增還款', 'ok');
}
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

  // 每月損益
  document.getElementById('monthlyYear').onchange = renderMonthly;
  document.getElementById('btnExportMonthly').onclick = exportMonthlyExcel;

  // 借款
  document.getElementById('btnAddLoan').onclick = addLoan;

  // 清空帳戶
  document.getElementById('btnClearAccount').onclick = async () => {
    const acc = getCurrentAccount();
    if (!acc) return;
    const ok = await confirmDialog('清空帳戶資料', `確定清空「${acc.name}」的所有未實現、投資明細、已實現資料？（帳戶與調整紀錄保留）`);
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

  // 沒有帳戶就建立一個預設的（名稱保留可改）
  if (!State.data.accounts.length) {
    State.data.accounts.push(emptyAccount('元大－主帳戶', '元大'));
    State.data.currentAccountId = State.data.accounts[0].id;
    save();
  }
  State.currentAccountId = State.data.currentAccountId || State.data.accounts[0].id;

  bindEvents();
  refreshAccountSelector();
  renderAll();
}

document.addEventListener('DOMContentLoaded', init);
