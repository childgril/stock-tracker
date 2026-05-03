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
    unrealized: [],
    unrealizedSnapshotDate: null,
    snapshots: [],
    trades: [],
    realized: [],
    adjustments: {},
    loans: [],
    marginCalls: [],
    dividends: {
      // 只用逐筆紀錄：{id, code, name, exDate, payDate, cash, stockShares, tax, note}
      entries: []
    }
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
      // 資料遷移：舊版本帳戶補上欄位
      if (data.accounts) {
        for (const acc of data.accounts) {
          if (!Array.isArray(acc.loans)) acc.loans = [];
          if (!Array.isArray(acc.marginCalls)) acc.marginCalls = [];
          if (!acc.dividends || typeof acc.dividends !== 'object') {
            acc.dividends = { entries: [] };
          }
          if (!Array.isArray(acc.dividends.entries)) acc.dividends.entries = [];
          // 舊版本的「累計模式」(aggregates) → 轉成沒日期的逐筆紀錄
          if (Array.isArray(acc.dividends.aggregates) && acc.dividends.aggregates.length > 0) {
            for (const a of acc.dividends.aggregates) {
              acc.dividends.entries.push({
                id: 'D-' + Date.now() + '-' + Math.random().toString(36).slice(2,6),
                code: a.code, name: a.name || '',
                exDate: '', payDate: '',
                cash: a.cashTotal || 0, stockShares: a.stockTotal || 0,
                tax: 0, note: '（從累計模式遷移）'
              });
            }
            delete acc.dividends.aggregates;
          }
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
  if (!sel) return;
  sel.innerHTML = '';
  for (const a of State.data.accounts) {
    const opt = document.createElement('option');
    opt.value = a.id;
    opt.textContent = a.name;
    sel.appendChild(opt);
  }
  sel.value = State.currentAccountId || '';

  document.getElementById('currentAccountName').textContent =
    getCurrentAccount() ? getCurrentAccount().name : '（無）';
  document.getElementById('accountTitle').textContent =
    getCurrentAccount() ? `帳戶明細：${getCurrentAccount().name}` : '帳戶明細';
}

// ---------- 帳戶順序調整對話框 ----------
async function reorderAccountsDialog() {
  if (!State.data.accounts.length) return toast('沒有帳戶可調整', 'err');

  const root = document.getElementById('modalRoot');
  // 製作清單 HTML
  const renderList = (items) => items.map((a, i) => `
    <li class="reorder-item" data-id="${a.id}" draggable="true">
      <span class="reorder-handle">⋮⋮</span>
      <span class="reorder-num">${i + 1}.</span>
      <span class="reorder-name">${a.name}</span>
      <span class="reorder-buttons">
        <button class="btn-mini" data-act="up" title="上移">▲</button>
        <button class="btn-mini" data-act="down" title="下移">▼</button>
      </span>
    </li>
  `).join('');

  // 工作副本（取消時不會改原資料）
  let working = State.data.accounts.slice();

  root.innerHTML = `
    <div class="modal-overlay">
      <div class="modal" style="min-width:400px;max-width:520px">
        <h3>調整帳戶順序</h3>
        <p class="hint" style="margin-top:0">拖曳左側 ⋮⋮ 圖示，或用 ▲▼ 按鈕調整順序</p>
        <ul class="reorder-list" id="reorderList">${renderList(working)}</ul>
        <div class="modal-actions">
          <button class="btn-mini" data-act="cancel">取消</button>
          <button class="btn-mini primary" data-act="ok">儲存</button>
        </div>
      </div>
    </div>
  `;

  const listEl = root.querySelector('#reorderList');

  const refreshList = () => {
    listEl.innerHTML = renderList(working);
    bindRow();
  };

  const move = (id, delta) => {
    const idx = working.findIndex(a => a.id === id);
    const newIdx = idx + delta;
    if (newIdx < 0 || newIdx >= working.length) return;
    const [m] = working.splice(idx, 1);
    working.splice(newIdx, 0, m);
    refreshList();
  };

  let dragId = null;
  const bindRow = () => {
    listEl.querySelectorAll('.reorder-item').forEach(li => {
      li.querySelector('[data-act="up"]').onclick = (e) => { e.stopPropagation(); move(li.dataset.id, -1); };
      li.querySelector('[data-act="down"]').onclick = (e) => { e.stopPropagation(); move(li.dataset.id, +1); };

      li.ondragstart = (e) => {
        dragId = li.dataset.id;
        li.classList.add('dragging');
        try { e.dataTransfer.setData('text/plain', dragId); } catch {}
        e.dataTransfer.effectAllowed = 'move';
      };
      li.ondragend = () => {
        li.classList.remove('dragging');
        listEl.querySelectorAll('.drag-over').forEach(x => x.classList.remove('drag-over'));
      };
      li.ondragover = (e) => {
        e.preventDefault();
        if (li.dataset.id !== dragId) {
          listEl.querySelectorAll('.drag-over').forEach(x => x.classList.remove('drag-over'));
          li.classList.add('drag-over');
        }
      };
      li.ondragleave = () => li.classList.remove('drag-over');
      li.ondrop = (e) => {
        e.preventDefault();
        li.classList.remove('drag-over');
        const targetId = li.dataset.id;
        if (!dragId || dragId === targetId) return;
        const srcIdx = working.findIndex(a => a.id === dragId);
        const tgtIdx = working.findIndex(a => a.id === targetId);
        const [m] = working.splice(srcIdx, 1);
        const newTgtIdx = working.findIndex(a => a.id === targetId);
        working.splice(newTgtIdx, 0, m);
        refreshList();
      };
    });
  };
  bindRow();

  root.querySelector('[data-act="cancel"]').onclick = () => { root.innerHTML = ''; };
  root.querySelector('[data-act="ok"]').onclick = () => {
    State.data.accounts = working;
    save();
    refreshAccountSelector();
    renderAll();
    root.innerHTML = '';
    toast('已儲存新順序', 'ok');
  };
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
        r.refund = saved.refund || 0;
        r.note = saved.note || '';
        if (saved.adjust || saved.refund || saved.note) restored++;
      } else {
        r.adjust = 0;
        r.refund = 0;
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

// 已實現損益的「總調整 = 盈虧調整 + 回補資金」
function adjustTotal(r) {
  return (r.adjust || 0) + (r.refund || 0);
}
// 實際盈虧 = 原始盈虧 + 總調整
function actualPL(r) {
  return (r.pl || 0) + adjustTotal(r);
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
    '利息','融券手續費','盈虧調整','回補資金','調整合計','備註',
    '盈虧','實際盈虧'
  ];
  const rows = acc.realized.map(r => [
    r.code, r.name, r.sellCategory,
    r.sellDate, r.sellPrice, r.buyDate, r.buyPrice,
    r.qty, r.cost, r.fee, r.tax,
    r.interest || 0, r.shortFee || 0,
    r.adjust || 0, r.refund || 0, adjustTotal(r), r.note || '',
    r.pl, actualPL(r)
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
  renderMarginCalls();
  renderDividends();
  renderPeriodInfo();
}

// ---------- 交易期間（每頁右上角）----------
function getDateRange(accounts) {
  // 從一個或多個帳戶的所有資料中找出最早與最晚日期
  let min = null, max = null;
  const consider = (d) => {
    if (!d) return;
    const s = String(d).trim();
    if (!/^\d{4}\/\d{1,2}\/\d{1,2}/.test(s)) return;
    const norm = s.slice(0, 10);
    if (!min || norm < min) min = norm;
    if (!max || norm > max) max = norm;
  };

  for (const acc of accounts) {
    if (!acc) continue;
    for (const t of (acc.trades || [])) consider(t.date);
    for (const r of (acc.realized || [])) {
      consider(r.sellDate);
      consider(r.buyDate);
    }
    for (const s of (acc.snapshots || [])) consider(s.date);
    if (acc.unrealizedSnapshotDate) consider(acc.unrealizedSnapshotDate);
  }
  return { min, max };
}

function formatPeriodHTML(range) {
  if (!range.min || !range.max) {
    return `<span class="label">交易期間</span><span class="range" style="color:var(--text-muted);font-weight:400">尚無資料</span>`;
  }
  // 算天數
  let daysHtml = '';
  try {
    const d1 = new Date(range.min.replace(/\//g, '-'));
    const d2 = new Date(range.max.replace(/\//g, '-'));
    if (!isNaN(d1) && !isNaN(d2)) {
      const days = Math.round((d2 - d1) / 86400000) + 1;
      daysHtml = `<span class="days">共 ${days} 天</span>`;
    }
  } catch (e) {}
  return `
    <span class="label">交易期間</span>
    <span class="range">${range.min} ～ ${range.max}</span>
    ${daysHtml}
  `;
}

function renderPeriodInfo() {
  // 各帳戶頁面用「目前帳戶」
  const acc = getCurrentAccount();
  const accRange = acc ? getDateRange([acc]) : { min: null, max: null };
  const accHtml = formatPeriodHTML(accRange);

  // 總覽用「全部帳戶」+ 可展開的各帳戶清單
  const accounts = State.data.accounts || [];
  const allRange = getDateRange(accounts);
  let allHtml = formatPeriodHTML(allRange);

  // 多於 1 個帳戶才顯示展開按鈕
  if (accounts.length > 1) {
    const isOpen = State._periodExpanded === true;
    allHtml += `
      <button class="period-toggle" id="periodToggleBtn" title="展開/收合各帳戶期間">
        ${isOpen ? '▲ 收合各帳戶' : '▼ 各帳戶期間'}
      </button>
    `;
    if (isOpen) {
      // 各帳戶期間卡片
      const items = accounts.map(a => {
        const r = getDateRange([a]);
        let days = '';
        if (r.min && r.max) {
          try {
            const d1 = new Date(r.min.replace(/\//g, '-'));
            const d2 = new Date(r.max.replace(/\//g, '-'));
            if (!isNaN(d1) && !isNaN(d2)) {
              days = `${Math.round((d2 - d1) / 86400000) + 1} 天`;
            }
          } catch (e) {}
        }
        const rangeText = (r.min && r.max) ? `${r.min} ～ ${r.max}` : '尚無資料';
        return `
          <li class="period-item">
            <div class="acc-name">${a.name || '未命名'}</div>
            <div class="acc-range">${rangeText}</div>
            ${days ? `<div class="acc-days">${days}</div>` : ''}
          </li>
        `;
      }).join('');
      allHtml += `
        <div class="period-detail-popup">
          <div class="period-detail-title">各帳戶交易期間</div>
          <ul class="period-detail-list">${items}</ul>
        </div>
      `;
    }
  }

  // 各 section 對應
  const map = {
    'period-overview': allHtml,
    'period-account': accHtml,
    'period-realized': accHtml,
    'period-unrealized': accHtml,
    'period-trades': accHtml,
    'period-loans': accHtml,
    'period-margincall': accHtml,
    'period-dividends': accHtml,
    'period-import': accHtml
  };
  for (const [id, html] of Object.entries(map)) {
    const el = document.getElementById(id);
    if (el) el.innerHTML = html;
  }

  // 綁定展開按鈕
  const toggleBtn = document.getElementById('periodToggleBtn');
  if (toggleBtn) {
    toggleBtn.onclick = (e) => {
      e.stopPropagation();
      State._periodExpanded = !State._periodExpanded;
      renderPeriodInfo();
    };
  }
}

// ---------- 帳戶聚合 ----------
function aggregateAccount(acc) {
  const totalMarket = acc.unrealized.reduce((s,x) => s+x.marketValue, 0);
  const totalCost   = acc.unrealized.reduce((s,x) => s+x.cost, 0);
  const unrealizedPL= acc.unrealized.reduce((s,x) => s+x.pl, 0);
  const totalInterest = acc.realized.reduce((s,r) => s+(r.interest||0), 0);
  const totalShortFee = acc.realized.reduce((s,r) => s+(r.shortFee||0), 0);
  const realizedPLRaw = acc.realized.reduce((s,r) => s + actualPL(r), 0);

  // 借款利息累計
  const loans = acc.loans || [];
  const loanInterestPaid = loans.reduce((s, l) =>
    s + (l.interestPayments || []).reduce((ss, p) => ss + (p.amount || 0), 0), 0);
  const loanPrincipal = loans.reduce((s, l) => s + (l.principal || 0), 0);
  const loanRepaid = loans.reduce((s, l) =>
    s + (l.repayments || []).reduce((ss, p) => ss + (p.amount || 0), 0), 0);
  const loanRemaining = loanPrincipal - loanRepaid;

  // 融資回補
  const marginCalls = acc.marginCalls || [];
  const mcTotal = marginCalls.reduce((s, m) => s + (m.amount || 0), 0);
  const mcPayout = marginCalls.reduce((s, m) =>
    s + (m.payouts || []).reduce((ss, p) => ss + (p.amount || 0), 0), 0);
  const mcRemaining = mcTotal - mcPayout;

  // 股利
  const div = acc.dividends || { aggregates: [], entries: [] };
  const dividendCash = computeDividendTotalCash(acc);
  const dividendStock = computeDividendTotalStock(acc);

  // 已實現損益要扣借款利息（你要求的：借款利息計入實際損益）
  // 並加入現金股利（股利是這個帳戶股票產生的）
  const realizedPL = realizedPLRaw - loanInterestPaid + dividendCash;

  return {
    totalMarket, totalCost, unrealizedPL,
    totalInterest, totalShortFee,
    realizedPL,
    realizedPLRaw,
    loanInterestPaid,
    loanPrincipal, loanRepaid, loanRemaining,
    mcTotal, mcPayout, mcRemaining,
    dividendCash, dividendStock
  };
}

// ---------- 股利統計輔助 ----------
// 全部都從逐筆紀錄加總而來
function computeDividendByCode(acc) {
  const div = acc.dividends || { entries: [] };
  const map = new Map();
  for (const e of div.entries) {
    if (!map.has(e.code)) {
      map.set(e.code, { code: e.code, name: e.name || '', cash: 0, stock: 0, tax: 0, count: 0, entries: [] });
    }
    const s = map.get(e.code);
    if (!s.name && e.name) s.name = e.name;
    s.cash += (e.cash || 0);
    s.stock += (e.stockShares || 0);
    s.tax += (e.tax || 0);
    s.count++;
    s.entries.push(e);
  }
  return [...map.values()];
}

function computeDividendTotalCash(acc) {
  return computeDividendByCode(acc).reduce((s, x) => s + x.cash, 0);
}
function computeDividendTotalStock(acc) {
  return computeDividendByCode(acc).reduce((s, x) => s + x.stock, 0);
}

function setVal(id, val, withClass = false) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = typeof val === 'number' ? fmt(val, { sign: withClass }) : val;
  if (withClass) el.className = 'value ' + plClass(val);
}

// ---------- 總覽 ----------
function renderOverview() {
  let M=0, C=0, U=0, R=0, I=0, S=0, LI=0, LB=0, MCR=0, DC=0;
  const perAccount = [];
  for (const a of State.data.accounts) {
    const g = aggregateAccount(a);
    M += g.totalMarket; C += g.totalCost; U += g.unrealizedPL;
    R += g.realizedPL; I += g.totalInterest; S += g.totalShortFee;
    LI += g.loanInterestPaid; LB += g.loanRemaining;
    MCR += g.mcRemaining;
    DC += g.dividendCash;
    perAccount.push({ name: a.name, ...g });
  }
  setVal('ovTotalMarket', M);
  setVal('ovTotalCost', C);
  setVal('ovUnrealizedPL', U, true);
  setVal('ovRealizedPL', R, true);
  setVal('ovTotalInterest', I);
  setVal('ovTotalShortFee', S);
  setVal('ovMarginCallRemaining', MCR);
  setVal('ovTotalDividend', DC);

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
    ['acMarket','acCost','acUnrealizedPL','acRealizedPL','acInterest','acShortFee','acLoanInterest','acLoanBalance','acMarginCallRemaining','acDividendCash']
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
  setVal('acMarginCallRemaining', g.mcRemaining);
  setVal('acDividendCash', g.dividendCash);

  const lines = [];
  lines.push(`<p class="hint">未實現損益：<strong>${acc.unrealized.length}</strong> 檔（快照日 ${acc.unrealizedSnapshotDate || '—'}）　|　投資明細：<strong>${acc.trades.length}</strong> 筆　|　已實現損益：<strong>${acc.realized.length}</strong> 筆　|　歷史快照：<strong>${(acc.snapshots||[]).length}</strong> 筆　|　借款：<strong>${(acc.loans||[]).length}</strong> 筆</p>`);
  status.innerHTML = lines.join('');

  renderMonthCompare();
  renderMonthly();
  renderDayTrades();
  renderStockAnalysis();
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
    m.adjust += adjustTotal(r);
    m.realizedItems.push(r);
  }

  // 投資明細：以成交日為準
  for (const t of (acc.trades || [])) {
    const k = getMonthKey(t.date);
    if (!k) continue;
    const m = ensure(k);
    const isConv = (t.category === '資轉現' || t.action === '資轉現');
    if (isConv) {
      // 資轉現：只計入結清利息，不算進交易筆數和成交金額
      m.interest += (t.marginInterest || 0);
      m.tradeItems.push(t);
    } else {
      m.tradeCount++;
      m.tradeAmount += (t.amount || 0);
      m.tradeItems.push(t);
    }
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

  // 融資回補：以回補日為準（純參考，不計入實際損益）
  for (const mc of (acc.marginCalls || [])) {
    const k = getMonthKey(mc.date);
    if (!k) continue;
    const m = ensure(k);
    m.marginCall = (m.marginCall || 0) + (mc.amount || 0);
  }

  // 股利：用「逐筆紀錄的發放日」分月
  // 累計模式沒有日期 → 統一歸到「無日期」（在實際損益計算中還是會總額計入帳戶）
  // 這裡只把「有日期」的逐筆紀錄分到對應月
  for (const e of (acc.dividends?.entries || [])) {
    const k = getMonthKey(e.payDate || e.exDate);
    if (!k) continue;
    const m = ensure(k);
    m.dividendCash = (m.dividendCash || 0) + (e.cash || 0);
  }

  // 計算實際損益（已實現 + 調整 - 借款利息 + 現金股利）
  for (const m of months.values()) {
    m.actual = m.realizedPL + m.adjust - m.loanInterest + (m.dividendCash || 0);
    if (!m.marginCall) m.marginCall = 0;
    if (!m.dividendCash) m.dividendCash = 0;
  }

  return [...months.values()].sort((a, b) => b.key.localeCompare(a.key));
}

// ============================================================
// 月份比較分析
// ============================================================
// 把 buildMonthlyData 的結果加上當沖損益（不在月資料裡）
function buildMonthlyDataWithDayTrade(acc) {
  const months = buildMonthlyData(acc);
  // 當沖按月加總
  const dtByMonth = new Map();
  for (const dt of analyzeDayTrades(acc)) {
    const mKey = (dt.date || '').slice(0, 7).replace('/', '-');
    if (!mKey) continue;
    if (!dtByMonth.has(mKey)) dtByMonth.set(mKey, 0);
    dtByMonth.set(mKey, dtByMonth.get(mKey) + (dt.netPL || 0));
  }
  for (const m of months) {
    m.dayTradePL = dtByMonth.get(m.key) || 0;
  }
  return months;
}

// 取得指標的顯示名稱
function cmpMetricLabel(metric) {
  const map = {
    actual: '實際損益',
    realizedPL: '已實現損益',
    interest: '融資利息',
    shortFee: '融券手續費',
    dividendCash: '現金股利',
    adjust: '調整金額',
    dayTradePL: '當沖損益',
    tradeAmount: '總成交金額'
  };
  return map[metric] || metric;
}

// 把月份 key 'YYYY-MM' 變成顯示用的 '2024 年 6 月'
function fmtMonthLabel(key) {
  const m = (key || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return key;
  return `${m[1]} 年 ${parseInt(m[2])} 月`;
}

// 計算變化百分比
function calcChange(current, base) {
  const diff = current - base;
  let pct = null;
  if (base !== 0) {
    pct = (diff / Math.abs(base)) * 100;
  }
  return { diff, pct };
}

// 渲染變化箭頭儲存格
function renderChangeCell(diff, pct, options = {}) {
  const lowerIsBetter = options.lowerIsBetter || false;
  const isUp = diff > 0;
  const isFlat = diff === 0;
  // 預設情境：上升=綠色（好），下降=紅色（壞）
  // 若指標是「越低越好」（如利息、手續費），則反轉
  let cls;
  if (isFlat) cls = 'flat';
  else if ((isUp && !lowerIsBetter) || (!isUp && lowerIsBetter)) cls = 'up';
  else cls = 'down';

  const icon = isFlat ? '—' : (isUp ? '▲' : '▼');
  const pctStr = pct === null ? '—' : `${pct > 0 ? '+' : ''}${pct.toFixed(1)}%`;
  const diffStr = fmt(diff, { sign: true });

  // 視覺化條（依 |pct| 寬度，最大 100%）
  let bar = '';
  if (pct !== null && !isFlat) {
    const w = Math.min(50, Math.abs(pct) / 2); // pct 50% 滿條
    const fillStyle = isUp
      ? `width:${w}%; left:50%`
      : `width:${w}%; right:50%; left:auto`;
    bar = `<span class="cmp-bar"><span class="fill ${cls}" style="${fillStyle}"></span></span>`;
  }

  return `
    <div class="cmp-arrow ${cls}">
      <span class="icon">${icon}</span>
      <span>${diffStr}</span>
      <span style="font-size:11px;opacity:0.85">(${pctStr})</span>
    </div>
  `;
}

// 取資料表中某月的指標數值
function getMetricValue(monthData, metric) {
  if (!monthData) return 0;
  return monthData[metric] || 0;
}

// 三種比較模式產生「比較對」清單：[{label, current: monthData, base: monthData, baseLabel}]
function buildComparePairs(months, mode, options = {}) {
  const pairs = [];
  if (!months.length) return pairs;

  // months 已從新到舊排序
  const byKey = new Map(months.map(m => [m.key, m]));

  if (mode === 'yearly') {
    // 年度跨月：每個月找去年同月
    for (const m of months) {
      const [y, mo] = m.key.split('-').map(Number);
      const prevKey = `${y - 1}-${String(mo).padStart(2, '0')}`;
      const base = byKey.get(prevKey);
      if (base) {
        pairs.push({
          currentLabel: fmtMonthLabel(m.key),
          baseLabel: fmtMonthLabel(prevKey),
          current: m,
          base
        });
      }
    }
  } else if (mode === 'consecutive') {
    // 連續月：每個月跟前一個月（按時間連續，跨年也算）
    for (let i = 0; i < months.length - 1; i++) {
      const cur = months[i];
      const prev = months[i + 1];  // months 是新到舊，所以 i+1 是更舊的
      pairs.push({
        currentLabel: fmtMonthLabel(cur.key),
        baseLabel: fmtMonthLabel(prev.key),
        current: cur,
        base: prev
      });
    }
  } else if (mode === 'custom') {
    // 自選：options.monthA / options.monthB
    const a = byKey.get(options.monthA);
    const b = byKey.get(options.monthB);
    if (a && b) {
      pairs.push({
        currentLabel: fmtMonthLabel(options.monthA),
        baseLabel: fmtMonthLabel(options.monthB),
        current: a,
        base: b
      });
    }
  }

  return pairs;
}

// 哪些指標「越低越好」？
const LOWER_IS_BETTER_METRICS = new Set(['interest', 'shortFee']);

function renderMonthCompare() {
  const acc = getCurrentAccount();
  const wrap = document.getElementById('monthCompareArea');
  if (!wrap) return;

  if (!acc) {
    wrap.innerHTML = '<div class="cmp-empty">尚未選擇帳戶</div>';
    return;
  }

  const mode = document.getElementById('cmpMode')?.value || 'yearly';
  const metric = document.getElementById('cmpMetric')?.value || 'actual';
  const monthASel = document.getElementById('cmpMonthA');
  const monthBSel = document.getElementById('cmpMonthB');

  const months = buildMonthlyDataWithDayTrade(acc);

  // 自選模式：填月份下拉並顯示
  if (mode === 'custom') {
    if (monthASel) monthASel.style.display = '';
    if (monthBSel) monthBSel.style.display = '';
    // 重建選項（保留現選）
    const cur = monthASel?.value || '';
    const cur2 = monthBSel?.value || '';
    if (monthASel) {
      monthASel.innerHTML = '<option value="">A 月份…</option>' +
        months.map(m => `<option value="${m.key}" ${m.key===cur?'selected':''}>${fmtMonthLabel(m.key)}</option>`).join('');
    }
    if (monthBSel) {
      monthBSel.innerHTML = '<option value="">B 月份…</option>' +
        months.map(m => `<option value="${m.key}" ${m.key===cur2?'selected':''}>${fmtMonthLabel(m.key)}</option>`).join('');
    }
  } else {
    if (monthASel) monthASel.style.display = 'none';
    if (monthBSel) monthBSel.style.display = 'none';
  }

  if (!months.length) {
    wrap.innerHTML = '<div class="cmp-empty">尚無資料可比較。請先匯入已實現損益或投資明細</div>';
    return;
  }

  const pairs = buildComparePairs(months, mode, {
    monthA: monthASel?.value || '',
    monthB: monthBSel?.value || ''
  });

  if (!pairs.length) {
    if (mode === 'custom') {
      wrap.innerHTML = '<div class="cmp-empty">請從上方選擇兩個月份來比較</div>';
    } else if (mode === 'yearly') {
      wrap.innerHTML = '<div class="cmp-empty">沒有「同月份不同年」的可比資料（至少要有兩年的同月份資料）</div>';
    } else {
      wrap.innerHTML = '<div class="cmp-empty">至少要有兩個月的資料才能做連續月比較</div>';
    }
    return;
  }

  const lowerIsBetter = LOWER_IS_BETTER_METRICS.has(metric);
  const metricName = cmpMetricLabel(metric);

  // 建表
  const headerLabel = mode === 'yearly' ? '本月 vs 去年同月'
                    : mode === 'consecutive' ? '本月 vs 上月'
                    : '比較';

  let html = `
    <div class="hint" style="margin:-4px 0 8px;font-size:12px">
      📈 ${metricName}　${lowerIsBetter ? '（越低越好，下降以綠色顯示）' : '（越高越好，上升以綠色顯示）'}
    </div>
    <table class="cmp-table">
      <thead><tr>
        <th>${mode === 'custom' ? 'A 月份' : '本月'}</th>
        <th>${metricName}（A）</th>
        <th>${mode === 'custom' ? 'B 月份' : (mode === 'yearly' ? '去年同月' : '上月')}</th>
        <th>${metricName}（B）</th>
        <th>變化</th>
      </tr></thead>
      <tbody>
  `;

  for (const p of pairs) {
    const curVal = getMetricValue(p.current, metric);
    const baseVal = getMetricValue(p.base, metric);
    const { diff, pct } = calcChange(curVal, baseVal);
    const curCls = curVal > 0 ? 'pos' : (curVal < 0 ? 'neg' : '');
    const baseCls = baseVal > 0 ? 'pos' : (baseVal < 0 ? 'neg' : '');

    html += `
      <tr>
        <td><div class="cmp-period">${p.currentLabel}</div></td>
        <td class="${curCls}">${fmt(curVal, { sign: true })}</td>
        <td><div class="cmp-period">${p.baseLabel}</div></td>
        <td class="${baseCls}">${fmt(baseVal, { sign: true })}</td>
        <td>${renderChangeCell(diff, pct, { lowerIsBetter })}</td>
      </tr>
    `;
  }

  // 摘要列：總和（自選模式不顯示）
  if (pairs.length > 1) {
    let sumCur = 0, sumBase = 0;
    for (const p of pairs) {
      sumCur += getMetricValue(p.current, metric);
      sumBase += getMetricValue(p.base, metric);
    }
    const { diff, pct } = calcChange(sumCur, sumBase);
    const curCls = sumCur > 0 ? 'pos' : (sumCur < 0 ? 'neg' : '');
    const baseCls = sumBase > 0 ? 'pos' : (sumBase < 0 ? 'neg' : '');
    html += `
      <tr style="background:var(--surface-2);font-weight:600">
        <td>合計（${pairs.length} 組）</td>
        <td class="${curCls}">${fmt(sumCur, { sign: true })}</td>
        <td>—</td>
        <td class="${baseCls}">${fmt(sumBase, { sign: true })}</td>
        <td>${renderChangeCell(diff, pct, { lowerIsBetter })}</td>
      </tr>
    `;
  }

  html += `</tbody></table>`;
  wrap.innerHTML = html;
}

const _expandedMonths = new Set();

function renderMonthly() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#monthlyTable tbody');
  const yearSel = document.getElementById('monthlyYear');
  if (!acc) return;

  const data = buildMonthlyData(acc);
  // 算當沖（按月匯總）
  const dayTrades = analyzeDayTrades(acc);
  const dtByMonth = aggregateDayTradesByMonth(dayTrades);

  // 年份下拉
  const years = [...new Set(data.map(m => m.key.slice(0,4)))].sort().reverse();
  const currentYear = yearSel.value;
  yearSel.innerHTML = '<option value="">所有年份</option>' +
    years.map(y => `<option value="${y}" ${y===currentYear?'selected':''}>${y}</option>`).join('');

  const filtered = currentYear ? data.filter(m => m.key.startsWith(currentYear)) : data;

  if (!filtered.length) {
    tb.innerHTML = '<tr><td colspan="16" class="empty-state">尚無資料（請先匯入已實現損益或投資明細）</td></tr>';
    return;
  }

  const rows = [];
  for (const m of filtered) {
    const expanded = _expandedMonths.has(m.key);
    const [year, month] = m.key.split('-');
    const dt = dtByMonth.get(m.key);
    rows.push(`
      <tr class="month-row" data-month="${m.key}">
        <td><span class="month-toggle ${expanded?'expanded':''}">▶</span></td>
        <td><span class="year-tag">${year}</span><strong>${parseInt(month)}月</strong></td>
        <td class="${plClass(m.realizedPL)}">${fmt(m.realizedPL,{sign:true})}</td>
        <td>${fmt(m.interest)}</td>
        <td>${fmt(m.shortFee)}</td>
        <td class="${m.loanInterest?'neg':''}">${m.loanInterest ? '-'+fmt(m.loanInterest) : '—'}</td>
        <td>${m.marginCall ? fmt(m.marginCall) : '—'}</td>
        <td class="${m.dividendCash?'pos':''}">${m.dividendCash ? '+'+fmt(m.dividendCash) : '—'}</td>
        <td class="${plClass(m.adjust)}">${m.adjust?fmt(m.adjust,{sign:true}):'—'}</td>
        <td class="hl ${plClass(m.actual)}">${fmt(m.actual,{sign:true})}</td>
        <td>${m.tradeCount}</td>
        <td>${fmt(m.tradeAmount)}</td>
        <td class="hl">${dt ? dt.count : '—'}</td>
        <td class="hl ${dt?plClass(dt.netPL):''}">${dt ? fmt(dt.netPL,{sign:true}) : '—'}</td>
        <td class="hl ${dt && dt.winRate>=50 ? 'pos' : (dt && dt.winRate<50 ? 'neg' : '')}">${dt ? dt.winRate.toFixed(1)+'%' : '—'}</td>
        <td class="hl ${dt?plClass(dt.rate):''}">${dt ? fmtPct(dt.rate) : '—'}</td>
      </tr>
    `);
    if (expanded) {
      rows.push(`<tr class="month-detail-row"><td colspan="16">${renderMonthDetail(m)}</td></tr>`);
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
      const adjT = adjustTotal(r);
      const actual = actualPL(r);
      html.push(`<tr>
        <td>${r.code}</td><td>${r.name||''}</td><td>${r.sellCategory||''}</td>
        <td>${r.sellDate}</td><td>${r.buyDate||'—'}</td>
        <td>${fmt(r.qty)}</td>
        <td>${fmt(r.sellPrice,{decimals:2})}</td>
        <td>${fmt(r.buyPrice,{decimals:2})}</td>
        <td class="${plClass(r.pl)}">${fmt(r.pl,{sign:true})}</td>
        <td>${fmt(r.interest||0)}</td>
        <td>${fmt(r.shortFee||0)}</td>
        <td class="${plClass(adjT)}">${adjT?fmt(adjT,{sign:true}):'—'}</td>
        <td class="${plClass(actual)}"><strong>${fmt(actual,{sign:true})}</strong></td>
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

  // 該月融資回補
  const monthMCs = (getCurrentAccount().marginCalls || []).filter(mc => getMonthKey(mc.date) === m.key);
  if (monthMCs.length > 0) {
    html.push('<div class="nested">');
    html.push('<h4 style="padding:10px 14px 0">⚠️ 該月融資回補（' + monthMCs.length + ' 筆）</h4>');
    html.push('<table class="loan-subtable"><thead><tr>');
    html.push('<th>編號</th><th>日期</th><th>回補金額</th><th>原因</th><th>已領回</th><th>未領回</th><th>狀態</th>');
    html.push('</tr></thead><tbody>');
    for (const mc of monthMCs) {
      const paid = (mc.payouts || []).reduce((s, p) => s + (p.amount || 0), 0);
      const remaining = (mc.amount || 0) - paid;
      const status = remaining <= 0 ? '已全額領回' : (paid > 0 ? '部分領回' : '未領回');
      html.push(`<tr>
        <td>${mc.id}</td>
        <td>${mc.date}</td>
        <td>${fmt(mc.amount)}</td>
        <td>${mc.reason||''}</td>
        <td>${fmt(paid)}</td>
        <td class="${remaining>0?'neg':''}">${fmt(remaining)}</td>
        <td>${status}</td>
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

  const dayTrades = analyzeDayTrades(acc);
  const dtByMonth = aggregateDayTradesByMonth(dayTrades);

  const headers = ['月份','已實現損益','融資利息','融券手續費','借款利息','融資回補','現金股利','調整金額','實際損益','交易筆數','總成交金額',
                   '當沖筆數','當沖損益','當沖勝率(%)','當沖報酬率(%)'];
  const rows = data.map(m => {
    const dt = dtByMonth.get(m.key);
    return [
      m.key, m.realizedPL, m.interest, m.shortFee,
      m.loanInterest, m.marginCall || 0, m.dividendCash || 0,
      m.adjust, m.actual,
      m.tradeCount, m.tradeAmount,
      dt ? dt.count : 0,
      dt ? dt.netPL : 0,
      dt ? +dt.winRate.toFixed(2) : 0,
      dt ? +dt.rate.toFixed(2) : 0
    ];
  });
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '每月損益');
  const ts = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `每月損益_${acc.name}_${ts}.xlsx`);
  toast('已匯出 Excel', 'ok');
}

// ---------- 已實現損益 ----------
// ============================================================
// 已實現 ↔ 投資明細 配對提醒與手動配對
// ============================================================

// 找出未配對的已實現、沒人引用的賣出投資明細
function findMatchProblems(acc) {
  const unmatched = (acc.realized || []).filter(r => r._unmatched);

  // 哪些 trade 已經被引用？
  const usedTradeKeys = new Set();
  for (const r of (acc.realized || [])) {
    if (r._matchedTradeKey) usedTradeKeys.add(r._matchedTradeKey);
    if (r._manualMatchTradeKey) usedTradeKeys.add(r._manualMatchTradeKey);
  }

  // 投資明細中「賣出 + 融資/融券」但沒被任何已實現引用的
  const orphanTrades = (acc.trades || []).filter(t => {
    if (t.action !== '賣') return false;
    if (t.category !== '融資' && t.category !== '融券') return false;
    return !usedTradeKeys.has(Parsers.makeTradeKey(t));
  });

  return { unmatched, orphanTrades };
}

function renderMatchAlerts(acc) {
  const wrap = document.getElementById('matchAlerts');
  if (!wrap) return;
  const { unmatched, orphanTrades } = findMatchProblems(acc);

  if (!unmatched.length && !orphanTrades.length) {
    wrap.innerHTML = '';
    return;
  }

  let html = '<div class="match-alert">';

  if (unmatched.length > 0) {
    html += `⚠️ 有 <strong>${unmatched.length}</strong> 筆已實現損益（融資/融券）找不到對應的投資明細，
             利息或融券手續費可能不準確。請點該列右側「🔗 配對」手動指定。`;
  }

  if (orphanTrades.length > 0) {
    html += unmatched.length > 0 ? '<br>' : '';
    html += `❓ 投資明細中有 <strong>${orphanTrades.length}</strong> 筆「賣出融資/融券」沒有對應的已實現損益（可能是已實現損益沒匯入完整）。`;
    html += `<details><summary>展開查看這些投資明細</summary>
      <table class="alert-table">
        <thead><tr><th>日期</th><th>代號</th><th>名稱</th><th>類別</th><th>數量</th><th>單價</th><th>利息</th><th>融券手續費</th></tr></thead>
        <tbody>`;
    for (const t of orphanTrades) {
      html += `<tr>
        <td>${t.date}</td>
        <td>${t.code}</td>
        <td>${t.name||''}</td>
        <td>${t.category}</td>
        <td>${fmt(t.qty)}</td>
        <td>${fmt(t.price,{decimals:2})}</td>
        <td>${fmt(t.marginInterest||0)}</td>
        <td>${fmt(t.shortFee||0)}</td>
      </tr>`;
    }
    html += '</tbody></table></details>';
  }

  html += '</div>';
  wrap.innerHTML = html;
}

// 手動配對對話框
async function showManualMatchDialog(realized) {
  const acc = getCurrentAccount();
  if (!acc) return;

  // 候選清單：所有「賣出融資/融券」的投資明細
  const allCandidates = (acc.trades || []).filter(t =>
    t.action === '賣' && (t.category === '融資' || t.category === '融券')
  );
  const sameCodeCategory = allCandidates.filter(t =>
    t.code === realized.code && t.category === realized.sellCategory
  );

  if (allCandidates.length === 0) {
    return toast('投資明細裡沒有任何「賣出融資/融券」紀錄可配對', 'err');
  }

  // 哪些已被其他已實現「手動」引用？（自動配對的不算「佔用」，因為一旦你手動配對會把它解開）
  const manuallyUsedKeys = new Set();
  for (const r of acc.realized) {
    if (r === realized) continue;
    if (r._manualMatchTradeKey) manuallyUsedKeys.add(r._manualMatchTradeKey);
  }

  const buildListHTML = (showAll) => {
    const list = showAll ? allCandidates : sameCodeCategory;
    if (!list.length) return '<div class="empty-state" style="padding:20px">沒有同代號同類別的賣出紀錄。請勾選「顯示全部」</div>';
    return `
      <table class="alert-table">
        <thead><tr><th>選</th><th>日期</th><th>代號</th><th>名稱</th><th>類別</th><th>數量</th><th>單價</th><th>利息</th><th>融券費</th><th>狀態</th></tr></thead>
        <tbody>
        ${list.map(t => {
          const k = Parsers.makeTradeKey(t);
          const used = manuallyUsedKeys.has(k);
          const isCurrent = (realized._manualMatchTradeKey === k);
          const disabled = used && !isCurrent ? 'disabled' : '';
          const stateColor = used && !isCurrent ? 'color:#aaa' : isCurrent ? 'color:#22c55e;font-weight:600' : '';
          const stateText = used && !isCurrent ? '其他已實現手動引用' : isCurrent ? '目前配對' : '可選';
          return `
            <tr style="${used && !isCurrent ? 'opacity:0.4' : ''}">
              <td><input type="radio" name="match-pick" value="${k}" ${isCurrent?'checked':''} ${disabled}></td>
              <td>${t.date}</td>
              <td>${t.code}</td>
              <td>${t.name||''}</td>
              <td>${t.category}</td>
              <td><strong>${fmt(t.qty)}</strong></td>
              <td>${fmt(t.price,{decimals:2})}</td>
              <td>${fmt(t.marginInterest||0)}</td>
              <td>${fmt(t.shortFee||0)}</td>
              <td style="${stateColor}">${stateText}</td>
            </tr>
          `;
        }).join('')}
        </tbody>
      </table>
    `;
  };

  const modalPromise = showModal({
    title: `手動配對 - ${realized.code} ${realized.name} ${realized.sellDate} ${fmt(realized.qty)}股`,
    html: `
      <div style="max-width:820px;min-width:720px">
        <p class="hint" style="margin-top:0">
          選一筆投資明細作為這筆已實現損益的對應交易。<br>
          ⚙️ 系統會自動把「同代號 + 同賣出日 + 同類別」的其他未配對已實現一起併入這個 trade，
          利息與融券手續費按沖銷成本比例分攤。
        </p>
        <label style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
          <input type="checkbox" id="mm-showall"> 顯示全部「賣出融資/融券」紀錄（不限同代號同類別）
        </label>
        <div id="mm-list">${buildListHTML(false)}</div>
      </div>
    `,
    onConfirm: (body) => {
      const picked = body.querySelector('input[name="match-pick"]:checked');
      return picked ? picked.value : null;
    }
  });

  setTimeout(() => {
    const cb = document.getElementById('mm-showall');
    const listWrap = document.getElementById('mm-list');
    if (cb && listWrap) {
      cb.onchange = () => {
        listWrap.innerHTML = buildListHTML(cb.checked);
      };
    }
  }, 0);

  const result = await modalPromise;
  if (!result) return;

  // 找出對應的 trade
  const trade = acc.trades.find(t => Parsers.makeTradeKey(t) === result);
  if (!trade) return toast('找不到對應的投資明細', 'err');

  // 找「同代號 + 同賣出日 + 同類別」的所有未配對已實現（包含這筆）
  const sameGroup = acc.realized.filter(r =>
    r.code === realized.code &&
    r.sellDate === realized.sellDate &&
    r.sellCategory === realized.sellCategory &&
    (r === realized || (
      // 必須是「未手動配對到別的 trade」的（包含未配對 + 自動配對）
      !r._manualMatchTradeKey ||
      r._manualMatchTradeKey === result // 已配對到同一個的也納入
    ))
  );

  const realizedQty = sameGroup.reduce((s, r) => s + (r.qty || 0), 0);
  const tradeQty = trade.qty || 0;

  // 數量比對
  let proceed = true;
  let willGroup = sameGroup;

  if (realizedQty === tradeQty) {
    // 完美 → 自動成組
    if (sameGroup.length > 1) {
      const ok = await confirmDialog(
        '自動群組配對',
        `偵測到「${realized.code} ${realized.sellDate} ${realized.sellCategory}」共有 ${sameGroup.length} 筆已實現（合計 ${fmt(realizedQty)} 股）<br>
         剛好等於這筆 trade 的 ${fmt(tradeQty)} 股。<br><br>
         系統會把這 ${sameGroup.length} 筆都配對到同一個 trade，利息 ${fmt(trade.marginInterest||0)} 按比例分攤。<br><br>
         繼續嗎？`
      );
      if (!ok) {
        // 使用者取消群組 → 改成只配當前這筆
        const ok2 = await confirmDialog(
          '只配對這 1 筆？',
          `這樣只有 ${realized.code} ${realized.sellDate} ${fmt(realized.qty)} 股會配對到此 trade，<br>
           其他 ${sameGroup.length - 1} 筆仍未配對。利息 ${fmt(trade.marginInterest||0)} 全部給這筆。<br><br>
           真的要嗎？`
        );
        if (!ok2) return;
        willGroup = [realized];
      }
    }
  } else if (realizedQty < tradeQty) {
    // 已實現加總不足 → 提示
    const ok = await confirmDialog(
      '⚠️ 數量不一致',
      `這筆 trade 是 <b>${fmt(tradeQty)}</b> 股，但同條件的已實現只加總到 <b>${fmt(realizedQty)}</b> 股（共 ${sameGroup.length} 筆），<br>
       還差 ${fmt(tradeQty - realizedQty)} 股沒配對。<br><br>
       仍要強制配對嗎？利息 ${fmt(trade.marginInterest||0)} 會全部分給這 ${sameGroup.length} 筆，<br>
       會多算（多出去的部分視為已實現本身漏匯）。`
    );
    if (!ok) return;
  } else {
    // 已實現加總超過 trade → 警告，預設不允許全選
    const ok = await confirmDialog(
      '⚠️ 已實現數量超過 trade',
      `同條件的已實現加總 ${fmt(realizedQty)} 股已超過這筆 trade 的 ${fmt(tradeQty)} 股。<br>
       資料可能有問題（trade 漏匯或拆分錯誤）。<br><br>
       仍要把這 ${sameGroup.length} 筆都配對到這 trade 嗎？利息會被稀釋。`
    );
    if (!ok) {
      // 使用者拒絕 → 改成只配當前這筆
      const ok2 = await confirmDialog(
        '只配對這 1 筆？',
        `只把當前的 ${realized.code} ${fmt(realized.qty)} 股配對到此 trade，其他不動。`
      );
      if (!ok2) return;
      willGroup = [realized];
    }
  }

  // 套用：把 willGroup 內所有已實現都設定 _manualMatchTradeKey
  for (const r of willGroup) {
    r._manualMatchTradeKey = result;
  }
  // 不在 willGroup 裡的同條件已實現：清掉 manual link（萬一之前指過別的）
  if (willGroup !== sameGroup) {
    for (const r of sameGroup) {
      if (!willGroup.includes(r) && r._manualMatchTradeKey === result) {
        delete r._manualMatchTradeKey;
      }
    }
  }

  // 重跑 enrichment（會處理群組分攤）
  Parsers.enrichRealizedWithInterest(acc.realized, acc.trades);
  save();
  renderAll();
  toast(willGroup.length > 1
    ? `已群組配對 ${willGroup.length} 筆，利息按比例分攤`
    : '已手動配對', 'ok');
}

// ============================================================
// 編輯調整金額對話框
// 兩個獨立欄位：盈虧調整、回補資金（被迫沖銷的補繳款）
// 都會直接相加成「調整總額」，再加到原始盈虧得到實際盈虧
// ============================================================
async function showAdjustDialog(realized) {
  const acc = getCurrentAccount();
  if (!acc) return;

  const initAdjust = realized.adjust || 0;
  const initRefund = realized.refund || 0;
  const initNote = realized.note || '';
  const original = realized.pl || 0;

  const result = await showModal({
    title: `編輯調整 - ${realized.code} ${realized.name} ${realized.sellDate}`,
    html: `
      <div style="min-width:460px;max-width:560px">
        <div style="background:var(--surface-2);padding:10px 12px;border-radius:6px;margin-bottom:12px;font-size:13px">
          原始盈虧：<strong class="${plClass(original)}">${fmt(original, {sign:true})}</strong>
        </div>

        <label style="display:block;margin-bottom:14px">
          <div style="font-weight:600;font-size:13px;margin-bottom:4px">盈虧調整</div>
          <input type="number" id="adj-pl" value="${initAdjust}" step="any" style="width:100%">
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px">
            修正券商給的盈虧錯誤（如轉融券沖銷後的價差）。正值=往上調，負值=往下調
          </div>
        </label>

        <label style="display:block;margin-bottom:14px">
          <div style="font-weight:600;font-size:13px;margin-bottom:4px">回補資金</div>
          <input type="number" id="adj-refund" value="${initRefund}" step="any" style="width:100%">
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px">
            被迫沖銷產生的補繳款（多付出去填負值，意外退回填正值）
          </div>
        </label>

        <label style="display:block;margin-bottom:14px">
          <div style="font-weight:600;font-size:13px;margin-bottom:4px">備註</div>
          <input type="text" id="adj-note" value="${initNote.replace(/"/g,'&quot;')}" style="width:100%" placeholder="例：11/15 轉融券沖銷">
        </label>

        <div style="background:var(--primary-light);padding:10px 12px;border-radius:6px;font-size:13px" id="adj-preview">
          <div>調整合計：<strong id="adj-preview-total">—</strong></div>
          <div style="margin-top:4px">實際盈虧：<strong id="adj-preview-actual">—</strong></div>
        </div>
      </div>
    `,
    confirmText: '儲存',
    onConfirm: (body) => ({
      adjust: parseFloat(body.querySelector('#adj-pl').value) || 0,
      refund: parseFloat(body.querySelector('#adj-refund').value) || 0,
      note: body.querySelector('#adj-note').value
    })
  });

  // modal 渲染後綁即時預覽
  setTimeout(() => {
    const adjI = document.getElementById('adj-pl');
    const refI = document.getElementById('adj-refund');
    const totalEl = document.getElementById('adj-preview-total');
    const actualEl = document.getElementById('adj-preview-actual');
    if (!adjI || !refI || !totalEl || !actualEl) return;

    const recalc = () => {
      const a = parseFloat(adjI.value) || 0;
      const r = parseFloat(refI.value) || 0;
      const total = a + r;
      const actual = original + total;
      const tCls = total > 0 ? 'pos' : (total < 0 ? 'neg' : '');
      const aCls = actual > 0 ? 'pos' : (actual < 0 ? 'neg' : '');
      totalEl.innerHTML = `<span class="${tCls}">${fmt(total, {sign:true})}</span>`;
      actualEl.innerHTML = `<span class="${aCls}">${fmt(actual, {sign:true})}</span>`;
    };
    adjI.oninput = recalc;
    refI.oninput = recalc;
    recalc();
  }, 0);

  if (!result) return;

  realized.adjust = result.adjust;
  realized.refund = result.refund;
  realized.note = result.note;

  // 同步存到 adjustments，這樣重新匯入會被還原
  const k = realizedKey(realized);
  acc.adjustments[k] = {
    adjust: result.adjust,
    refund: result.refund,
    note: result.note
  };

  save();
  renderAll();
  toast('已儲存調整', 'ok');
}

// ============================================================
// 編輯 / 刪除已實現損益
// ============================================================
async function editRealizedDialog(realized) {
  const acc = getCurrentAccount();
  if (!acc) return;

  const result = await showModal({
    title: `編輯已實現 - ${realized.code} ${realized.name||''}`,
    html: `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;min-width:560px">
        <label>代號<input type="text" id="rz-code" value="${realized.code||''}"></label>
        <label>名稱<input type="text" id="rz-name" value="${(realized.name||'').replace(/"/g,'&quot;')}"></label>
        <label>類別
          <select id="rz-cat">
            <option value="現股" ${realized.sellCategory==='現股'?'selected':''}>現股</option>
            <option value="融資" ${realized.sellCategory==='融資'?'selected':''}>融資</option>
            <option value="融券" ${realized.sellCategory==='融券'?'selected':''}>融券</option>
            <option value="現股當沖" ${realized.sellCategory==='現股當沖'?'selected':''}>現股當沖</option>
          </select>
        </label>
        <label>數量<input type="number" id="rz-qty" value="${realized.qty||0}"></label>
        <label>賣出日<input type="date" id="rz-selldate" value="${(realized.sellDate||'').replace(/\//g,'-')}"></label>
        <label>賣價<input type="number" id="rz-sellprice" value="${realized.sellPrice||0}" step="0.01"></label>
        <label>買進日<input type="date" id="rz-buydate" value="${(realized.buyDate||'').replace(/\//g,'-')}"></label>
        <label>買價<input type="number" id="rz-buyprice" value="${realized.buyPrice||0}" step="0.01"></label>
        <label>沖銷成本<input type="number" id="rz-cost" value="${realized.cost||0}"></label>
        <label>原始盈虧<input type="number" id="rz-pl" value="${realized.pl||0}"></label>
        <label>手續費<input type="number" id="rz-fee" value="${realized.fee||0}"></label>
        <label>交易稅<input type="number" id="rz-tax" value="${realized.tax||0}"></label>
        <label>利息<input type="number" id="rz-interest" value="${realized.interest||0}"></label>
        <label>融券手續費<input type="number" id="rz-shortfee" value="${realized.shortFee||0}"></label>
      </div>
    `,
    confirmText: '儲存',
    onConfirm: (body) => ({
      code: body.querySelector('#rz-code').value.trim(),
      name: body.querySelector('#rz-name').value.trim(),
      sellCategory: body.querySelector('#rz-cat').value,
      qty: parseFloat(body.querySelector('#rz-qty').value) || 0,
      sellDate: body.querySelector('#rz-selldate').value,
      sellPrice: parseFloat(body.querySelector('#rz-sellprice').value) || 0,
      buyDate: body.querySelector('#rz-buydate').value,
      buyPrice: parseFloat(body.querySelector('#rz-buyprice').value) || 0,
      cost: parseFloat(body.querySelector('#rz-cost').value) || 0,
      pl: parseFloat(body.querySelector('#rz-pl').value) || 0,
      fee: parseFloat(body.querySelector('#rz-fee').value) || 0,
      tax: parseFloat(body.querySelector('#rz-tax').value) || 0,
      interest: parseFloat(body.querySelector('#rz-interest').value) || 0,
      shortFee: parseFloat(body.querySelector('#rz-shortfee').value) || 0
    })
  });
  if (!result) return;

  // 用新值覆蓋
  realized.code = result.code;
  realized.name = result.name;
  realized.sellCategory = result.sellCategory;
  realized.qty = result.qty;
  realized.sellDate = Parsers.formatDate(result.sellDate);
  realized.sellPrice = result.sellPrice;
  realized.buyDate = result.buyDate ? Parsers.formatDate(result.buyDate) : '';
  realized.buyPrice = result.buyPrice;
  realized.cost = result.cost;
  realized.pl = result.pl;
  realized.fee = result.fee;
  realized.tax = result.tax;
  realized.interest = result.interest;
  realized.shortFee = result.shortFee;
  realized._userEdited = true;  // 標記：手動編輯過，匯入時不要被覆蓋

  save();
  renderAll();
  toast('已更新', 'ok');
}

async function deleteRealizedConfirm(realized) {
  const acc = getCurrentAccount();
  if (!acc) return;
  const ok = await confirmDialog(
    '刪除已實現損益',
    `確定刪除這筆？<br><br><b>${realized.code} ${realized.name||''}</b><br>${realized.sellDate} 賣出 ${fmt(realized.qty)} 股，盈虧 ${fmt(realized.pl,{sign:true})}<br><br>同時會移除「調整金額/備註」記憶，下次重新匯入該筆時不會自動還原。`
  );
  if (!ok) return;

  const idx = acc.realized.indexOf(realized);
  if (idx < 0) return;
  acc.realized.splice(idx, 1);

  // 同時清掉 adjustments 紀錄
  const k = realizedKey(realized);
  if (acc.adjustments && acc.adjustments[k]) delete acc.adjustments[k];

  save();
  renderAll();
  toast('已刪除', 'ok');
}

function renderRealized() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#realizedTable tbody');
  const alertWrap = document.getElementById('matchAlerts');
  if (alertWrap) alertWrap.innerHTML = '';
  if (!acc || !acc.realized.length) {
    tb.innerHTML = '<tr><td colspan="18" class="empty-state">尚無已實現損益資料</td></tr>';
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
    adjust += adjustTotal(r);
    actual += actualPL(r);
    interest += (r.interest || 0);
    shortFee += (r.shortFee || 0);
  }
  setVal('rzOriginal', original, true);
  setVal('rzAdjust', adjust, true);
  setVal('rzActual', actual, true);
  setVal('rzInterest', interest);
  setVal('rzShortFee', shortFee);

  // 渲染配對提醒
  renderMatchAlerts(acc);

  tb.innerHTML = filtered.map((r, idx) => {
    const realIdx = acc.realized.indexOf(r);
    const adjT = adjustTotal(r);
    const actual = actualPL(r);
    const isMargin = (r.sellCategory === '融資' || r.sellCategory === '融券');
    let rowClass = '';
    // 配對狀態（只跟融資/融券有關）
    let matchPart = '';
    if (r._unmatched) {
      rowClass = 'unmatched-row';
      matchPart = `<button class="btn-mini" data-act="manual-match" data-idx="${realIdx}" style="padding:2px 6px;font-size:11px">🔗</button>`;
    } else if (r._matchSource === 'manual') {
      rowClass = 'manual-match-row';
      matchPart = `<span class="match-source-tag manual" title="手動配對">手動</span><button class="btn-mini" data-act="unmatch" data-idx="${realIdx}" style="padding:2px 5px;font-size:10px">✕</button>`;
    } else if (isMargin && (r._matchSource === 'auto' || r._matchSource === 'auto-split')) {
      matchPart = '<span class="match-source-tag auto" title="自動配對">自動</span>';
    }
    // 編輯 + 刪除按鈕（永遠顯示）
    const editPart = `<button class="btn-mini" data-act="edit-realized" data-idx="${realIdx}" style="padding:2px 6px;font-size:11px" title="編輯">✏️</button>`;
    const delPart = `<button class="btn-mini danger" data-act="del-realized" data-idx="${realIdx}" style="padding:2px 6px;font-size:11px" title="刪除">🗑️</button>`;
    const actionCell = `<div style="display:flex;gap:3px;align-items:center;flex-wrap:wrap">${matchPart}${editPart}${delPart}</div>`;

    // 調整金額按鈕
    let adjBtnClass = 'adjust-btn';
    let adjLabel = '＋ 調整';
    if (adjT !== 0) {
      adjBtnClass += ' has-value ' + (adjT > 0 ? 'pos' : 'neg');
      adjLabel = fmt(adjT, {sign: true});
    }

    return `
      <tr data-idx="${realIdx}" class="${rowClass}">
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
        <td class="hl"><button class="${adjBtnClass}" data-act="edit-adjust" data-idx="${realIdx}">${adjLabel}</button></td>
        <td title="${(r.note || '').replace(/"/g,'&quot;')}" style="max-width:140px;overflow:hidden;text-overflow:ellipsis">${r.note || '—'}</td>
        <td class="${plClass(r.pl)}">${fmt(r.pl, {sign:true})}</td>
        <td class="hl ${plClass(actual)}">${fmt(actual, {sign:true})}</td>
        <td>${actionCell}</td>
      </tr>
    `;
  }).join('');

  // 綁定按鈕（配對 + 調整金額）
  tb.querySelectorAll('button[data-act]').forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      const idx = parseInt(btn.dataset.idx, 10);
      const r = acc.realized[idx];
      if (!r) return;
      const act = btn.dataset.act;
      if (act === 'manual-match') await showManualMatchDialog(r);
      else if (act === 'unmatch') {
        delete r._manualMatchTradeKey;
        Parsers.enrichRealizedWithInterest(acc.realized, acc.trades);
        save();
        renderAll();
        toast('已解除手動配對', 'ok');
      }
      else if (act === 'edit-adjust') await showAdjustDialog(r);
      else if (act === 'edit-realized') await editRealizedDialog(r);
      else if (act === 'del-realized') await deleteRealizedConfirm(r);
    };
  });
}

// ---------- 未實現損益 ----------
// ============================================================
// 抓即時股價（Yahoo Finance）
// ============================================================

// 是否已自動拓過（每次切到頁籤只自動拓一次，不要重複觸發）
let _priceAutoFetched = false;
let _priceFetching = false;

function setPriceStatus(text, kind = '') {
  const el = document.getElementById('priceStatus');
  if (!el) return;
  el.textContent = text;
  el.className = 'price-status ' + kind;
}

// ---------- TWSE/TPEX OpenAPI 快取（避免重複拉全市場） ----------
let _priceCache = {
  twseAt: 0,    // 上次抓 TWSE 的時間（毫秒）
  twseMap: null, // 代號 → {price, name}
  tpexAt: 0,
  tpexMap: null
};
const PRICE_CACHE_TTL = 5 * 60 * 1000; // 5 分鐘

async function fetchTWSEPrices() {
  // 證交所每日收盤行情（全市場一次拿）
  // 欄位範例：{ Code, Name, ClosingPrice, ChangePrice, ... }
  const url = 'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_AVG_ALL';
  const resp = await fetch(url, { method: 'GET' });
  if (!resp.ok) throw new Error(`TWSE HTTP ${resp.status}`);
  const data = await resp.json();
  const map = new Map();
  for (const r of data) {
    const code = String(r.Code || '').trim();
    const price = parseFloat(r.ClosingPrice);
    if (code && !isNaN(price) && price > 0) {
      map.set(code, { price, name: r.Name || '' });
    }
  }
  return map;
}

async function fetchTPEXPrices() {
  // 櫃買中心上櫃每日收盤
  const url = 'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes';
  const resp = await fetch(url, { method: 'GET' });
  if (!resp.ok) throw new Error(`TPEX HTTP ${resp.status}`);
  const data = await resp.json();
  const map = new Map();
  for (const r of data) {
    // 欄位名稱可能是 SecuritiesCompanyCode 或 Code
    const code = String(r.SecuritiesCompanyCode || r.Code || '').trim();
    const price = parseFloat(r.Close || r.ClosingPrice);
    const name = r.CompanyName || r.Name || '';
    if (code && !isNaN(price) && price > 0) {
      map.set(code, { price, name });
    }
  }
  return map;
}

// 主流程：抓股價並覆寫到目前帳戶的 unrealized
async function refreshPrices(silent = false) {
  if (_priceFetching) return;
  const acc = getCurrentAccount();
  if (!acc || !acc.unrealized || !acc.unrealized.length) {
    if (!silent) setPriceStatus('無持股可更新', 'error');
    return;
  }

  _priceFetching = true;
  setPriceStatus('正在抓取證交所/櫃買中心資料…', 'loading');

  try {
    const now = Date.now();

    // TWSE：5 分鐘內有快取就用
    if (!_priceCache.twseMap || (now - _priceCache.twseAt) > PRICE_CACHE_TTL) {
      try {
        _priceCache.twseMap = await fetchTWSEPrices();
        _priceCache.twseAt = now;
      } catch (e) {
        console.warn('TWSE 抓取失敗，繼續嘗試 TPEX：', e);
        if (!_priceCache.twseMap) _priceCache.twseMap = new Map();
      }
    }

    // TPEX：5 分鐘內有快取就用
    if (!_priceCache.tpexMap || (now - _priceCache.tpexAt) > PRICE_CACHE_TTL) {
      try {
        _priceCache.tpexMap = await fetchTPEXPrices();
        _priceCache.tpexAt = now;
      } catch (e) {
        console.warn('TPEX 抓取失敗：', e);
        if (!_priceCache.tpexMap) _priceCache.tpexMap = new Map();
      }
    }

    const twse = _priceCache.twseMap;
    const tpex = _priceCache.tpexMap;

    if (twse.size === 0 && tpex.size === 0) {
      throw new Error('證交所和櫃買中心都抓不到資料（網路問題或服務暫時無法使用）');
    }

    // 套用：覆寫 unrealized 表的 price、marketValue、pl、rate
    let updated = 0, failed = [];
    for (const x of acc.unrealized) {
      const code = String(x.code || '').trim();
      const info = twse.get(code) || tpex.get(code);
      if (info && info.price > 0) {
        x.price = info.price;
        x.marketValue = info.price * (x.qty || 0);
        x.pl = x.marketValue - (x.cost || 0);
        if (x.cost > 0) {
          const r = (x.pl / x.cost) * 100;
          x.rate = (r > 0 ? '+' : '') + r.toFixed(2) + '%';
        }
        x._priceUpdated = true;
        x._priceUpdatedAt = new Date().toISOString();
        updated++;
      } else {
        x._priceUpdated = false;
        failed.push(x.code);
      }
    }

    save();
    const ts = new Date();
    const timeStr = ts.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
    const totalSrc = `（上市 ${twse.size} 檔，上櫃 ${tpex.size} 檔）`;
    if (failed.length === 0) {
      setPriceStatus(`✓ 已更新 ${updated} 檔 ${totalSrc}，最後更新 ${timeStr}`, 'success');
    } else {
      setPriceStatus(`✓ ${updated} 檔成功，${failed.length} 檔找不到 (${failed.slice(0,3).join(',')}${failed.length>3?'...':''})，更新時間 ${timeStr}`, 'success');
    }
    renderAll();
  } catch (e) {
    console.error('refreshPrices error:', e);
    setPriceStatus(`✗ 更新失敗：${e.message}`, 'error');
  } finally {
    _priceFetching = false;
  }
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

  tb.innerHTML = acc.unrealized.map(x => {
    const priceClass = x._priceUpdated ? 'price-updated-cell' : '';
    const failedTag = (x._priceUpdated === false) ? '<span class="price-failed-tag" title="無法從 Yahoo 取得即時價">未取得</span>' : '';
    return `
    <tr>
      <td>${x.code}</td>
      <td>${x.name}${failedTag}</td>
      <td>${x.category}</td>
      <td>${fmt(x.qty)}</td>
      <td class="${priceClass}">${fmt(x.price, {decimals:2})}</td>
      <td class="${priceClass}">${fmt(x.marketValue)}</td>
      <td>${fmt(x.cost)}</td>
      <td>${fmt(x.avgCost, {decimals:4})}</td>
      <td>${fmt(x.fee)}</td>
      <td>${fmt(x.tax)}</td>
      <td>${fmt(x.interest)}</td>
      <td class="${plClass(x.pl)} ${priceClass}">${fmt(x.pl, {sign:true})}</td>
      <td class="${plClass(x.pl)} ${priceClass}">${x.rate || (x.cost ? fmtPct(x.pl/x.cost*100) : '—')}</td>
    </tr>
  `;}).join('');

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
// 當沖分析
// 規則：同一交易日 + 同一股票，買賣配對的數量視為當沖
// 例：當天買 1500 股 + 賣 1000 股 → 當沖 1000 股
// ============================================================
function analyzeDayTrades(acc) {
  // 1. 把投資明細按 (日期, 代號) 分組
  const groups = new Map();
  for (const t of (acc.trades || [])) {
    if (!t.date || !t.code) continue;
    if (t.category === '資轉現' || t.action === '資轉現') continue; // 資轉現不算
    const k = `${t.date}|${t.code}`;
    if (!groups.has(k)) groups.set(k, { date: t.date, code: t.code, name: t.name, buys: [], sells: [] });
    const g = groups.get(k);
    if (t.action === '買') g.buys.push(t);
    else if (t.action === '賣') g.sells.push(t);
  }

  // 2. 算每組的當沖配對
  const dayTrades = []; // 每筆當沖（一檔股票的某天當沖明細）
  for (const g of groups.values()) {
    if (g.buys.length === 0 || g.sells.length === 0) continue;
    const buyQty = g.buys.reduce((s,t) => s + (t.qty || 0), 0);
    const sellQty = g.sells.reduce((s,t) => s + (t.qty || 0), 0);
    const matchQty = Math.min(buyQty, sellQty); // 當沖配對股數
    if (matchQty <= 0) continue;

    // 加權平均價
    const buyAmount = g.buys.reduce((s,t) => s + (t.amount || 0), 0);
    const sellAmount = g.sells.reduce((s,t) => s + (t.amount || 0), 0);
    const avgBuyPrice = buyQty > 0 ? buyAmount / buyQty : 0;
    const avgSellPrice = sellQty > 0 ? sellAmount / sellQty : 0;

    // 當沖損益 = (賣價 - 買價) × 配對股數 - 該日的手續費和稅
    const grossPL = (avgSellPrice - avgBuyPrice) * matchQty;
    const buyFee = g.buys.reduce((s,t) => s + (t.fee || 0), 0);
    const sellFee = g.sells.reduce((s,t) => s + (t.fee || 0), 0);
    const sellTax = g.sells.reduce((s,t) => s + (t.tax || 0), 0);
    // 手續費按配對比例分攤
    const buyFeeRatio = buyQty > 0 ? matchQty / buyQty : 0;
    const sellFeeRatio = sellQty > 0 ? matchQty / sellQty : 0;
    const allocFee = buyFee * buyFeeRatio + sellFee * sellFeeRatio;
    const allocTax = sellTax * sellFeeRatio;

    const netPL = grossPL - allocFee - allocTax;
    const turnover = avgBuyPrice * matchQty + avgSellPrice * matchQty; // 雙邊成交金額
    const buySideAmount = avgBuyPrice * matchQty; // 用買進金額算報酬率

    dayTrades.push({
      date: g.date,
      code: g.code,
      name: g.name || '',
      matchQty,
      avgBuyPrice,
      avgSellPrice,
      buyAmount: buySideAmount,
      sellAmount: avgSellPrice * matchQty,
      turnover,
      grossPL,
      fee: allocFee,
      tax: allocTax,
      netPL,
      rate: buySideAmount > 0 ? (netPL / buySideAmount) * 100 : 0,
      buyCount: g.buys.length,
      sellCount: g.sells.length
    });
  }

  return dayTrades.sort((a,b) => b.date.localeCompare(a.date));
}

// 按日期匯總當沖（每天一行）
function aggregateDayTradesByDay(dayTrades) {
  const byDate = new Map();
  for (const dt of dayTrades) {
    if (!byDate.has(dt.date)) {
      byDate.set(dt.date, {
        date: dt.date, count: 0, totalQty: 0,
        turnover: 0, buyAmount: 0,
        netPL: 0, fee: 0, tax: 0,
        winCount: 0, lossCount: 0,
        items: []
      });
    }
    const d = byDate.get(dt.date);
    d.count++;
    d.totalQty += dt.matchQty;
    d.turnover += dt.turnover;
    d.buyAmount += dt.buyAmount;
    d.netPL += dt.netPL;
    d.fee += dt.fee;
    d.tax += dt.tax;
    if (dt.netPL > 0) d.winCount++;
    else if (dt.netPL < 0) d.lossCount++;
    d.items.push(dt);
  }
  // 算每天的勝率/報酬率
  for (const d of byDate.values()) {
    d.winRate = d.count > 0 ? (d.winCount / d.count) * 100 : 0;
    d.rate = d.buyAmount > 0 ? (d.netPL / d.buyAmount) * 100 : 0;
  }
  return [...byDate.values()].sort((a,b) => b.date.localeCompare(a.date));
}

// 把當沖按月份匯總（給每月損益表用）
function aggregateDayTradesByMonth(dayTrades) {
  const byMonth = new Map();
  for (const dt of dayTrades) {
    const k = getMonthKey(dt.date);
    if (!k) continue;
    if (!byMonth.has(k)) {
      byMonth.set(k, { count: 0, turnover: 0, buyAmount: 0, netPL: 0, winCount: 0 });
    }
    const m = byMonth.get(k);
    m.count++;
    m.turnover += dt.turnover;
    m.buyAmount += dt.buyAmount;
    m.netPL += dt.netPL;
    if (dt.netPL > 0) m.winCount++;
  }
  for (const m of byMonth.values()) {
    m.winRate = m.count > 0 ? (m.winCount / m.count) * 100 : 0;
    m.rate = m.buyAmount > 0 ? (m.netPL / m.buyAmount) * 100 : 0;
  }
  return byMonth;
}

// 把當沖按股票代號匯總（給個股分析用）
function aggregateDayTradesByCode(dayTrades) {
  const byCode = new Map();
  for (const dt of dayTrades) {
    if (!byCode.has(dt.code)) {
      byCode.set(dt.code, { count: 0, turnover: 0, buyAmount: 0, netPL: 0, winCount: 0 });
    }
    const s = byCode.get(dt.code);
    s.count++;
    s.turnover += dt.turnover;
    s.buyAmount += dt.buyAmount;
    s.netPL += dt.netPL;
    if (dt.netPL > 0) s.winCount++;
  }
  for (const s of byCode.values()) {
    s.winRate = s.count > 0 ? (s.winCount / s.count) * 100 : 0;
    s.rate = s.buyAmount > 0 ? (s.netPL / s.buyAmount) * 100 : 0;
  }
  return byCode;
}

// ---------- 每日當沖渲染 ----------
const _expandedDays = new Set();

function renderDayTradeCards(dayTrades) {
  const wrap = document.getElementById('dayTradeCards');
  if (!wrap) return;
  if (!dayTrades || !dayTrades.length) {
    wrap.innerHTML = '';
    return;
  }
  const days = aggregateDayTradesByDay(dayTrades);
  const totalCount = dayTrades.length;
  const totalDays = days.length;
  const totalNetPL = dayTrades.reduce((s, x) => s + x.netPL, 0);
  const totalBuyAmount = dayTrades.reduce((s, x) => s + x.buyAmount, 0);
  const totalTurnover = dayTrades.reduce((s, x) => s + x.turnover, 0);
  const winCount = dayTrades.filter(x => x.netPL > 0).length;
  const winRate = totalCount > 0 ? (winCount / totalCount) * 100 : 0;
  const avgRate = totalBuyAmount > 0 ? (totalNetPL / totalBuyAmount) * 100 : 0;
  const profitDay = days.filter(d => d.netPL > 0).length;
  const lossDay = days.filter(d => d.netPL < 0).length;
  const dayWinRate = totalDays > 0 ? (profitDay / totalDays) * 100 : 0;

  const cards = [
    `<div class="card"><div class="label">當沖總筆數</div><div class="value">${totalCount}</div><div style="font-size:11px;color:var(--text-muted);margin-top:2px">共 ${totalDays} 個交易日</div></div>`,
    `<div class="card"><div class="label">當沖總損益</div><div class="value ${plClass(totalNetPL)}">${fmt(totalNetPL,{sign:true})}</div></div>`,
    `<div class="card"><div class="label">總成交金額（雙邊）</div><div class="value">${fmt(totalTurnover)}</div></div>`,
    `<div class="card"><div class="label">當沖勝率</div><div class="value ${winRate>=50?'pos':'neg'}">${winRate.toFixed(1)}%</div><div style="font-size:11px;color:var(--text-muted);margin-top:2px">${winCount}/${totalCount} 筆獲利</div></div>`,
    `<div class="card"><div class="label">日勝率</div><div class="value ${dayWinRate>=50?'pos':'neg'}">${dayWinRate.toFixed(1)}%</div><div style="font-size:11px;color:var(--text-muted);margin-top:2px">${profitDay} 賺 / ${lossDay} 賠</div></div>`,
    `<div class="card"><div class="label">總報酬率</div><div class="value ${plClass(avgRate)}">${fmtPct(avgRate)}</div><div style="font-size:11px;color:var(--text-muted);margin-top:2px">損益÷買進金額</div></div>`
  ];
  wrap.innerHTML = cards.join('');
}

function renderDayTrades() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#dayTradeTable tbody');
  const yearSel = document.getElementById('dayTradeYear');
  if (!tb) return;

  if (!acc) {
    tb.innerHTML = '<tr><td colspan="9" class="empty-state">尚未選擇帳戶</td></tr>';
    renderDayTradeCards(null);
    return;
  }

  const allDayTrades = analyzeDayTrades(acc);
  if (!allDayTrades.length) {
    tb.innerHTML = '<tr><td colspan="9" class="empty-state">沒有當沖紀錄（同一日同一股票需有買有賣才算當沖）</td></tr>';
    renderDayTradeCards(null);
    return;
  }

  // 年份下拉
  const years = [...new Set(allDayTrades.map(d => d.date.slice(0,4)))].sort().reverse();
  const currentYear = yearSel.value;
  yearSel.innerHTML = '<option value="">所有年份</option>' +
    years.map(y => `<option value="${y}" ${y===currentYear?'selected':''}>${y}</option>`).join('');

  const filtered = currentYear
    ? allDayTrades.filter(d => d.date.startsWith(currentYear))
    : allDayTrades;

  renderDayTradeCards(filtered);

  if (!filtered.length) {
    tb.innerHTML = '<tr><td colspan="9" class="empty-state">該年份無當沖紀錄</td></tr>';
    return;
  }

  const days = aggregateDayTradesByDay(filtered);

  const rows = [];
  for (const d of days) {
    const expanded = _expandedDays.has(d.date);
    rows.push(`
      <tr class="month-row" data-date="${d.date}">
        <td><span class="month-toggle ${expanded?'expanded':''}">▶</span></td>
        <td><strong>${d.date}</strong></td>
        <td>${d.count}</td>
        <td>${fmt(d.totalQty)}</td>
        <td>${fmt(d.turnover)}</td>
        <td class="${plClass(d.netPL)}"><strong>${fmt(d.netPL,{sign:true})}</strong></td>
        <td><span class="pos">${d.winCount}</span> / <span class="neg">${d.lossCount}</span></td>
        <td class="${d.winRate>=50?'pos':'neg'}">${d.winRate.toFixed(1)}%</td>
        <td class="${plClass(d.rate)}">${fmtPct(d.rate)}</td>
      </tr>
    `);
    if (expanded) {
      rows.push(`<tr class="month-detail-row"><td colspan="9">${renderDayTradeDetail(d)}</td></tr>`);
    }
  }
  tb.innerHTML = rows.join('');

  tb.querySelectorAll('.month-row').forEach(row => {
    row.onclick = () => {
      const k = row.dataset.date;
      if (_expandedDays.has(k)) _expandedDays.delete(k); else _expandedDays.add(k);
      renderDayTrades();
    };
  });
}

function renderDayTradeDetail(d) {
  const html = ['<div class="nested">'];
  html.push('<h4 style="padding:10px 14px 0">當日當沖明細（' + d.items.length + ' 檔）</h4>');
  html.push('<table class="loan-subtable"><thead><tr>');
  html.push('<th>代號</th><th>名稱</th><th>當沖股數</th><th>平均買價</th><th>平均賣價</th><th>毛損益</th><th>手續費+稅</th><th>淨損益</th><th>報酬率</th>');
  html.push('</tr></thead><tbody>');
  for (const dt of d.items) {
    html.push(`<tr>
      <td>${dt.code}</td>
      <td>${dt.name}</td>
      <td>${fmt(dt.matchQty)}</td>
      <td>${fmt(dt.avgBuyPrice,{decimals:2})}</td>
      <td>${fmt(dt.avgSellPrice,{decimals:2})}</td>
      <td class="${plClass(dt.grossPL)}">${fmt(dt.grossPL,{sign:true})}</td>
      <td>${fmt(dt.fee + dt.tax)}</td>
      <td class="${plClass(dt.netPL)}"><strong>${fmt(dt.netPL,{sign:true})}</strong></td>
      <td class="${plClass(dt.rate)}">${fmtPct(dt.rate)}</td>
    </tr>`);
  }
  html.push('</tbody></table></div>');
  return html.join('');
}

function exportDayTradesExcel() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const dayTrades = analyzeDayTrades(acc);
  if (!dayTrades.length) return toast('沒有當沖紀錄可匯出', 'err');

  const wb = XLSX.utils.book_new();

  const h1 = ['日期','代號','名稱','當沖股數','平均買價','平均賣價','買進金額','賣出金額','雙邊成交','毛損益','手續費','交易稅','淨損益','報酬率(%)'];
  const r1 = dayTrades.map(d => [
    d.date, d.code, d.name, d.matchQty,
    +d.avgBuyPrice.toFixed(2), +d.avgSellPrice.toFixed(2),
    Math.round(d.buyAmount), Math.round(d.sellAmount), Math.round(d.turnover),
    Math.round(d.grossPL), Math.round(d.fee), Math.round(d.tax), Math.round(d.netPL),
    +d.rate.toFixed(2)
  ]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h1, ...r1]), '每筆當沖');

  const days = aggregateDayTradesByDay(dayTrades);
  const h2 = ['日期','當沖檔數','當沖股數','雙邊成交金額','淨損益','勝/敗','勝率(%)','報酬率(%)'];
  const r2 = days.map(d => [
    d.date, d.count, d.totalQty, Math.round(d.turnover),
    Math.round(d.netPL), `${d.winCount}/${d.lossCount}`,
    +d.winRate.toFixed(2), +d.rate.toFixed(2)
  ]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h2, ...r2]), '每日當沖');

  const ts = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `當沖分析_${acc.name}_${ts}.xlsx`);
  toast('已匯出 Excel', 'ok');
}


function buildStockAnalysis(acc) {
  // key: code, value: 統計
  const stocks = new Map();
  const ensure = (code, name) => {
    if (!stocks.has(code)) {
      stocks.set(code, {
        code, name: name || '',
        buyCount: 0, sellCount: 0,
        buyAmount: 0, sellAmount: 0,
        buyQty: 0, sellQty: 0,    // 累計買/賣股數（用來算平均價）
        realizedCount: 0,
        realizedPL: 0,    // 從已實現損益的「盈虧」加總
        adjust: 0,
        interest: 0,      // 融資利息（從已實現補上的）
        shortFee: 0,      // 融券手續費
        tradeInterest: 0, // 投資明細上所有與這檔有關的利息（含資轉現）
        currentQty: 0,    // 目前持有股數（買-賣，含資轉現的轉換）
        marketValue: 0,
        cost: 0,
        unrealizedPL: 0,
        firstTradeDate: null,
        lastTradeDate: null,
        trades: [],
        realizedItems: []
      });
    }
    const s = stocks.get(code);
    if (name && !s.name) s.name = name;
    return s;
  };

  // 從投資明細統計
  for (const t of (acc.trades || [])) {
    if (!t.code) continue;
    const s = ensure(t.code, t.name);
    s.trades.push(t);
    s.tradeInterest += (t.marginInterest || 0);

    // 第一/最後交易日
    if (!s.firstTradeDate || (t.date && t.date < s.firstTradeDate)) s.firstTradeDate = t.date;
    if (!s.lastTradeDate || (t.date && t.date > s.lastTradeDate)) s.lastTradeDate = t.date;

    // 資轉現：不算買賣次數，但會更新持股（融資轉現股，數量不變）
    if (t.category === '資轉現' || t.action === '資轉現') {
      // 不影響 currentQty（融資減一筆、現股加一筆，淨零）
      continue;
    }

    if (t.action === '買') {
      s.buyCount++;
      s.buyAmount += (t.amount || 0);
      s.buyQty += (t.qty || 0);
      s.currentQty += (t.qty || 0);
    } else if (t.action === '賣') {
      s.sellCount++;
      s.sellAmount += (t.amount || 0);
      s.sellQty += (t.qty || 0);
      s.currentQty -= (t.qty || 0);
    }
  }

  // 從已實現損益統計
  for (const r of (acc.realized || [])) {
    if (!r.code) continue;
    const s = ensure(r.code, r.name);
    s.realizedCount++;
    s.realizedPL += (r.pl || 0);
    s.adjust += adjustTotal(r);
    s.interest += (r.interest || 0);
    s.shortFee += (r.shortFee || 0);
    s.realizedItems.push(r);
  }

  // 從未實現損益統計（持有市值/成本）
  for (const x of (acc.unrealized || [])) {
    if (!x.code) continue;
    const s = ensure(x.code, x.name);
    s.marketValue += (x.marketValue || 0);
    s.cost += (x.cost || 0);
    s.unrealizedPL += (x.pl || 0);
    // 如果交易明細沒匯入或不完整，用未實現的數量當持股數
    if (s.currentQty <= 0 && (x.qty || 0) > 0) {
      s.currentQty = x.qty;
    }
  }

  // 計算實際損益和平均報酬率
  for (const s of stocks.values()) {
    s.actualPL = s.realizedPL + s.adjust;
    // 平均報酬率：累計實現損益 / 累計買進金額
    s.avgRate = s.buyAmount > 0 ? (s.actualPL / s.buyAmount) * 100 : 0;
    // 平均買價、平均賣價（用累計金額/股數，避免單筆價格簡單平均的偏差）
    s.avgBuyPrice = s.buyQty > 0 ? s.buyAmount / s.buyQty : 0;
    s.avgSellPrice = s.sellQty > 0 ? s.sellAmount / s.sellQty : 0;
    // 持有判斷：未實現有市值或交易明細推算還有股數
    s.holding = s.marketValue > 0 || s.currentQty > 0;
  }

  return [...stocks.values()].sort((a, b) => {
    // 持有中的優先，再按實際損益由大到小
    if (a.holding !== b.holding) return a.holding ? -1 : 1;
    return b.actualPL - a.actualPL;
  });
}

const _expandedStocks = new Set();

// ---------- 個股分析摘要卡片 ----------
function renderStockAnalyticsCards(stocks) {
  const wrap = document.getElementById('stockAnalyticsCards');
  if (!wrap) return;
  if (!stocks || !stocks.length) {
    wrap.innerHTML = '';
    return;
  }

  // 已實現過的股票（有買賣紀錄）
  const closed = stocks.filter(s => s.realizedCount > 0);
  const profitable = closed.filter(s => s.actualPL > 0);
  const losing = closed.filter(s => s.actualPL < 0);
  const winRate = closed.length > 0 ? (profitable.length / closed.length) * 100 : 0;

  // 平均單筆獲利/虧損（依股票算）
  const totalProfit = profitable.reduce((s, x) => s + x.actualPL, 0);
  const totalLoss = Math.abs(losing.reduce((s, x) => s + x.actualPL, 0));
  const avgProfit = profitable.length > 0 ? totalProfit / profitable.length : 0;
  const avgLoss = losing.length > 0 ? totalLoss / losing.length : 0;
  const profitFactor = totalLoss > 0 ? totalProfit / totalLoss : (totalProfit > 0 ? Infinity : 0);

  // 最賺/最賠的股票
  const sortedByPL = closed.slice().sort((a,b) => b.actualPL - a.actualPL);
  const bestStock = sortedByPL[0];
  const worstStock = sortedByPL[sortedByPL.length - 1];

  // 持股集中度（前 3 大占比）
  const holdings = stocks.filter(s => s.holding && s.marketValue > 0)
    .sort((a,b) => b.marketValue - a.marketValue);
  const totalMarket = holdings.reduce((s,x) => s + x.marketValue, 0);
  const top3Market = holdings.slice(0, 3).reduce((s,x) => s + x.marketValue, 0);
  const concentrationPct = totalMarket > 0 ? (top3Market / totalMarket) * 100 : 0;

  // 最常交易的股票（按進出總次數）
  const sortedByTrades = stocks.slice()
    .filter(s => s.buyCount + s.sellCount > 0)
    .sort((a,b) => (b.buyCount + b.sellCount) - (a.buyCount + a.sellCount));
  const mostTraded = sortedByTrades[0];

  // 累計利息成本占比（利息 / 總獲利）
  const totalInterest = stocks.reduce((s,x) => s + x.interest + x.shortFee, 0);

  const cards = [];

  if (closed.length > 0) {
    cards.push(`
      <div class="card">
        <div class="label">勝率</div>
        <div class="value ${winRate>=50?'pos':'neg'}">${winRate.toFixed(1)}%</div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">${profitable.length}/${closed.length} 檔獲利</div>
      </div>
    `);

    cards.push(`
      <div class="card">
        <div class="label">盈虧比</div>
        <div class="value ${profitFactor>=1?'pos':'neg'}">${profitFactor === Infinity ? '∞' : profitFactor.toFixed(2)}</div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">總獲利 ÷ 總虧損</div>
      </div>
    `);

    cards.push(`
      <div class="card">
        <div class="label">平均獲利 / 虧損</div>
        <div class="value" style="font-size:16px">
          <span class="pos">+${fmt(avgProfit)}</span>
          <span style="color:var(--text-muted)"> / </span>
          <span class="neg">-${fmt(avgLoss)}</span>
        </div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">${avgLoss>0 ? `風報比 1:${(avgProfit/avgLoss).toFixed(2)}` : ''}</div>
      </div>
    `);

    if (bestStock) {
      cards.push(`
        <div class="card">
          <div class="label">最賺的股票</div>
          <div class="value pos" style="font-size:16px">${bestStock.code} ${bestStock.name||''}</div>
          <div style="font-size:13px;color:var(--success);margin-top:2px;font-weight:600">+${fmt(bestStock.actualPL)}</div>
        </div>
      `);
    }
    if (worstStock && worstStock !== bestStock && worstStock.actualPL < 0) {
      cards.push(`
        <div class="card">
          <div class="label">最賠的股票</div>
          <div class="value neg" style="font-size:16px">${worstStock.code} ${worstStock.name||''}</div>
          <div style="font-size:13px;color:var(--danger);margin-top:2px;font-weight:600">${fmt(worstStock.actualPL,{sign:true})}</div>
        </div>
      `);
    }
  }

  if (holdings.length > 0) {
    cards.push(`
      <div class="card">
        <div class="label">持股集中度</div>
        <div class="value">${concentrationPct.toFixed(1)}%</div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">前 3 大占總市值（共 ${holdings.length} 檔）</div>
      </div>
    `);
  }

  if (mostTraded) {
    cards.push(`
      <div class="card">
        <div class="label">最常交易</div>
        <div class="value" style="font-size:16px">${mostTraded.code} ${mostTraded.name||''}</div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">買${mostTraded.buyCount}次/賣${mostTraded.sellCount}次</div>
      </div>
    `);
  }

  if (totalInterest > 0) {
    const realizedTotal = closed.reduce((s,x) => s + x.actualPL, 0);
    const interestRatio = realizedTotal !== 0 ? (totalInterest / Math.abs(realizedTotal)) * 100 : 0;
    cards.push(`
      <div class="card">
        <div class="label">累計利息成本</div>
        <div class="value">${fmt(totalInterest)}</div>
        <div style="font-size:11px;color:var(--text-muted);margin-top:2px">${interestRatio>0?`占實現損益 ${interestRatio.toFixed(1)}%`:''}</div>
      </div>
    `);
  }

  wrap.innerHTML = cards.join('');
}

function renderStockAnalysis() {
  const acc = getCurrentAccount();
  const tb = document.querySelector('#stockTable tbody');
  if (!tb) return;
  if (!acc) {
    tb.innerHTML = '<tr><td colspan="8" class="empty-state">尚未選擇帳戶</td></tr>';
    renderStockAnalyticsCards(null);
    return;
  }

  // 沒有任何資料時提示先匯入
  if (!acc.trades.length && !acc.realized.length && !acc.unrealized.length) {
    tb.innerHTML = '<tr><td colspan="8" class="empty-state">尚無資料。請先到「⬆️ 匯入資料」匯入投資明細、已實現損益或未實現損益</td></tr>';
    renderStockAnalyticsCards(null);
    return;
  }

  let stocks = buildStockAnalysis(acc);
  // 當沖統計（按代號）
  const dtByCode = aggregateDayTradesByCode(analyzeDayTrades(acc));

  renderStockAnalyticsCards(stocks);

  // 完全沒解析到任何股票
  if (!stocks.length) {
    tb.innerHTML = '<tr><td colspan="8" class="empty-state">資料解析後沒有股票（請檢查匯入的資料）</td></tr>';
    return;
  }

  const totalCount = stocks.length;
  const search = (document.getElementById('stockSearch')?.value || '').toLowerCase();
  const filter = document.getElementById('stockFilter')?.value || 'all';

  // 篩選
  stocks = stocks.filter(s => {
    if (search && !(s.code.toLowerCase().includes(search) || (s.name||'').toLowerCase().includes(search))) return false;
    if (filter === 'holding' && !s.holding) return false;
    if (filter === 'closed' && s.holding) return false;
    if (filter === 'profit' && s.actualPL <= 0) return false;
    if (filter === 'loss' && s.actualPL >= 0) return false;
    return true;
  });

  if (!stocks.length) {
    tb.innerHTML = `<tr><td colspan="8" class="empty-state">沒有符合條件的股票（總共 ${totalCount} 檔，請放寬搜尋或篩選）</td></tr>`;
    return;
  }

  const rows = [];
  for (const s of stocks) {
    const expanded = _expandedStocks.has(s.code);
    const holdingTag = s.holding ? `<span class="year-tag" style="background:var(--success-light);color:var(--success)">持有 ${fmt(s.currentQty)}</span>` : '';
    const dt = dtByCode.get(s.code);

    // 報酬率（用累計買進金額作為基數）
    let rateStr = '—';
    let rateClass = '';
    if (s.buyAmount > 0) {
      const r = (s.actualPL / s.buyAmount) * 100;
      rateStr = (r > 0 ? '+' : '') + r.toFixed(2) + '%';
      rateClass = plClass(r);
    }

    rows.push(`
      <tr class="month-row" data-code="${s.code}">
        <td><span class="month-toggle ${expanded?'expanded':''}">▶</span></td>
        <td><strong>${s.code}</strong></td>
        <td>${s.name||''} ${holdingTag}</td>
        <td>${s.holding ? fmt(s.currentQty) : '—'}</td>
        <td class="${plClass(s.realizedPL)}">${fmt(s.realizedPL,{sign:true})}</td>
        <td class="hl ${plClass(s.actualPL)}"><strong>${fmt(s.actualPL,{sign:true})}</strong></td>
        <td class="${rateClass}">${rateStr}</td>
        <td class="hl ${dt?plClass(dt.netPL):''}">${dt ? fmt(dt.netPL,{sign:true}) : '—'}</td>
      </tr>
    `);
    if (expanded) {
      rows.push(`<tr class="month-detail-row"><td colspan="8">${renderStockDetail(s, dt)}</td></tr>`);
    }
  }
  tb.innerHTML = rows.join('');

  tb.querySelectorAll('.month-row').forEach(row => {
    row.onclick = () => {
      const code = row.dataset.code;
      if (_expandedStocks.has(code)) _expandedStocks.delete(code); else _expandedStocks.add(code);
      renderStockAnalysis();
    };
  });
}

function renderStockDetail(s, dt) {
  const html = [];
  const interestSum = (s.interest || 0) + (s.shortFee || 0);
  const buyPrice = s.avgBuyPrice ? fmt(s.avgBuyPrice, {decimals:2}) : '—';
  const sellPrice = s.avgSellPrice ? fmt(s.avgSellPrice, {decimals:2}) : '—';

  // 上方：完整指標卡片（精簡主表沒顯示的欄位都放這裡）
  html.push(`
    <div class="stock-detail-panel">
      <div class="indicator">
        <div class="label">進出筆數</div>
        <div class="value">${s.buyCount} 買 / ${s.sellCount} 賣</div>
      </div>
      <div class="indicator">
        <div class="label">已實現次數</div>
        <div class="value">${s.realizedCount}</div>
      </div>
      <div class="indicator">
        <div class="label">平均買價</div>
        <div class="value">${buyPrice}</div>
      </div>
      <div class="indicator">
        <div class="label">平均賣價</div>
        <div class="value">${sellPrice}</div>
      </div>
      <div class="indicator">
        <div class="label">累計買進</div>
        <div class="value">${fmt(s.buyAmount)}</div>
      </div>
      <div class="indicator">
        <div class="label">累計賣出</div>
        <div class="value">${fmt(s.sellAmount)}</div>
      </div>
      <div class="indicator">
        <div class="label">融資利息</div>
        <div class="value">${fmt(s.interest||0)}</div>
      </div>
      <div class="indicator">
        <div class="label">融券手續費</div>
        <div class="value">${fmt(s.shortFee||0)}</div>
      </div>
      <div class="indicator">
        <div class="label">利息+融券費</div>
        <div class="value">${interestSum ? fmt(interestSum) : '—'}</div>
      </div>
      ${dt ? `
        <div class="indicator">
          <div class="label">當沖筆數</div>
          <div class="value">${dt.count}</div>
        </div>
        <div class="indicator">
          <div class="label">當沖勝率</div>
          <div class="value ${dt.winRate>=50?'pos':'neg'}">${dt.winRate.toFixed(1)}%</div>
        </div>
        <div class="indicator">
          <div class="label">當沖損益</div>
          <div class="value ${plClass(dt.netPL)}">${fmt(dt.netPL,{sign:true})}</div>
        </div>
      ` : ''}
    </div>
  `);

  // 該股已實現
  if (s.realizedItems.length > 0) {
    html.push('<div class="nested">');
    html.push('<h4 style="padding:10px 14px 0">💰 已實現紀錄（' + s.realizedItems.length + ' 筆）</h4>');
    html.push('<table class="loan-subtable"><thead><tr>');
    html.push('<th>賣出日</th><th>買進日</th><th>類別</th><th>數量</th><th>賣價</th><th>買價</th><th>盈虧</th><th>利息</th><th>調整</th><th>實際</th>');
    html.push('</tr></thead><tbody>');
    for (const r of s.realizedItems) {
      const adjT = adjustTotal(r);
      const actual = actualPL(r);
      html.push(`<tr>
        <td>${r.sellDate}</td>
        <td>${r.buyDate||'—'}</td>
        <td>${r.sellCategory||''}</td>
        <td>${fmt(r.qty)}</td>
        <td>${fmt(r.sellPrice,{decimals:2})}</td>
        <td>${fmt(r.buyPrice,{decimals:2})}</td>
        <td class="${plClass(r.pl)}">${fmt(r.pl,{sign:true})}</td>
        <td>${fmt(r.interest||0)}</td>
        <td class="${plClass(adjT)}">${adjT?fmt(adjT,{sign:true}):'—'}</td>
        <td class="${plClass(actual)}"><strong>${fmt(actual,{sign:true})}</strong></td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
  }

  // 該股交易紀錄
  if (s.trades.length > 0) {
    html.push('<div class="nested">');
    html.push('<h4 style="padding:10px 14px 0">📋 交易紀錄（' + s.trades.length + ' 筆）</h4>');
    html.push('<table class="loan-subtable"><thead><tr>');
    html.push('<th>日期</th><th>買賣</th><th>類別</th><th>數量</th><th>單價</th><th>價金</th><th>手續費</th><th>利息</th><th>備註</th>');
    html.push('</tr></thead><tbody>');
    for (const t of s.trades.slice().sort((a,b) => (a.date||'').localeCompare(b.date||''))) {
      const isConv = (t.category === '資轉現' || t.action === '資轉現');
      const noteCell = isConv ? (t.note || '資轉現') : '';
      html.push(`<tr ${isConv?'style="background:var(--primary-light)"':''}>
        <td>${t.date}</td>
        <td>${t.action||''}</td>
        <td>${t.category||''}</td>
        <td>${fmt(t.qty)}</td>
        <td>${fmt(t.price,{decimals:2})}</td>
        <td>${fmt(t.amount)}</td>
        <td>${fmt(t.fee)}</td>
        <td>${fmt(t.marginInterest||0)}</td>
        <td>${noteCell}</td>
      </tr>`);
    }
    html.push('</tbody></table></div>');
  }

  // 持有中市值
  if (s.holding) {
    html.push('<div class="nested" style="padding:14px">');
    html.push(`<h4>📦 目前持有</h4>`);
    html.push(`<p style="margin:6px 0;color:var(--text-dim)">數量 ${fmt(s.currentQty)} 股　|　市值 ${fmt(s.marketValue)}　|　成本 ${fmt(s.cost)}　|　未實現損益 <span class="${plClass(s.unrealizedPL)}"><strong>${fmt(s.unrealizedPL,{sign:true})}</strong></span></p>`);
    html.push('</div>');
  }

  if (html.length === 0) html.push('<div class="empty-state">該股無細項資料</div>');
  return html.join('');
}

function exportStocksExcel() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const stocks = buildStockAnalysis(acc);
  if (!stocks.length) return toast('沒有資料可匯出', 'err');

  const headers = ['代號','名稱','持有中股數','進出次數','已實現次數',
                   '平均買價','平均賣價',
                   '累計買進','累計賣出','累計實現損益','調整金額','累計利息',
                   '累計融券費','實際損益','平均報酬率(%)','市值','成本','未實現損益'];
  const rows = stocks.map(s => [
    s.code, s.name, s.currentQty,
    `${s.buyCount}/${s.sellCount}`, s.realizedCount,
    s.avgBuyPrice ? +s.avgBuyPrice.toFixed(2) : 0,
    s.avgSellPrice ? +s.avgSellPrice.toFixed(2) : 0,
    s.buyAmount, s.sellAmount, s.realizedPL, s.adjust, s.interest,
    s.shortFee, s.actualPL, s.avgRate.toFixed(2),
    s.marketValue, s.cost, s.unrealizedPL
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '個股損益分析');
  const ts = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `個股損益_${acc.name}_${ts}.xlsx`);
  toast('已匯出 Excel', 'ok');
}

// ============================================================
// 資轉現
// ============================================================

// 計算「目前還未賣出的融資買進」清單
function getOpenMarginBuys(acc) {
  // 把每筆融資買進記下尚未配對的剩餘數量
  // 用「先進先出」配對掉融資賣出和資轉現
  const buys = []; // {trade, remainingQty}
  const trades = (acc.trades || []).slice().sort((a,b) => (a.date||'').localeCompare(b.date||''));

  for (const t of trades) {
    if (t.category === '融資' && t.action === '買') {
      buys.push({ trade: t, remainingQty: t.qty || 0 });
    } else if (t.category === '融資' && t.action === '賣') {
      // 從最早的融資買進開始扣
      let need = t.qty || 0;
      for (const b of buys) {
        if (need <= 0) break;
        if (b.remainingQty <= 0) continue;
        if (b.trade.code !== t.code) continue;
        const used = Math.min(b.remainingQty, need);
        b.remainingQty -= used;
        need -= used;
      }
    } else if ((t.category === '資轉現' || t.action === '資轉現') && t._convertedFromMarginBuy) {
      // 從指定的融資買進扣掉
      let need = t.qty || 0;
      for (const b of buys) {
        if (need <= 0) break;
        if (b.remainingQty <= 0) continue;
        if (b.trade.code !== t.code) continue;
        // 嚴格比對：資轉現紀錄會帶 _convertedFromMarginBuy = trade key
        if (tradeKey(b.trade) !== t._convertedFromMarginBuy) continue;
        const used = Math.min(b.remainingQty, need);
        b.remainingQty -= used;
        need -= used;
      }
    }
  }

  return buys.filter(b => b.remainingQty > 0);
}

function tradeKey(t) {
  return `${t.date}|${t.code}|${t.action}|${t.category}|${t.qty}|${t.price}`;
}

async function showConvertToCashDialog() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');

  const open = getOpenMarginBuys(acc);
  if (!open.length) {
    return toast('目前沒有未平倉的融資買進可轉現', 'err');
  }

  const today = todayStr().replace(/\//g, '-');
  const optionsHtml = open.map((b, i) => {
    const t = b.trade;
    return `<option value="${i}">${t.date}　${t.code} ${t.name||''}　${fmt(b.remainingQty)}股 @${fmt(t.price,{decimals:2})}　成本約 ${fmt(b.remainingQty * t.price)}</option>`;
  }).join('');

  const result = await showModal({
    title: '🔄 新增資轉現',
    html: `
      <div style="display:grid;gap:10px;min-width:480px">
        <label>從哪筆融資買進轉？
          <select id="cv-source" style="width:100%">${optionsHtml}</select>
        </label>
        <label>資轉現日期
          <input type="date" id="cv-date" value="${today}" style="width:100%">
        </label>
        <label>結清融資利息（券商計算）
          <input type="number" id="cv-interest" placeholder="例：320" style="width:100%">
        </label>
        <label>備註
          <input type="text" id="cv-note" placeholder="可不填" style="width:100%">
        </label>
        <p class="hint" style="margin:0;font-size:12px">系統會自動產生兩筆紀錄：「融資結清」+「現股建倉」（成本沿用原融資買進）</p>
      </div>
    `,
    onConfirm: (body) => ({
      sourceIdx: parseInt(body.querySelector('#cv-source').value, 10),
      date: body.querySelector('#cv-date').value,
      interest: parseFloat(body.querySelector('#cv-interest').value) || 0,
      note: body.querySelector('#cv-note').value.trim()
    })
  });
  if (!result) return;

  const src = open[result.sourceIdx];
  if (!src) return toast('來源無效', 'err');

  const sourceTrade = src.trade;
  const qty = src.remainingQty; // 全數轉現（如果你想部分轉，未來再加數量輸入）
  const date = Parsers.formatDate(result.date);
  const sourceKey = tradeKey(sourceTrade);

  // 產生「資轉現-融資結清」紀錄
  const closeRec = {
    date,
    code: sourceTrade.code,
    name: sourceTrade.name,
    kind: '整股',
    action: '資轉現',
    category: '資轉現',
    rawCategory: '資轉現-融資結清',
    qty,
    price: sourceTrade.price,
    amount: qty * sourceTrade.price,
    fee: 0,
    tax: 0,
    receivable: 0,
    marginAmount: sourceTrade.marginAmount || 0,
    ownFund: 0,
    marginInterest: result.interest,  // 結清利息
    shortFee: 0,
    borrowFee: 0,
    interestTax: 0,
    nhi: 0,
    pl: 0,
    settleDate: '',
    currency: sourceTrade.currency || '新台幣',
    note: result.note ? `資轉現-融資結清｜${result.note}` : '資轉現-融資結清',
    _convertedFromMarginBuy: sourceKey,  // 連結到原融資買進
    _convertId: 'CV-' + Date.now()
  };

  // 產生「資轉現-現股建倉」紀錄
  const openCashRec = {
    date,
    code: sourceTrade.code,
    name: sourceTrade.name,
    kind: '整股',
    action: '資轉現',
    category: '資轉現',
    rawCategory: '資轉現-現股建倉',
    qty,
    price: sourceTrade.price,  // 沿用原融資買進的單價
    amount: qty * sourceTrade.price,
    fee: 0,
    tax: 0,
    receivable: 0,
    marginAmount: 0,
    ownFund: 0,
    marginInterest: 0,
    shortFee: 0,
    borrowFee: 0,
    interestTax: 0,
    nhi: 0,
    pl: 0,
    settleDate: '',
    currency: sourceTrade.currency || '新台幣',
    note: result.note ? `資轉現-現股建倉｜${result.note}` : '資轉現-現股建倉',
    _convertedFromMarginBuy: sourceKey,
    _convertId: closeRec._convertId
  };

  acc.trades.push(closeRec);
  acc.trades.push(openCashRec);
  acc.trades.sort((a,b) => (a.date||'').localeCompare(b.date||''));

  save();
  renderAll();
  toast(`資轉現已登記：${sourceTrade.code} ${sourceTrade.name||''} ${fmt(qty)}股`, 'ok');
}



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

// ============================================================
// 融資回補（追繳）
// ============================================================

function mcGenId(acc) {
  const existing = (acc.marginCalls || []).map(m => m.id);
  let n = existing.length + 1;
  while (existing.includes('MC-' + String(n).padStart(3,'0'))) n++;
  return 'MC-' + String(n).padStart(3, '0');
}

function renderMarginCalls() {
  const acc = getCurrentAccount();
  const list = document.getElementById('marginCallList');
  const nameSpan = document.getElementById('mcAccountName');
  if (!list) return;

  if (!acc) {
    if (nameSpan) nameSpan.textContent = '—';
    list.innerHTML = '<div class="empty-state">尚未選擇帳戶</div>';
    ['mcTotalIn','mcTotalOut','mcRemaining','mcActiveCount'].forEach(id => setVal(id, '—'));
    return;
  }
  if (nameSpan) nameSpan.textContent = acc.name;

  const mcs = acc.marginCalls || [];
  const total = mcs.reduce((s, m) => s + (m.amount || 0), 0);
  const paid = mcs.reduce((s, m) =>
    s + (m.payouts || []).reduce((ss, p) => ss + (p.amount || 0), 0), 0);
  const remaining = total - paid;
  const active = mcs.filter(m => {
    const p = (m.payouts || []).reduce((s, p) => s + (p.amount || 0), 0);
    return p < (m.amount || 0);
  }).length;

  setVal('mcTotalIn', total);
  setVal('mcTotalOut', paid);
  setVal('mcRemaining', remaining);
  setVal('mcActiveCount', active);

  if (!mcs.length) {
    list.innerHTML = '<div class="empty-state">尚無融資回補紀錄。點上方「＋ 新增回補」開始記錄。</div>';
    return;
  }

  // 排序：未領回的優先、再依日期遞減
  const sorted = mcs.slice().sort((a, b) => {
    const aPaid = (a.payouts || []).reduce((s, p) => s + (p.amount || 0), 0);
    const bPaid = (b.payouts || []).reduce((s, p) => s + (p.amount || 0), 0);
    const aActive = aPaid < (a.amount || 0);
    const bActive = bPaid < (b.amount || 0);
    if (aActive !== bActive) return aActive ? -1 : 1;
    return (b.date || '').localeCompare(a.date || '');
  });

  list.innerHTML = sorted.map(mc => renderMarginCallCard(mc)).join('');
  bindMarginCallEvents();
}

function renderMarginCallCard(mc) {
  const total = mc.amount || 0;
  const paid = (mc.payouts || []).reduce((s, p) => s + (p.amount || 0), 0);
  const remaining = total - paid;
  const fullyPaid = remaining <= 0;
  const status = fullyPaid ? '已全額領回' : (paid > 0 ? '部分領回' : '未領回');

  let payoutRows = (mc.payouts || []).slice().reverse().map((p, i) => `
    <tr>
      <td>${p.date}</td>
      <td>${fmt(p.amount)}</td>
      <td>${p.note || ''}</td>
      <td><button class="btn-mini danger" data-act="del-payout" data-mc="${mc.id}" data-idx="${(mc.payouts||[]).length-1-i}" style="padding:2px 6px;font-size:11px">刪</button></td>
    </tr>
  `).join('') || '<tr><td colspan="4" class="empty-state">尚未領回</td></tr>';

  return `
    <div class="loan-card">
      <div class="loan-card-header ${fullyPaid?'settled':''}">
        <span class="loan-id">${mc.id}</span>
        <span class="loan-purpose">${mc.reason || '（無說明）'}</span>
        <span class="loan-status ${fullyPaid?'settled':''}">${status}</span>
      </div>
      <div class="loan-card-body">
        <div class="loan-fields">
          <div class="loan-field"><div class="label">回補日期</div><div class="value">${mc.date || '—'}</div></div>
          <div class="loan-field"><div class="label">回補金額</div><div class="value">${fmt(total)}</div></div>
          <div class="loan-field"><div class="label">維持率</div><div class="value">${mc.ratio ? mc.ratio + '%' : '—'}</div></div>
          <div class="loan-field"><div class="label">已領回</div><div class="value">${fmt(paid)}</div></div>
          <div class="loan-field"><div class="label">未領回</div><div class="value ${remaining>0?'neg':'pos'}">${fmt(remaining)}</div></div>
          ${mc.note ? `<div class="loan-field" style="grid-column:1/-1"><div class="label">備註</div><div class="value" style="font-size:13px;font-weight:400">${mc.note}</div></div>` : ''}
        </div>

        <div class="loan-subsection">
          <h4>💰 領回紀錄</h4>
          <table class="loan-subtable">
            <thead><tr><th>日期</th><th>金額</th><th>備註</th><th></th></tr></thead>
            <tbody>${payoutRows}</tbody>
          </table>
        </div>

        <div class="loan-actions">
          ${!fullyPaid ? `<button class="btn-mini primary" data-act="add-payout" data-mc="${mc.id}">＋ 新增領回</button>` : ''}
          <button class="btn-mini" data-act="edit" data-mc="${mc.id}">編輯</button>
          <button class="btn-mini danger" data-act="delete" data-mc="${mc.id}">刪除回補</button>
        </div>
      </div>
    </div>
  `;
}

function bindMarginCallEvents() {
  document.querySelectorAll('#marginCallList button[data-act]').forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      const act = btn.dataset.act;
      const mcId = btn.dataset.mc;
      const acc = getCurrentAccount();
      const mc = acc.marginCalls.find(x => x.id === mcId);
      if (!mc) return;

      if (act === 'add-payout') await addMarginCallPayout(mc);
      else if (act === 'edit') await editMarginCall(mc);
      else if (act === 'delete') {
        const ok = await confirmDialog('刪除回補', `確定刪除「${mc.id}」？所有領回紀錄會一併消失。`);
        if (ok) {
          acc.marginCalls = acc.marginCalls.filter(x => x.id !== mcId);
          save(); renderAll();
          toast('已刪除', 'ok');
        }
      }
      else if (act === 'del-payout') {
        const idx = parseInt(btn.dataset.idx, 10);
        mc.payouts.splice(idx, 1);
        save(); renderAll();
      }
    };
  });
}

async function addMarginCall() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const today = todayStr().replace(/\//g, '-');
  const result = await showModal({
    title: '⚠️ 新增融資回補',
    html: `
      <div style="display:grid;gap:10px;min-width:380px">
        <label>回補日期<input type="date" id="mc-date" value="${today}"></label>
        <label>回補金額<input type="number" id="mc-amount" placeholder="例：100000"></label>
        <label>維持率(%)（可選）<input type="number" id="mc-ratio" placeholder="例：125" step="0.01"></label>
        <label>觸發原因<input type="text" id="mc-reason" value="跌破130%維持率" placeholder="例：跌破130%維持率、追繳"></label>
        <label>備註<input type="text" id="mc-note" placeholder="可不填"></label>
      </div>
    `,
    onConfirm: (body) => ({
      date: body.querySelector('#mc-date').value,
      amount: parseFloat(body.querySelector('#mc-amount').value) || 0,
      ratio: parseFloat(body.querySelector('#mc-ratio').value) || 0,
      reason: body.querySelector('#mc-reason').value.trim(),
      note: body.querySelector('#mc-note').value.trim()
    })
  });
  if (!result || !result.amount) return;

  const mc = {
    id: mcGenId(acc),
    date: Parsers.formatDate(result.date),
    amount: result.amount,
    ratio: result.ratio || null,
    reason: result.reason || '',
    note: result.note || '',
    payouts: [],
    createdAt: new Date().toISOString()
  };
  acc.marginCalls.push(mc);
  save();
  renderAll();
  toast(`已新增回補 ${mc.id}`, 'ok');
}

async function editMarginCall(mc) {
  const result = await showModal({
    title: `編輯回補 ${mc.id}`,
    html: `
      <div style="display:grid;gap:10px;min-width:380px">
        <label>回補日期<input type="date" id="mc-date" value="${(mc.date||'').replace(/\//g,'-')}"></label>
        <label>回補金額<input type="number" id="mc-amount" value="${mc.amount||0}"></label>
        <label>維持率(%)<input type="number" id="mc-ratio" value="${mc.ratio||''}" step="0.01"></label>
        <label>觸發原因<input type="text" id="mc-reason" value="${(mc.reason||'').replace(/"/g,'&quot;')}"></label>
        <label>備註<input type="text" id="mc-note" value="${(mc.note||'').replace(/"/g,'&quot;')}"></label>
      </div>
    `,
    onConfirm: (body) => ({
      date: body.querySelector('#mc-date').value,
      amount: parseFloat(body.querySelector('#mc-amount').value) || 0,
      ratio: parseFloat(body.querySelector('#mc-ratio').value) || 0,
      reason: body.querySelector('#mc-reason').value.trim(),
      note: body.querySelector('#mc-note').value.trim()
    })
  });
  if (!result) return;
  mc.date = Parsers.formatDate(result.date);
  mc.amount = result.amount;
  mc.ratio = result.ratio || null;
  mc.reason = result.reason;
  mc.note = result.note;
  save();
  renderAll();
  toast('已更新', 'ok');
}

async function addMarginCallPayout(mc) {
  const today = todayStr().replace(/\//g, '-');
  const paid = (mc.payouts || []).reduce((s, p) => s + (p.amount || 0), 0);
  const remaining = (mc.amount || 0) - paid;

  const result = await showModal({
    title: `新增領回 - ${mc.id}`,
    html: `
      <div style="display:grid;gap:10px;min-width:380px">
        <label>日期<input type="date" id="po-date" value="${today}"></label>
        <label>領回金額（剩餘 ${fmt(remaining)}）<input type="number" id="po-amount"></label>
        <label>備註<input type="text" id="po-note" placeholder="可不填"></label>
        <label><input type="checkbox" id="po-full"> 全額領回（自動填入剩餘金額）</label>
      </div>
    `,
    onConfirm: (body) => {
      const full = body.querySelector('#po-full').checked;
      return {
        date: body.querySelector('#po-date').value,
        amount: full ? remaining : (parseFloat(body.querySelector('#po-amount').value) || 0),
        note: body.querySelector('#po-note').value.trim()
      };
    }
  });
  if (!result || !result.amount) return;

  if (!mc.payouts) mc.payouts = [];
  mc.payouts.push({
    date: Parsers.formatDate(result.date),
    amount: result.amount,
    note: result.note
  });
  mc.payouts.sort((a, b) => a.date.localeCompare(b.date));
  save();
  renderAll();
  toast('已新增領回', 'ok');
}

// ============================================================
// 股利紀錄
// ============================================================

function divGenId() {
  return 'D-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6);
}

function renderDividends() {
  const acc = getCurrentAccount();
  const list = document.getElementById('dividendList');
  if (!list) return;

  if (!acc) {
    list.innerHTML = '<div class="empty-state">尚未選擇帳戶</div>';
    ['divTotalCash','divTotalStock','divEntryCount','divStockCount'].forEach(id => setVal(id, '—'));
    return;
  }

  const div = acc.dividends || { entries: [] };
  const byCode = computeDividendByCode(acc);
  const totalCash = byCode.reduce((s, x) => s + x.cash, 0);
  const totalStock = byCode.reduce((s, x) => s + x.stock, 0);

  setVal('divTotalCash', totalCash);
  setVal('divTotalStock', totalStock);
  setVal('divEntryCount', div.entries.length);
  setVal('divStockCount', byCode.length);

  if (!byCode.length && !div.entries.length) {
    list.innerHTML = '<div class="empty-state">尚無股利紀錄。請從「⬆️ 匯入資料」上傳除權息總覽檔，或點上方「＋ 新增逐筆紀錄」手動加入。</div>';
    return;
  }

  const mode = document.getElementById('divViewMode')?.value || 'byStock';
  if (mode === 'byStock') list.innerHTML = renderDividendsByStock(byCode);
  else if (mode === 'byEntry') list.innerHTML = renderDividendsByEntry(div.entries);
  else if (mode === 'byYear') list.innerHTML = renderDividendsByYear(div.entries);

  bindDividendEvents();
}

function renderDividendsByStock(byCode) {
  if (!byCode.length) return '<div class="empty-state">無資料</div>';
  const rows = byCode.slice().sort((a, b) => b.cash - a.cash).map(s => {
    return `
      <tr class="dividend-stock-row" data-code="${s.code}">
        <td><strong>${s.code}</strong></td>
        <td>${s.name}</td>
        <td class="pos">${fmt(s.cash)}</td>
        <td>${s.stock ? fmt(s.stock) : '—'}</td>
        <td>${fmt(s.tax)}</td>
        <td>${s.count} 筆</td>
        <td>
          <button class="btn-mini" data-act="add-entry" data-code="${s.code}" data-name="${s.name.replace(/"/g,'&quot;')}">＋ 新增逐筆</button>
          <button class="btn-mini danger" data-act="del-stock" data-code="${s.code}">刪除全部</button>
        </td>
      </tr>
    `;
  }).join('');

  // 展開逐筆紀錄
  const detailRows = byCode
    .filter(s => s.entries.length > 0)
    .map(s => {
      const entries = s.entries.slice().sort((a, b) => (b.payDate || b.exDate || '').localeCompare(a.payDate || a.exDate || ''));
      const inner = entries.map(e => `
        <tr>
          <td>${e.exDate || '—'}</td>
          <td>${e.payDate || '—'}</td>
          <td>${e.note || ''}</td>
          <td class="pos">${fmt(e.cash || 0)}</td>
          <td>${e.stockShares ? fmt(e.stockShares) : '—'}</td>
          <td>${fmt(e.tax || 0)}</td>
          <td>
            <button class="btn-mini" data-act="edit-entry" data-id="${e.id}">編輯</button>
            <button class="btn-mini danger" data-act="del-entry" data-id="${e.id}">刪</button>
          </td>
        </tr>
      `).join('');
      return `
        <tr class="month-detail-row"><td colspan="7">
          <div class="nested" style="padding:14px">
            <h4>${s.code} ${s.name} - 逐筆紀錄（${entries.length} 筆，合計現金 ${fmt(s.cash)}）</h4>
            <table class="dividend-entry-table">
              <thead><tr><th>除息日</th><th>發放日</th><th>備註</th><th>現金</th><th>股票股數</th><th>扣繳稅</th><th></th></tr></thead>
              <tbody>${inner}</tbody>
            </table>
          </div>
        </td></tr>
      `;
    }).join('');

  return `
    <div class="table-scroll">
      <table class="data-table">
        <thead><tr>
          <th>代號</th><th>名稱</th><th>現金股利合計</th><th>股票股數</th><th>扣繳稅</th><th>筆數</th><th>操作</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    ${detailRows ? `<div style="margin-top:18px"><h3>逐筆紀錄展開</h3><table class="data-table"><tbody>${detailRows}</tbody></table></div>` : ''}
  `;
}

function renderDividendsByEntry(entries) {
  if (!entries.length) {
    return '<div class="empty-state">尚無逐筆紀錄。從「以股票檢視」點某檔的「＋ 新增逐筆」即可加入。</div>';
  }
  const sorted = entries.slice().sort((a, b) => (b.payDate || b.exDate || '').localeCompare(a.payDate || a.exDate || ''));
  const rows = sorted.map(e => `
    <tr>
      <td>${e.exDate || '—'}</td>
      <td>${e.payDate || '—'}</td>
      <td>${e.code}</td>
      <td>${e.name || ''}</td>
      <td class="pos">${fmt(e.cash || 0)}</td>
      <td>${e.stockShares ? fmt(e.stockShares) : '—'}</td>
      <td>${fmt(e.tax || 0)}</td>
      <td>${e.note || ''}</td>
      <td>
        <button class="btn-mini" data-act="edit-entry" data-id="${e.id}">編輯</button>
        <button class="btn-mini danger" data-act="del-entry" data-id="${e.id}">刪</button>
      </td>
    </tr>
  `).join('');
  return `
    <div class="table-scroll">
      <table class="data-table">
        <thead><tr>
          <th>除息日</th><th>發放日</th><th>代號</th><th>名稱</th><th>現金</th><th>股票股數</th><th>扣繳稅</th><th>備註</th><th>操作</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderDividendsByYear(entries) {
  // 把逐筆按年份合計
  const byYear = new Map();
  const ensure = (y) => {
    if (!byYear.has(y)) byYear.set(y, { year: y, cash: 0, stock: 0, tax: 0, count: 0, codes: new Set() });
    return byYear.get(y);
  };
  for (const e of entries) {
    const d = e.payDate || e.exDate || '';
    const y = (d.match(/^(\d{4})/) || [])[1] || '無日期';
    const r = ensure(y);
    r.cash += (e.cash || 0);
    r.stock += (e.stockShares || 0);
    r.tax += (e.tax || 0);
    r.count++;
    r.codes.add(e.code);
  }

  if (!byYear.size) return '<div class="empty-state">無資料</div>';

  const sorted = [...byYear.values()].sort((a, b) => {
    if (a.year === '無日期') return 1;
    if (b.year === '無日期') return -1;
    return b.year.localeCompare(a.year);
  });
  const rows = sorted.map(y => `
    <tr>
      <td><strong>${y.year}</strong></td>
      <td>${y.codes.size} 檔</td>
      <td>${y.count} 筆</td>
      <td class="pos">${fmt(y.cash)}</td>
      <td>${y.stock ? fmt(y.stock) : '—'}</td>
      <td>${fmt(y.tax)}</td>
    </tr>
  `).join('');
  return `
    <div class="table-scroll">
      <table class="data-table">
        <thead><tr>
          <th>年度</th><th>配息檔數</th><th>筆數</th><th>現金股利</th><th>股票股利</th><th>扣繳稅</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function bindDividendEvents() {
  document.querySelectorAll('#dividendList button[data-act], #dividendList .dividend-stock-row').forEach(el => {
    if (el.tagName === 'TR') {
      // 點 row 不做事（避免和按鈕衝突）
      return;
    }
    el.onclick = async (e) => {
      e.stopPropagation();
      const acc = getCurrentAccount();
      const act = el.dataset.act;

      if (act === 'add-entry') {
        await addDividendEntry(el.dataset.code, el.dataset.name);
      } else if (act === 'edit-entry') {
        const entry = acc.dividends.entries.find(x => x.id === el.dataset.id);
        if (entry) await editDividendEntry(entry);
      } else if (act === 'del-entry') {
        const ok = await confirmDialog('刪除逐筆紀錄', '確定刪除這筆股利紀錄？');
        if (ok) {
          acc.dividends.entries = acc.dividends.entries.filter(x => x.id !== el.dataset.id);
          save(); renderAll();
        }
      } else if (act === 'del-stock') {
        const ok = await confirmDialog('刪除股票所有股利紀錄', `確定刪除「${el.dataset.code}」的所有股利紀錄（含所有逐筆）？`);
        if (ok) {
          acc.dividends.entries = acc.dividends.entries.filter(x => x.code !== el.dataset.code);
          save(); renderAll();
        }
      }
    };
  });
}

async function addDividendEntry(presetCode = '', presetName = '') {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');

  const today = todayStr().replace(/\//g, '-');
  const result = await showModal({
    title: '新增逐筆股利紀錄',
    html: `
      <div style="display:grid;gap:10px;min-width:380px">
        <label>股票代號<input type="text" id="d-code" value="${presetCode}" placeholder="例：2330"></label>
        <label>股票名稱<input type="text" id="d-name" value="${presetName.replace(/"/g,'&quot;')}" placeholder="例：台積電"></label>
        <label>除息日（除權日）<input type="date" id="d-exdate"></label>
        <label>發放日<input type="date" id="d-paydate" value="${today}"></label>
        <label>現金股利<input type="number" id="d-cash" placeholder="例：3000" step="0.01"></label>
        <label>股票股利股數（無則填 0）<input type="number" id="d-stock" value="0"></label>
        <label>扣繳稅款<input type="number" id="d-tax" value="0"></label>
        <label>備註<input type="text" id="d-note" placeholder="例：第二季配息"></label>
      </div>
    `,
    onConfirm: (body) => ({
      code: body.querySelector('#d-code').value.trim(),
      name: body.querySelector('#d-name').value.trim(),
      exDate: body.querySelector('#d-exdate').value,
      payDate: body.querySelector('#d-paydate').value,
      cash: parseFloat(body.querySelector('#d-cash').value) || 0,
      stockShares: parseFloat(body.querySelector('#d-stock').value) || 0,
      tax: parseFloat(body.querySelector('#d-tax').value) || 0,
      note: body.querySelector('#d-note').value.trim()
    })
  });
  if (!result || !result.code) return;

  acc.dividends.entries.push({
    id: divGenId(),
    code: result.code,
    name: result.name,
    exDate: result.exDate ? Parsers.formatDate(result.exDate) : '',
    payDate: result.payDate ? Parsers.formatDate(result.payDate) : '',
    cash: result.cash,
    stockShares: result.stockShares,
    tax: result.tax,
    note: result.note
  });
  save();
  renderAll();
  toast('已新增逐筆紀錄', 'ok');
}

async function editDividendEntry(entry) {
  const result = await showModal({
    title: '編輯股利紀錄',
    html: `
      <div style="display:grid;gap:10px;min-width:380px">
        <label>股票代號<input type="text" id="d-code" value="${entry.code}"></label>
        <label>股票名稱<input type="text" id="d-name" value="${(entry.name||'').replace(/"/g,'&quot;')}"></label>
        <label>除息日<input type="date" id="d-exdate" value="${(entry.exDate||'').replace(/\//g,'-')}"></label>
        <label>發放日<input type="date" id="d-paydate" value="${(entry.payDate||'').replace(/\//g,'-')}"></label>
        <label>現金股利<input type="number" id="d-cash" value="${entry.cash||0}" step="0.01"></label>
        <label>股票股利股數<input type="number" id="d-stock" value="${entry.stockShares||0}"></label>
        <label>扣繳稅款<input type="number" id="d-tax" value="${entry.tax||0}"></label>
        <label>備註<input type="text" id="d-note" value="${(entry.note||'').replace(/"/g,'&quot;')}"></label>
      </div>
    `,
    onConfirm: (body) => ({
      code: body.querySelector('#d-code').value.trim(),
      name: body.querySelector('#d-name').value.trim(),
      exDate: body.querySelector('#d-exdate').value,
      payDate: body.querySelector('#d-paydate').value,
      cash: parseFloat(body.querySelector('#d-cash').value) || 0,
      stockShares: parseFloat(body.querySelector('#d-stock').value) || 0,
      tax: parseFloat(body.querySelector('#d-tax').value) || 0,
      note: body.querySelector('#d-note').value.trim()
    })
  });
  if (!result) return;
  Object.assign(entry, {
    code: result.code,
    name: result.name,
    exDate: result.exDate ? Parsers.formatDate(result.exDate) : '',
    payDate: result.payDate ? Parsers.formatDate(result.payDate) : '',
    cash: result.cash,
    stockShares: result.stockShares,
    tax: result.tax,
    note: result.note
  });
  save();
  renderAll();
  toast('已更新', 'ok');
}

async function importDividendFile(file) {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');

  let result;
  try {
    result = await Parsers.parseDividendsFile(file);
  } catch (e) {
    console.error(e);
    return toast('解析失敗：' + e.message, 'err');
  }
  if (!result.items || !result.items.length) {
    return toast('檔案內沒有可解析的股利資料', 'err');
  }

  // 跳對話框讓使用者填統一的除息日
  const today = todayStr().replace(/\//g, '-');
  const previewRows = result.items.slice(0, 6).map(x =>
    `<tr><td>${x.code}</td><td>${x.name}</td><td>${fmt(x.cashTotal)}</td><td>${x.stockTotal||0}</td></tr>`
  ).join('');
  const moreText = result.items.length > 6 ? `<tr><td colspan="4" style="text-align:center;color:var(--text-muted)">…還有 ${result.items.length - 6} 檔</td></tr>` : '';

  const confirmed = await showModal({
    title: `匯入股利紀錄（${result.items.length} 檔）`,
    html: `
      <div style="min-width:500px;max-width:640px">
        <p class="hint" style="margin-top:0">每一檔會新增一筆「逐筆紀錄」追加到目前帳戶。</p>

        <label style="display:block;margin-bottom:10px">
          <div style="font-weight:600;font-size:13px;margin-bottom:4px">除息日（必填）</div>
          <input type="date" id="div-exdate" value="${today}" style="width:100%">
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px">所有匯入的紀錄都會用這個除息日</div>
        </label>

        <label style="display:block;margin-bottom:10px">
          <div style="font-weight:600;font-size:13px;margin-bottom:4px">發放日（可選）</div>
          <input type="date" id="div-paydate" style="width:100%">
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px">不填的話會自動使用除息日</div>
        </label>

        <label style="display:block;margin-bottom:14px">
          <div style="font-weight:600;font-size:13px;margin-bottom:4px">備註（可選，所有筆共用）</div>
          <input type="text" id="div-note" placeholder="例：2024 年度配息" style="width:100%">
        </label>

        <div style="background:var(--surface-2);padding:10px 12px;border-radius:6px;font-size:12px">
          <div style="font-weight:600;margin-bottom:6px">預覽（前 6 筆）</div>
          <table class="alert-table">
            <thead><tr><th>代號</th><th>名稱</th><th>現金股利</th><th>股票股數</th></tr></thead>
            <tbody>${previewRows}${moreText}</tbody>
          </table>
        </div>
      </div>
    `,
    confirmText: '匯入',
    onConfirm: (body) => ({
      exDate: body.querySelector('#div-exdate').value,
      payDate: body.querySelector('#div-paydate').value,
      note: body.querySelector('#div-note').value.trim()
    })
  });

  if (!confirmed) return;
  if (!confirmed.exDate) return toast('請填除息日', 'err');

  const exDate = Parsers.formatDate(confirmed.exDate);
  const payDate = confirmed.payDate ? Parsers.formatDate(confirmed.payDate) : exDate;
  const note = confirmed.note;

  let added = 0;
  for (const x of result.items) {
    acc.dividends.entries.push({
      id: 'D-' + Date.now() + '-' + Math.random().toString(36).slice(2, 6) + '-' + added,
      code: x.code,
      name: x.name,
      exDate,
      payDate,
      cash: x.cashTotal || 0,
      stockShares: x.stockTotal || 0,
      tax: 0,
      note
    });
    added++;
  }
  save();
  renderAll();
  toast(`已追加 ${added} 筆股利紀錄（除息日 ${exDate}）`, 'ok');
}

function exportDividendsExcel() {
  const acc = getCurrentAccount();
  if (!acc) return toast('請先選擇帳戶', 'err');
  const byCode = computeDividendByCode(acc);
  if (!byCode.length) return toast('沒有資料可匯出', 'err');

  const wb = XLSX.utils.book_new();

  // Sheet 1：以股票檢視
  const h1 = ['代號','名稱','模式','現金股利合計','股票股利股數','逐筆紀錄筆數'];
  const r1 = byCode.map(s => [
    s.code, s.name, s.mode === 'detailed' ? '逐筆' : '累計',
    s.cash, s.stock, s.entries.length
  ]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h1, ...r1]), '以股票');

  // Sheet 2：逐筆紀錄
  const entries = acc.dividends.entries;
  if (entries.length > 0) {
    const h2 = ['除息日','發放日','代號','名稱','現金股利','股票股數','扣繳稅','備註'];
    const r2 = entries.slice().sort((a,b) => (b.payDate||b.exDate||'').localeCompare(a.payDate||a.exDate||''))
      .map(e => [e.exDate||'', e.payDate||'', e.code, e.name||'', e.cash||0, e.stockShares||0, e.tax||0, e.note||'']);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([h2, ...r2]), '逐筆紀錄');
  }

  const ts = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `股利紀錄_${acc.name}_${ts}.xlsx`);
  toast('已匯出 Excel', 'ok');
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
    tb.innerHTML = '<tr><td colspan="17" class="empty-state">尚無投資明細</td></tr>';
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

  tb.innerHTML = filtered.map((t, i) => {
    const realIdx = acc.trades.indexOf(t);
    const isConv = (t.category === '資轉現' || t.action === '資轉現');
    const rowStyle = isConv ? 'style="background:rgba(37,99,235,0.06)"' : '';
    const noteCell = t.note ? `<span style="color:var(--primary-dark);font-size:12px">${t.note}</span>` : '';

    // 操作按鈕
    let actions = '';
    if (isConv && t._convertId) {
      // 資轉現紀錄：只給「刪除整組」按鈕
      actions = `<button class="btn-mini danger" data-act="del-conv" data-cvid="${t._convertId}" style="padding:2px 6px;font-size:11px" title="刪除整組（含對應另一筆）">🗑️整組</button>`;
    } else {
      actions = `
        <button class="btn-mini" data-act="edit-trade" data-idx="${realIdx}" style="padding:2px 6px;font-size:11px" title="編輯">✏️</button>
        <button class="btn-mini danger" data-act="del-trade" data-idx="${realIdx}" style="padding:2px 6px;font-size:11px" title="刪除">🗑️</button>
      `;
    }

    return `
    <tr ${rowStyle}>
      <td>${t.date}</td>
      <td>${t.code}</td>
      <td>${t.name||''}</td>
      <td>${t.action||''}</td>
      <td>${isConv ? `<span class="year-tag" style="background:var(--primary-light);color:var(--primary-dark)">${t.rawCategory||t.category}</span>` : (t.category||'')}</td>
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
      <td>${t.settleDate || noteCell || '—'}</td>
      <td><div style="display:flex;gap:3px">${actions}</div></td>
    </tr>
    `;
  }).join('') || '<tr><td colspan="17" class="empty-state">沒有符合條件的資料</td></tr>';

  tb.querySelectorAll('button[data-act]').forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      const act = btn.dataset.act;

      if (act === 'del-conv') {
        const cvid = btn.dataset.cvid;
        const ok = await confirmDialog('刪除資轉現紀錄', '會同時刪除「融資結清」與「現股建倉」兩筆紀錄。確定？');
        if (!ok) return;
        acc.trades = acc.trades.filter(t => t._convertId !== cvid);
        save();
        renderAll();
        toast('已刪除資轉現紀錄', 'ok');
      } else if (act === 'edit-trade') {
        const idx = parseInt(btn.dataset.idx, 10);
        const t = acc.trades[idx];
        if (t) await editTradeDialog(t);
      } else if (act === 'del-trade') {
        const idx = parseInt(btn.dataset.idx, 10);
        const t = acc.trades[idx];
        if (t) await deleteTradeConfirm(t);
      }
    };
  });
}

// ============================================================
// 編輯 / 刪除投資明細
// ============================================================
async function editTradeDialog(t) {
  const result = await showModal({
    title: `編輯投資明細 - ${t.code} ${t.name||''}`,
    html: `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;min-width:560px">
        <label>成交日<input type="date" id="t-date" value="${(t.date||'').replace(/\//g,'-')}"></label>
        <label>交割日<input type="date" id="t-settle" value="${(t.settleDate||'').replace(/\//g,'-')}"></label>
        <label>代號<input type="text" id="t-code" value="${t.code||''}"></label>
        <label>名稱<input type="text" id="t-name" value="${(t.name||'').replace(/"/g,'&quot;')}"></label>
        <label>買賣
          <select id="t-action">
            <option value="買" ${t.action==='買'?'selected':''}>買</option>
            <option value="賣" ${t.action==='賣'?'selected':''}>賣</option>
          </select>
        </label>
        <label>類別
          <select id="t-cat">
            <option value="現股" ${t.category==='現股'?'selected':''}>現股</option>
            <option value="融資" ${t.category==='融資'?'selected':''}>融資</option>
            <option value="融券" ${t.category==='融券'?'selected':''}>融券</option>
            <option value="現股當沖" ${t.category==='現股當沖'?'selected':''}>現股當沖</option>
          </select>
        </label>
        <label>數量<input type="number" id="t-qty" value="${t.qty||0}"></label>
        <label>單價<input type="number" id="t-price" value="${t.price||0}" step="0.01"></label>
        <label>價金<input type="number" id="t-amount" value="${t.amount||0}"></label>
        <label>應收付款<input type="number" id="t-receivable" value="${t.receivable||0}"></label>
        <label>手續費<input type="number" id="t-fee" value="${t.fee||0}"></label>
        <label>交易稅<input type="number" id="t-tax" value="${t.tax||0}"></label>
        <label>融資金額/保證金<input type="number" id="t-margin" value="${t.marginAmount||0}"></label>
        <label>融資/券利息<input type="number" id="t-interest" value="${t.marginInterest||0}"></label>
        <label>融券手續費<input type="number" id="t-shortfee" value="${t.shortFee||0}"></label>
        <label>損益<input type="number" id="t-pl" value="${t.pl||0}"></label>
        <label style="grid-column:1/-1">備註<input type="text" id="t-note" value="${(t.note||'').replace(/"/g,'&quot;')}"></label>
      </div>
    `,
    confirmText: '儲存',
    onConfirm: (body) => ({
      date: body.querySelector('#t-date').value,
      settleDate: body.querySelector('#t-settle').value,
      code: body.querySelector('#t-code').value.trim(),
      name: body.querySelector('#t-name').value.trim(),
      action: body.querySelector('#t-action').value,
      category: body.querySelector('#t-cat').value,
      qty: parseFloat(body.querySelector('#t-qty').value) || 0,
      price: parseFloat(body.querySelector('#t-price').value) || 0,
      amount: parseFloat(body.querySelector('#t-amount').value) || 0,
      receivable: parseFloat(body.querySelector('#t-receivable').value) || 0,
      fee: parseFloat(body.querySelector('#t-fee').value) || 0,
      tax: parseFloat(body.querySelector('#t-tax').value) || 0,
      marginAmount: parseFloat(body.querySelector('#t-margin').value) || 0,
      marginInterest: parseFloat(body.querySelector('#t-interest').value) || 0,
      shortFee: parseFloat(body.querySelector('#t-shortfee').value) || 0,
      pl: parseFloat(body.querySelector('#t-pl').value) || 0,
      note: body.querySelector('#t-note').value.trim()
    })
  });
  if (!result) return;

  t.date = Parsers.formatDate(result.date);
  t.settleDate = result.settleDate ? Parsers.formatDate(result.settleDate) : '';
  t.code = result.code;
  t.name = result.name;
  t.action = result.action;
  t.category = result.category;
  t.qty = result.qty;
  t.price = result.price;
  t.amount = result.amount;
  t.receivable = result.receivable;
  t.fee = result.fee;
  t.tax = result.tax;
  t.marginAmount = result.marginAmount;
  t.marginInterest = result.marginInterest;
  t.shortFee = result.shortFee;
  t.pl = result.pl;
  t.note = result.note;
  t._userEdited = true;

  // 編輯後重跑已實現 ↔ 投資明細 比對
  const acc = getCurrentAccount();
  Parsers.enrichRealizedWithInterest(acc.realized, acc.trades);

  save();
  renderAll();
  toast('已更新', 'ok');
}

async function deleteTradeConfirm(t) {
  const acc = getCurrentAccount();
  if (!acc) return;
  const ok = await confirmDialog(
    '刪除投資明細',
    `確定刪除這筆？<br><br><b>${t.code} ${t.name||''}</b><br>${t.date} ${t.action} ${t.category} ${fmt(t.qty)}股 @${fmt(t.price,{decimals:2})}<br><br>注意：如果這筆有對應的「已實現損益」配對紀錄，會自動失效需要重新配對。`
  );
  if (!ok) return;

  const idx = acc.trades.indexOf(t);
  if (idx < 0) return;
  acc.trades.splice(idx, 1);

  // 重跑已實現比對
  Parsers.enrichRealizedWithInterest(acc.realized, acc.trades);

  save();
  renderAll();
  toast('已刪除', 'ok');
}

// ============================================================
// 事件綁定
// ============================================================

// 安全綁定：抓不到 element 也不會炸
function bindClick(id, handler) {
  const el = document.getElementById(id);
  if (el) el.onclick = handler;
  else console.warn(`bindClick: element "${id}" not found`);
}
function bindChange(id, handler) {
  const el = document.getElementById(id);
  if (el) el.onchange = handler;
  else console.warn(`bindChange: element "${id}" not found`);
}
function bindInput(id, handler) {
  const el = document.getElementById(id);
  if (el) el.oninput = handler;
  else console.warn(`bindInput: element "${id}" not found`);
}

function bindEvents() {
  // Tab 切換
  document.querySelectorAll('.tab').forEach(btn => {
    btn.onclick = () => {
      document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
      // 切換頁籤時重新渲染（避免某些子表沒資料的情況）
      renderAll();
      // 切到未實現損益且本次還沒拓過 → 自動拓一次
      if (btn.dataset.tab === 'unrealized' && !_priceAutoFetched) {
        _priceAutoFetched = true;
        refreshPrices(true).catch(() => {});
      }
    };
  });

  // 帳戶選擇
  bindChange('accountSelect', (e) => {
    State.currentAccountId = e.target.value;
    State.data.currentAccountId = e.target.value;
    _priceAutoFetched = false; // 換帳戶 → 重新自動拓一次
    save();
    refreshAccountSelector();
    renderAll();
  });
  bindClick('btnNewAccount', newAccount);
  bindClick('btnRenameAccount', renameAccount);
  bindClick('btnReorderAccount', reorderAccountsDialog);
  bindClick('btnDeleteAccount', deleteAccount);

  // 備份
  bindClick('btnBackup', exportBackup);
  bindClick('btnRestore', () => document.getElementById('restoreFile').click());
  bindChange('restoreFile', (e) => {
    const f = e.target.files[0]; if (f) restoreBackup(f);
    e.target.value = '';
  });

  // 上傳
  bindUpload('upUnrealized', importUnrealized);
  bindUpload('upTrades', importTrades);
  bindUpload('upRealized', importRealized);

  // 快照日期預設今天
  const sd = document.getElementById('snapshotDate');
  if (sd) sd.valueAsDate = new Date();

  // 搜尋
  bindInput('rzSearch', renderRealized);
  bindInput('trSearch', renderTrades);
  bindChange('trFilterType', renderTrades);
  bindChange('trFilterAction', renderTrades);

  // 匯出已實現
  bindClick('btnExportRealized', exportRealizedExcel);

  // 每月損益
  bindChange('monthlyYear', renderMonthly);
  bindClick('btnExportMonthly', exportMonthlyExcel);

  // 月份比較分析
  bindChange('cmpMode', renderMonthCompare);
  bindChange('cmpMetric', renderMonthCompare);
  bindChange('cmpMonthA', renderMonthCompare);
  bindChange('cmpMonthB', renderMonthCompare);

  // 個股損益分析
  bindInput('stockSearch', renderStockAnalysis);
  bindChange('stockFilter', renderStockAnalysis);
  bindClick('btnExportStocks', exportStocksExcel);

  // 每日當沖
  bindChange('dayTradeYear', renderDayTrades);
  bindClick('btnExportDayTrades', exportDayTradesExcel);

  // 資轉現
  bindClick('btnConvertToCash', showConvertToCashDialog);

  // 借款
  bindClick('btnAddLoan', addLoan);

  // 融資回補
  bindClick('btnAddMarginCall', addMarginCall);

  // 即時股價
  bindClick('btnRefreshPrice', () => refreshPrices(false));

  // 股利紀錄
  bindClick('btnAddDividend', () => addDividendEntry());
  bindClick('btnExportDividends', exportDividendsExcel);
  bindChange('divViewMode', renderDividends);
  const upDiv = document.getElementById('upDividends');
  if (upDiv) {
    upDiv.onchange = (e) => {
      const f = e.target.files[0];
      if (f) importDividendFile(f);
      e.target.value = '';
    };
  }
  // 股利匯入：放在「⬆️ 匯入資料」頁的入口
  const upDivImp = document.getElementById('upDividendsImport');
  if (upDivImp) {
    upDivImp.onchange = (e) => {
      const f = e.target.files[0];
      if (f) importDividendFile(f);
      e.target.value = '';
    };
  }

  // 清空帳戶
  bindClick('btnClearAccount', async () => {
    const acc = getCurrentAccount();
    if (!acc) return;
    const ok = await confirmDialog('清空帳戶資料', `確定清空「${acc.name}」的所有未實現、投資明細、已實現資料？（帳戶與調整紀錄保留）`);
    if (!ok) return;
    acc.unrealized = []; acc.trades = []; acc.realized = []; acc.snapshots = [];
    save(); renderAll(); toast('已清空', 'ok');
  });
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
