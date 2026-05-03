// ============================================================
// 券商檔案解析器
// 支援：元大（Ｄ系列）、國泰、新光（M系列）
// 所有解析器都自動偵測格式，不需手動選擇
// ============================================================

const Parsers = (() => {

  // ---------- 工具函式 ----------
  function toNumber(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    const raw = String(v).trim();
    if (raw === '--' || raw === '-' || raw === 'N/A') return 0;
    const s = raw.replace(/,/g, '').replace(/\s/g, '').replace(/^\+/, '');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  function toString(v) {
    if (v == null) return '';
    if (v instanceof Date) return formatDate(v);
    return String(v).trim();
  }

  // 國泰格式："鴻海(2317)" 或 "00940元大台灣價值高息(00940)" 或 "緯穎-KY(6669)" → {code, name}
  function parseStockName(s) {
    if (s == null) return { code: '', name: '' };
    const str = String(s).trim();
    const m = str.match(/^(.+?)\(([A-Z0-9]+)\)\s*$/);
    if (m) return { code: m[2], name: m[1].trim() };
    if (/^\d+$/.test(str)) return { code: str, name: '' };
    return { code: '', name: str };
  }

  function formatDate(d) {
    if (!d) return '';
    if (typeof d === 'string') {
      // 民國年「114/02/19」→ 西元 2025/02/19
      const rocMatch = d.match(/^(\d{2,3})[/-](\d{1,2})[/-](\d{1,2})/);
      if (rocMatch) {
        const y = parseInt(rocMatch[1], 10);
        const m = rocMatch[2].padStart(2, '0');
        const day = rocMatch[3].padStart(2, '0');
        if (y < 200) return `${y + 1911}/${m}/${day}`;
        return `${y}/${m}/${day}`;
      }
      // 已是字串 → 統一成 yyyy/mm/dd
      const parts = d.replace(/-/g, '/').split('/').map(s => s.trim());
      if (parts.length === 3) {
        const yy = parts[0].length === 4 ? parts[0] : `20${parts[0]}`;
        return `${yy}/${parts[1].padStart(2,'0')}/${parts[2].padStart(2,'0')}`;
      }
      return d;
    }
    if (d instanceof Date) {
      const y = d.getFullYear();
      const m = String(d.getMonth()+1).padStart(2,'0');
      const day = String(d.getDate()).padStart(2,'0');
      return `${y}/${m}/${day}`;
    }
    return String(d);
  }

  // 把 SheetJS 讀進來的 worksheet 轉成二維陣列
  function sheetToRows(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  // ---------- 自動偵測格式 ----------
  function detectFormat(rows, type) {
    // 把前 5 列攤平
    const headerText = rows.slice(0, 5).flat().map(x => x == null ? '' : String(x)).join('|');

    // 新光特徵：第 1-2 列是「帳號」「8560-」開頭，欄位有「資息」「券息」「沖銷均價」「現沖借券費」「借貸還款」
    if (/(資息.*券息|現沖借券費|借貸還款|沖銷均價)/.test(headerText)) return 'sinopac';
    // 未實現的新光特徵：「商品.*股號」「股票總市值」「預估損益」「預估淨收付」
    if (type === 'unrealized' && /(預估淨收付|股票總市值|預估損益|可用股數)/.test(headerText)) return 'sinopac';

    // 國泰特徵：欄位明確、有「自備款/保證金」(中間斜線)、有「淨收額」「淨付額」、未實現有「股票現值」
    if (type === 'unrealized' && /股票現值|未實現損益率/.test(headerText)) return 'cathay';
    if (type === 'trades' && /(自備款\/保證金|融資金\/擔保金|淨收額|淨付額)/.test(headerText)) return 'cathay';
    if (type === 'realized' && /買進日期.*賣出日期.*股票名稱.*交易類別.*成交股數/.test(headerText)) return 'cathay';

    // 元大特徵：合計行有「TWD 新台幣」、有「沖銷原始成本」「融資金額/融券保證金」
    if (/(沖銷原始成本|融資金額.*融券保證金|二代健保|TWD 新台幣)/.test(headerText)) return 'yuanta';

    // 已實現的元大特徵：第 1 列就有 "代號" 在 A1
    if (type === 'realized') {
      const a1 = String(rows[0]?.[0] || '').trim();
      if (a1 === '代號') return 'yuanta';
    }
    if (type === 'unrealized') {
      const a1 = String(rows[0]?.[0] || '').trim();
      if (a1 === '明細') return 'yuanta';
    }
    if (type === 'trades') {
      if ((rows[0] || []).length >= 20) return 'yuanta';
    }

    return 'yuanta';
  }

  // ============================================================
  // 元大：未實現損益
  // 表頭兩列合併：股票代號/名稱、交易類別、庫存數量、現價、市值、成本金額、平均單價、手續費、交易稅、利息費用、損益、報酬率、幣別
  // 資料從第 3 列起，最後一列為合計（A 欄含「總市值」字樣）
  // ============================================================
  function parseUnrealizedYuanta(rows) {
    const items = [];
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      // 跳過合計列
      const firstCell = toString(r[0]);
      if (firstCell.includes('總市值') || firstCell.includes('總成本')) continue;
      const code = toString(r[1]);
      if (!code || !/^\d+$/.test(code)) continue;

      items.push({
        code,
        name: toString(r[2]),
        category: toString(r[3]),
        qty: toNumber(r[4]),
        price: toNumber(r[5]),
        marketValue: toNumber(r[6]),
        cost: toNumber(r[7]),
        avgCost: toNumber(r[8]),
        fee: toNumber(r[9]),
        tax: toNumber(r[10]),
        interest: toNumber(r[11]),
        pl: toNumber(r[12]),
        rate: toString(r[13]),
        currency: toString(r[14])
      });
    }
    return items;
  }

  // ============================================================
  // 元大：投資明細
  // 表頭兩列合併：成交日期、股票代號/名稱、交易種類、買賣、交易類別、成交數量/單價、價金、手續費、交易稅、應收付款、
  //              融資金額/融券保證金、自備款/擔保品、融資券利息、融券手續費、標借費、利息代扣稅款、二代健保補充費、損益、交割日、幣別
  // ============================================================
  function parseTradesYuanta(rows) {
    const items = [];
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const date = toString(r[0]);
      const code = toString(r[1]);
      if (!date || !code) continue;
      if (!/^\d+$/.test(code)) continue;

      items.push({
        date: formatDate(r[0]),
        code,
        name: toString(r[2]),
        kind: toString(r[3]),       // 整股
        action: toString(r[4]),     // 買 / 賣
        category: toString(r[5]),   // 現股 / 融資 / 融券
        qty: toNumber(r[6]),
        price: toNumber(r[7]),
        amount: toNumber(r[8]),
        fee: toNumber(r[9]),
        tax: toNumber(r[10]),
        receivable: toNumber(r[11]),
        marginAmount: toNumber(r[12]),     // 融資金額 / 融券保證金
        ownFund: toNumber(r[13]),          // 自備款 / 擔保品
        marginInterest: toNumber(r[14]),   // 融資券利息
        shortFee: toNumber(r[15]),         // 融券手續費
        borrowFee: toNumber(r[16]),        // 標借費
        interestTax: toNumber(r[17]),      // 利息代扣稅款
        nhi: toNumber(r[18]),              // 二代健保補充費
        pl: toNumber(r[19]),
        settleDate: formatDate(r[20]),
        currency: toString(r[21])
      });
    }
    return items;
  }

  // ============================================================
  // 元大：已實現損益
  // 單層表頭：代號、名稱、賣出日期、賣出委託書號、賣出交易類別、賣價、買進日期、買進委託書號、買進交易類別、買價、
  //          數量、沖銷原始成本、手續費、交易稅、盈虧
  // ============================================================
  function parseRealizedYuanta(rows) {
    const items = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const code = toString(r[0]);
      if (!code) continue;
      if (!/^\d+$/.test(code)) continue;

      items.push({
        code,
        name: toString(r[1]),
        sellDate: formatDate(r[2]),
        sellOrderNo: toString(r[3]),
        sellCategory: toString(r[4]),
        sellPrice: toNumber(r[5]),
        buyDate: formatDate(r[6]),
        buyOrderNo: toString(r[7]),
        buyCategory: toString(r[8]),
        buyPrice: toNumber(r[9]),
        qty: toNumber(r[10]),
        cost: toNumber(r[11]),
        fee: toNumber(r[12]),
        tax: toNumber(r[13]),
        pl: toNumber(r[14])
        // 後續會補上：interest, shortFee, adjust, note, actualPL
      });
    }
    return items;
  }

  // ============================================================
  // 國泰：未實現損益
  // 第 1 列：標題（現股損益）
  // 第 2 列：表頭（股票名稱、幣別、類別、庫存股數、持有成本、成交均價、現價、股票現值、未實現損益、未實現損益率%）
  // 資料從第 3 列起
  // ============================================================
  function parseUnrealizedCathay(rows) {
    const items = [];
    for (let i = 2; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const stockField = r[0];
      if (!stockField) continue;
      const { code, name } = parseStockName(stockField);
      if (!code) continue;

      const qty = toNumber(r[3]);
      const cost = toNumber(r[4]);
      const avgCost = toNumber(r[5]);
      const price = toNumber(r[6]);
      const marketValue = toNumber(r[7]);
      const pl = toNumber(r[8]);
      const rateRaw = toNumber(r[9]);
      // 國泰的「未實現損益率%」是小數（0.0373），轉成百分比字串
      const rateStr = rateRaw === 0 ? '—' : ((rateRaw > 0 ? '+' : '') + (rateRaw * 100).toFixed(2) + '%');

      items.push({
        code, name,
        category: toString(r[2]) || '現股',
        qty, price, marketValue, cost, avgCost,
        fee: 0, tax: 0, interest: 0, // 國泰未實現未提供
        pl,
        rate: rateStr,
        currency: toString(r[1])
      });
    }
    return items;
  }

  // ============================================================
  // 國泰：投資明細
  // 表頭一列：成交日期、交易類別、股票名稱、幣別、成交股數、成交單價、成交價金、手續費、交易稅、
  //          自備款/保證金、融資金/擔保金、融券費、標借費、利息、稅款、淨收額、淨付額
  //
  // 特殊處理：每筆「融資賣出/融券買進」後面跟一筆「原資買/原券賣」輔助列
  //          → A 欄是文字「原資買」「原券賣」，B 欄是股票，第 14 欄(利息) 是該筆的已付利息
  // ============================================================
  function parseTradesCathay(rows) {
    const items = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const a = r[0];
      // 跳過合計列
      if (a === '成交價金' || (typeof a === 'string' && /^成交價金$/.test(a))) continue;
      // 跳過輔助列「原資買」「原券賣」（已透過 lookahead 處理）
      if (typeof a === 'string' && /^(原資買|原券賣)$/.test(a.trim())) continue;
      // 主資料列：A 欄必須是日期
      if (!(a instanceof Date) && !/^\d{4}[/-]\d{1,2}[/-]\d{1,2}/.test(String(a || ''))) continue;

      const cat = toString(r[1]);
      const { code, name } = parseStockName(r[2]);
      if (!code) continue;

      const qty = toNumber(r[4]);
      const price = toNumber(r[5]);
      const amount = toNumber(r[6]);
      const fee = toNumber(r[7]);
      const tax = toNumber(r[8]);
      const ownFund = toNumber(r[9]);       // 自備款/保證金
      const marginAmount = toNumber(r[10]); // 融資金/擔保金
      let shortFee = toNumber(r[11]);       // 融券費
      const borrowFee = toNumber(r[12]);
      let interest = toNumber(r[13]);       // 利息
      const interestTax = toNumber(r[14]);
      const netReceive = toNumber(r[15]);
      const netPay = toNumber(r[16]);

      // 看下一列是否為輔助「原資買/原券賣」→ 有的話用它的利息
      const next = rows[i + 1];
      if (next && typeof next[0] === 'string' && /^(原資買|原券賣)$/.test(next[0].trim())) {
        const auxInterest = toNumber(next[13]);
        if (auxInterest && !interest) interest = auxInterest;
      }

      // 動作 + 類別
      let action, category;
      if (cat === '現股買進') { action = '買'; category = '現股'; }
      else if (cat === '現股賣出') { action = '賣'; category = '現股'; }
      else if (cat === '融資買進') { action = '買'; category = '融資'; }
      else if (cat === '融資賣出') { action = '賣'; category = '融資'; }
      else if (cat === '融券賣出') { action = '賣'; category = '融券'; }
      else if (cat === '融券買進') { action = '買'; category = '融券'; }
      else if (cat === '現沖買進') { action = '買'; category = '現股當沖'; }
      else if (cat === '現賣沖') { action = '賣'; category = '現股當沖'; }
      else if (cat === '現股當沖') { action = '賣'; category = '現股當沖'; }
      else { action = ''; category = cat; }

      items.push({
        date: formatDate(a),
        code, name,
        kind: '整股',
        action,
        category,
        rawCategory: cat,
        qty, price, amount, fee, tax,
        receivable: action === '賣' ? netReceive : -netPay,
        marginAmount,
        ownFund,
        marginInterest: interest,
        shortFee,
        borrowFee,
        interestTax,
        nhi: 0,
        pl: 0,
        settleDate: '',
        currency: toString(r[3])
      });
    }
    return items;
  }

  // ============================================================
  // 新光（SinoPac/M）：未實現損益
  // 第 1 列：查詢日期
  // 第 2-3 列：帳號、委託狀態
  // 第 4 列：合計
  // 第 5 列：表頭（商品、股號、類別、股數、可用股數、市價、市值、成本價、成本、預估損益、
  //              成交均價、報酬率(%)、預估淨收付、融資金、融券金、保證金、擔保品、利息、手續費、幣別）
  // 資料從第 6 列起
  // ============================================================
  function parseUnrealizedSinopac(rows) {
    const items = [];
    // 找表頭列
    let headerRow = -1;
    for (let i = 0; i < Math.min(8, rows.length); i++) {
      const r = rows[i] || [];
      const flat = r.map(x => String(x || '').trim()).join('|');
      if (/商品.*股號.*類別.*股數/.test(flat)) {
        headerRow = i; break;
      }
    }
    if (headerRow < 0) headerRow = 4;

    for (let i = headerRow + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const name = toString(r[0]);
      const code = toString(r[1]);
      if (!name && !code) continue;
      if (!code || (typeof r[1] !== 'number' && !/^[A-Z0-9]+$/.test(code))) continue;

      const rateRaw = toNumber(r[11]); // 已是百分比數字（如 -9.46）
      const rateStr = rateRaw === 0 ? '—' : ((rateRaw > 0 ? '+' : '') + rateRaw.toFixed(2) + '%');

      items.push({
        code, name,
        category: toString(r[2]) || '現股',  // 普通 → 改成「現股」
        qty: toNumber(r[3]),
        price: toNumber(r[5]),
        marketValue: toNumber(r[6]),
        cost: toNumber(r[8]),
        avgCost: toNumber(r[10]) || toNumber(r[7]),
        fee: toNumber(r[18]),
        tax: 0,
        interest: toNumber(r[17]),
        pl: toNumber(r[9]),
        rate: rateStr,
        currency: toString(r[19])
      });
    }
    return items;
  }

  // ============================================================
  // 新光：交易明細
  // 第 1-3 列：帳號、合計、總計
  // 第 4 列：表頭（委託日期、股票、代號、成交股數、單價、類別、價金、手續費、交易稅、
  //              融券手續費、預繳金、淨收付、資息、券息、損益、標借費、現沖借券費、沖銷均價、
  //              報酬率(%)、當沖、借貸還款、委託書號、幣別）
  // 資料從第 5 列起
  // ============================================================
  function parseTradesSinopac(rows) {
    const items = [];
    // 找表頭列
    let headerRow = -1;
    for (let i = 0; i < Math.min(8, rows.length); i++) {
      const r = rows[i] || [];
      const flat = r.map(x => String(x || '').trim()).join('|');
      if (/委託日期.*股票.*代號.*類別/.test(flat)) {
        headerRow = i; break;
      }
    }
    if (headerRow < 0) headerRow = 3;

    for (let i = headerRow + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || !r[0]) continue;
      const dateRaw = r[0];
      // 跳過合計列（A 欄是「總損益」「總報酬」等）
      if (typeof dateRaw === 'string' && !/^\d{2,4}[/-]\d{1,2}[/-]\d{1,2}/.test(dateRaw.trim())) continue;
      if (!(dateRaw instanceof Date) && typeof dateRaw !== 'string') continue;

      const code = toString(r[2]);
      if (!code) continue;
      const cat = toString(r[5]); // '現買' '現賣' '資買' '資賣' '券賣' '券買'

      let action, category;
      if (cat === '現買') { action = '買'; category = '現股'; }
      else if (cat === '現賣') { action = '賣'; category = '現股'; }
      else if (cat === '資買') { action = '買'; category = '融資'; }
      else if (cat === '資賣') { action = '賣'; category = '融資'; }
      else if (cat === '券賣') { action = '賣'; category = '融券'; }
      else if (cat === '券買') { action = '買'; category = '融券'; }
      else if (cat.includes('沖')) {
        category = '現股當沖';
        action = cat.includes('賣') ? '賣' : '買';
      }
      else { action = ''; category = cat; }

      // 資息(12) + 券息(13) 都是利息類，加總
      const interest = toNumber(r[12]) + toNumber(r[13]);
      const shortFee = toNumber(r[9]); // 融券手續費

      items.push({
        date: formatDate(dateRaw),
        code,
        name: toString(r[1]),
        kind: '整股',
        action, category,
        rawCategory: cat,
        qty: toNumber(r[3]),
        price: toNumber(r[4]),
        amount: toNumber(r[6]),
        fee: toNumber(r[7]),
        tax: toNumber(r[8]),
        receivable: toNumber(r[11]),
        marginAmount: 0,
        ownFund: 0,
        marginInterest: interest,
        shortFee,
        borrowFee: toNumber(r[15]),
        interestTax: 0,
        nhi: 0,
        pl: toNumber(r[14]),
        settleDate: '',
        currency: toString(r[22]),
        orderNo: toString(r[21])
      });
    }
    return items;
  }

  // ============================================================
  // 新光：已實現損益
  // 結構與「交易明細」幾乎一樣，差在沒有「委託日期」改叫「成交日期」、沒有「類別」這欄是直接用同一欄
  // 第 4 列：表頭（成交日期、股票、代號、類別、成交股數、單價、價金、手續費、交易稅、
  //              融券手續費、預繳金、淨收付、資息、券息、損益、...、委託書號、幣別）
  // 資料從第 5 列起
  // 注意：末尾可能有「已實現損益」「日期區間」之類的合計列要跳過
  // ============================================================
  function parseRealizedSinopac(rows) {
    const items = [];
    let headerRow = -1;
    for (let i = 0; i < Math.min(8, rows.length); i++) {
      const r = rows[i] || [];
      const flat = r.map(x => String(x || '').trim()).join('|');
      if (/成交日期.*股票.*代號.*類別.*成交股數/.test(flat)) {
        headerRow = i; break;
      }
    }
    if (headerRow < 0) headerRow = 3;

    for (let i = headerRow + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || !r[0]) continue;
      const dateRaw = r[0];
      // 跳過合計/註解列
      if (typeof dateRaw === 'string' && !/^\d{2,4}[/-]\d{1,2}[/-]\d{1,2}/.test(dateRaw.trim())) continue;
      if (!(dateRaw instanceof Date) && typeof dateRaw !== 'string') continue;

      const code = toString(r[2]);
      if (!code) continue;
      const cat = toString(r[3]);

      let sellCategory;
      if (cat === '現賣' || cat === '現買') sellCategory = '現股';
      else if (cat === '資賣' || cat === '資買') sellCategory = '融資';
      else if (cat === '券賣' || cat === '券買') sellCategory = '融券';
      else if (cat.includes('沖')) sellCategory = '現股當沖';
      else sellCategory = cat;

      const interest = toNumber(r[12]) + toNumber(r[13]); // 資息+券息
      const shortFee = toNumber(r[9]);
      const rateRaw = toNumber(r[18]);
      const rateStr = rateRaw === 0 ? '—' : ((rateRaw > 0 ? '+' : '') + rateRaw.toFixed(2) + '%');

      items.push({
        code,
        name: toString(r[1]),
        sellDate: formatDate(dateRaw),
        sellOrderNo: toString(r[21]),
        sellCategory,
        rawCategory: cat,
        sellPrice: toNumber(r[5]),
        // 新光的「沖銷均價」(17) 就是買進均價
        buyDate: '',
        buyOrderNo: '',
        buyCategory: sellCategory,
        buyPrice: toNumber(r[17]),
        qty: toNumber(r[4]),
        cost: toNumber(r[17]) * toNumber(r[4]), // 沖銷均價 × 數量
        fee: toNumber(r[7]),
        tax: toNumber(r[8]),
        pl: toNumber(r[14]),
        // 新光直接從這張表上就能拿到利息和融券手續費
        _preInterest: interest,
        _preShortFee: shortFee,
        rate: rateStr
      });
    }
    return items;
  }
  function parseRealizedCathay(rows) {
    const items = [];
    // 找表頭列
    let headerRow = -1;
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const r = rows[i] || [];
      const flat = r.map(x => String(x || '').trim()).join('|');
      if (/買進日期.*賣出日期.*股票名稱/.test(flat)) {
        headerRow = i; break;
      }
    }
    if (headerRow < 0) headerRow = 1; // 預設

    for (let i = headerRow + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || !r[0]) continue;
      const { code, name } = parseStockName(r[2]);
      if (!code) continue;

      const cat = toString(r[4]); // '現股賣出' / '融資賣出' / '融券買進' / '現沖賣出' 等
      // 對應到統一格式：sellCategory 用「融資/融券/現股/現股當沖」
      let sellCategory;
      if (cat.includes('融資')) sellCategory = '融資';
      else if (cat.includes('融券')) sellCategory = '融券';
      else if (cat.includes('沖')) sellCategory = '現股當沖';
      else sellCategory = '現股';

      const rateRaw = toNumber(r[9]);
      const rateStr = rateRaw === 0 ? '—' : ((rateRaw > 0 ? '+' : '') + (rateRaw * 100).toFixed(2) + '%');

      items.push({
        code, name,
        sellDate: formatDate(r[1]),
        sellOrderNo: '',
        sellCategory,
        rawCategory: cat,
        sellPrice: toNumber(r[7]),
        buyDate: formatDate(r[0]),
        buyOrderNo: '',
        buyCategory: sellCategory, // 國泰沒給單獨買進類別，假設一致
        buyPrice: toNumber(r[6]),
        qty: toNumber(r[5]),
        cost: 0, // 國泰沒給沖銷原始成本
        fee: 0,
        tax: 0,
        pl: toNumber(r[8]),
        rate: rateStr
      });
    }
    return items;
  }

  // ---------- 對外入口 ----------
  function parseUnrealized(workbook) {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheetToRows(ws);
    const fmt = detectFormat(rows, 'unrealized');
    if (fmt === 'sinopac') return { broker: 'sinopac', items: parseUnrealizedSinopac(rows) };
    if (fmt === 'cathay') return { broker: 'cathay', items: parseUnrealizedCathay(rows) };
    return { broker: 'yuanta', items: parseUnrealizedYuanta(rows) };
  }

  function parseTrades(workbook) {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheetToRows(ws);
    const fmt = detectFormat(rows, 'trades');
    if (fmt === 'sinopac') return { broker: 'sinopac', items: parseTradesSinopac(rows) };
    if (fmt === 'cathay') return { broker: 'cathay', items: parseTradesCathay(rows) };
    return { broker: 'yuanta', items: parseTradesYuanta(rows) };
  }

  function parseRealized(workbook) {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheetToRows(ws);
    const fmt = detectFormat(rows, 'realized');
    if (fmt === 'sinopac') return { broker: 'sinopac', items: parseRealizedSinopac(rows) };
    if (fmt === 'cathay') return { broker: 'cathay', items: parseRealizedCathay(rows) };
    return { broker: 'yuanta', items: parseRealizedYuanta(rows) };
  }

  // ============================================================
  // 已實現 + 投資明細 比對 → 補上利息與融券手續費
  //
  // 重點：同一張賣單（融資/融券）的「投資明細」上是一筆（如 3000 股）；
  //      但「已實現損益」會按買進批次拆成多筆（如 3 筆 1000 股）。
  //
  // 演算法：
  // 1. 把已實現損益按「代號 + 賣出日 + 賣出類別」分組
  // 2. 組總數量比對投資明細上同條件的賣出筆，數量相等者視為一張委託
  // 3. 投資明細的利息/融券手續費按「沖銷成本佔組總成本比例」拆分到組內每筆
  // 4. 同日同股同類別有多張委託時，依數量逐一消耗
  //
  // 特例：新光的已實現本身就含「資息」「券息」「融券手續費」欄位，
  //       會在 _preInterest / _preShortFee 上預先記錄，直接採用不需比對。
  // ============================================================
  function enrichRealizedWithInterest(realizedItems, tradeItems) {
    // 1. 先把所有列重設為 0；新光（_preInterest 存在）直接用預設值
    for (const r of realizedItems) {
      if (r._preInterest != null || r._preShortFee != null) {
        // 新光：直接用已實現表上的數字
        r.interest = r._preInterest || 0;
        r.shortFee = r._preShortFee || 0;
        r._unmatched = false;
        continue;
      }
      if (r.sellCategory !== '融資' && r.sellCategory !== '融券') {
        r.interest = 0; r.shortFee = 0; r._unmatched = false;
      } else {
        r.interest = 0; r.shortFee = 0; r._unmatched = true;
      }
    }

    // 2. 取出賣出融資/融券交易，按 (code, date, category) 分組（同條件可能有多張委託）
    const sellTrades = tradeItems.filter(t =>
      t.action === '賣' && (t.category === '融資' || t.category === '融券')
    );
    const tradeBuckets = new Map(); // key -> array of trade
    for (const t of sellTrades) {
      const k = `${t.code}|${t.date}|${t.category}`;
      if (!tradeBuckets.has(k)) tradeBuckets.set(k, []);
      tradeBuckets.get(k).push(t);
    }

    // 3. 把已實現按 (code, sellDate, sellCategory) 分組（保留原順序）
    //    跳過新光（已預先填好）
    const realizedBuckets = new Map();
    for (const r of realizedItems) {
      if (r._preInterest != null || r._preShortFee != null) continue;
      if (r.sellCategory !== '融資' && r.sellCategory !== '融券') continue;
      const k = `${r.code}|${r.sellDate}|${r.sellCategory}`;
      if (!realizedBuckets.has(k)) realizedBuckets.set(k, []);
      realizedBuckets.get(k).push(r);
    }

    let matched = 0;
    let total = 0;

    // 計入新光那些
    for (const r of realizedItems) {
      if (r._preInterest != null || r._preShortFee != null) {
        if (r.sellCategory === '融資' || r.sellCategory === '融券') {
          total++;
          matched++;
        }
      }
    }

    // 4. 組對組消耗
    for (const [key, rzGroup] of realizedBuckets) {
      total += rzGroup.length;
      const trades = tradeBuckets.get(key) || [];
      if (!trades.length) continue;

      const queue = rzGroup.slice();
      for (const trade of trades) {
        let remaining = trade.qty;
        const consumed = [];
        let costSum = 0;
        while (remaining > 0 && queue.length > 0) {
          const r = queue[0];
          if (r.qty <= remaining) {
            consumed.push(r);
            costSum += (r.cost || 0);
            remaining -= r.qty;
            queue.shift();
          } else {
            consumed.push(r);
            costSum += (r.cost || 0);
            remaining = 0;
            queue.shift();
          }
        }
        if (consumed.length === 0) continue;

        const interest = trade.marginInterest || 0;
        const sFee = trade.shortFee || 0;
        if (consumed.length === 1) {
          consumed[0].interest = interest;
          consumed[0].shortFee = sFee;
          consumed[0]._unmatched = false;
          matched++;
        } else {
          let allocatedI = 0, allocatedS = 0;
          for (let i = 0; i < consumed.length; i++) {
            const r = consumed[i];
            if (i === consumed.length - 1) {
              r.interest = interest - allocatedI;
              r.shortFee = sFee - allocatedS;
            } else {
              const ratio = costSum > 0 ? (r.cost || 0) / costSum : (1 / consumed.length);
              r.interest = Math.round(interest * ratio);
              r.shortFee = Math.round(sFee * ratio);
              allocatedI += r.interest;
              allocatedS += r.shortFee;
            }
            r._unmatched = false;
            matched++;
          }
        }
      }
    }

    return { matched, total };
  }

  // ============================================================
  // 股利檔案解析（券商提供的「除權息總覽」HTML 偽裝 .xls）
  // 欄位：代號 / 名稱 / 現金股利 / 股票股利
  // ============================================================
  function parseDividendsFromText(text) {
    // 用 DOMParser 把 HTML 字串解析成表格
    const doc = new DOMParser().parseFromString(text, 'text/html');
    const tables = doc.querySelectorAll('table');
    const items = [];
    for (const tbl of tables) {
      const rows = tbl.querySelectorAll('tr');
      let headerSeen = false;
      for (const tr of rows) {
        const cells = [...tr.querySelectorAll('td, th')].map(c => (c.textContent || '').trim());
        if (cells.length < 3) continue;
        // 偵測表頭
        if (!headerSeen && /代號/.test(cells[0]) && /名稱/.test(cells[1])) {
          headerSeen = true;
          continue;
        }
        // 資料列：代號必須是英數字組合
        const code = cells[0];
        if (!code || !/^[A-Z0-9]+$/i.test(code)) continue;
        const name = cells[1] || '';
        const cash = toNumber(cells[2]);
        const stock = toNumber(cells[3]);
        if (cash === 0 && stock === 0) continue; // 跳過全空列
        items.push({
          code,
          name,
          cashTotal: cash,
          stockTotal: stock,
        });
      }
      if (items.length > 0) break; // 已找到資料就停
    }
    return items;
  }

  // 直接讀檔案物件（File / Blob），自動處理 HTML 或真正 xls/xlsx
  async function parseDividendsFile(file) {
    // 偵測是 HTML 還是 binary
    const headBytes = new Uint8Array(await file.slice(0, 16).arrayBuffer());
    const headStr = new TextDecoder('utf-8', { fatal: false }).decode(headBytes);
    const isHtml = /<\s*(html|table|meta|\?xml)/i.test(headStr);

    if (isHtml) {
      // 用 text 讀，然後 DOMParser
      let text = await file.text();
      // 有些券商會用 BIG5 編碼，瀏覽器預設用 UTF-8 讀會亂碼
      // 偵測：如果 text 含 BIG5 亂碼模式（？？或大量 \uFFFD），改用 BIG5
      if (text.includes('\uFFFD') || /[\xC0-\xFF]{3,}/.test(text)) {
        try {
          const buf = await file.arrayBuffer();
          text = new TextDecoder('big5').decode(buf);
        } catch (e) { /* 沒 big5 支援就吞 */ }
      }
      return { broker: 'html', items: parseDividendsFromText(text) };
    }

    // 真正的 xlsx/xls：用 SheetJS
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = sheetToRows(ws);
    const items = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      const code = toString(r[0]);
      const name = toString(r[1]);
      // 跳過表頭
      if (!code || code === '代號' || !/^[A-Z0-9]+$/i.test(code)) continue;
      const cash = toNumber(r[2]);
      const stock = toNumber(r[3]);
      if (cash === 0 && stock === 0) continue;
      items.push({ code, name, cashTotal: cash, stockTotal: stock });
    }
    return { broker: 'xlsx', items };
  }

  return {
    parseUnrealized,
    parseTrades,
    parseRealized,
    parseDividendsFile,
    parseDividendsFromText,
    enrichRealizedWithInterest,
    detectFormat,
    toNumber,
    toString,
    formatDate
  };
})();
