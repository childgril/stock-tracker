// ============================================================
// 券商檔案解析器
// 目前支援：元大（Ｄ系列）
// 之後可加：國泰、新光
// ============================================================

const Parsers = (() => {

  // ---------- 工具函式 ----------
  function toNumber(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/,/g, '').replace(/\s/g, '').replace(/^\+/, '');
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
      // 已是字串 → 統一成 yyyy/mm/dd
      const parts = d.replace(/-/g, '/').split('/').map(s => s.trim());
      if (parts.length === 3) {
        const y = parts[0].length === 4 ? parts[0] : `20${parts[0]}`;
        return `${y}/${parts[1].padStart(2,'0')}/${parts[2].padStart(2,'0')}`;
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
    // 把前 3 列攤平
    const headerText = rows.slice(0, 3).flat().map(x => x == null ? '' : String(x)).join('|');

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
      // 元大未實現第一列有合併儲存格「明細」「股票」
      const a1 = String(rows[0]?.[0] || '').trim();
      if (a1 === '明細') return 'yuanta';
    }
    if (type === 'trades') {
      // 元大投資明細第一列 A1 = "成交日期"，但欄位有 22 欄
      if ((rows[0] || []).length >= 20) return 'yuanta';
    }

    return 'yuanta'; // 不確定就先試元大
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
  // 國泰：已實現損益
  // 第 1 列：空或標題；第 2 列：表頭（買進日期、賣出日期、股票名稱、幣別、交易類別、成交股數、買進單價、賣出單價、損益、報酬率）
  // 資料從第 3 列起
  // ============================================================
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
    if (fmt === 'cathay') return { broker: 'cathay', items: parseUnrealizedCathay(rows) };
    return { broker: 'yuanta', items: parseUnrealizedYuanta(rows) };
  }

  function parseTrades(workbook) {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheetToRows(ws);
    const fmt = detectFormat(rows, 'trades');
    if (fmt === 'cathay') return { broker: 'cathay', items: parseTradesCathay(rows) };
    return { broker: 'yuanta', items: parseTradesYuanta(rows) };
  }

  function parseRealized(workbook) {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const rows = sheetToRows(ws);
    const fmt = detectFormat(rows, 'realized');
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
  // ============================================================
  function enrichRealizedWithInterest(realizedItems, tradeItems) {
    // 1. 先把現股的也清成 0（避免殘留舊值）
    for (const r of realizedItems) {
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
    const realizedBuckets = new Map();
    for (const r of realizedItems) {
      if (r.sellCategory !== '融資' && r.sellCategory !== '融券') continue;
      const k = `${r.code}|${r.sellDate}|${r.sellCategory}`;
      if (!realizedBuckets.has(k)) realizedBuckets.set(k, []);
      realizedBuckets.get(k).push(r);
    }

    let matched = 0;
    let total = 0;

    // 4. 組對組消耗：同一個 key 下，逐一從 tradeBucket 取出 trade，
    //    按 trade.qty 從 realizedBucket 的開頭吃掉等量的數量
    for (const [key, rzGroup] of realizedBuckets) {
      total += rzGroup.length;
      const trades = tradeBuckets.get(key) || [];
      if (!trades.length) continue;

      // realized 群組依出現順序消耗
      const queue = rzGroup.slice();
      // 嘗試對每張交易消耗對應數量
      for (const trade of trades) {
        let remaining = trade.qty;
        const consumed = []; // 這張 trade 涵蓋了哪幾筆 realized
        let costSum = 0;
        while (remaining > 0 && queue.length > 0) {
          const r = queue[0];
          if (r.qty <= remaining) {
            consumed.push(r);
            costSum += (r.cost || 0);
            remaining -= r.qty;
            queue.shift();
          } else {
            // 已實現的單筆數量比 trade 剩下的還多 → 部分配對
            // 這種情況代表已實現拆得比交易明細還細（少見）
            // 退回去：把剩下的當作「無法精確配對」直接給整筆
            consumed.push(r);
            costSum += (r.cost || 0);
            remaining = 0;
            queue.shift();
          }
        }
        if (consumed.length === 0) continue;

        // 按沖銷成本比例分配利息和融券費
        const interest = trade.marginInterest || 0;
        const sFee = trade.shortFee || 0;
        if (consumed.length === 1) {
          consumed[0].interest = interest;
          consumed[0].shortFee = sFee;
          consumed[0]._unmatched = false;
          matched++;
        } else {
          // 用比例分；用「沖銷成本」分，最後一筆吃尾差
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

  return {
    parseUnrealized,
    parseTrades,
    parseRealized,
    enrichRealizedWithInterest,
    detectFormat,
    toNumber,
    toString,
    formatDate
  };
})();
