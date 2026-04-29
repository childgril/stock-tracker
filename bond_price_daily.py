"""
債券價格自動更新爬蟲
- 自動從 fcn_data.json Gist 讀取債券清單（含 CUSIP）
- 用 Playwright 抓 FINRA 公司債價格
- 結果寫入 fund_nav.json Gist 的 bonds 區塊
"""
import json, os, re, time, requests
from datetime import datetime

GITHUB_TOKEN  = os.environ['GITHUB_TOKEN']
GIST_NAV_ID   = os.environ.get('GIST_ID', '')
GIST_DATA_ID  = os.environ.get('GIST_DATA_ID', '')

GH_HEADERS = {
    'Authorization': 'token ' + GITHUB_TOKEN,
    'Accept': 'application/vnd.github.v3+json',
}


def load_bonds_from_gist():
    if not GIST_DATA_ID:
        return []
    try:
        r = requests.get(f'https://api.github.com/gists/{GIST_DATA_ID}', headers=GH_HEADERS, timeout=15)
        r.raise_for_status()
        content = r.json()['files']['fcn_data.json']['content']
        data = json.loads(content)
        bd_list = data.get('fcnBD', [])
        bonds = []
        seen = set()
        for bd in bd_list:
            cusip = bd.get('cusip', '')
            isin = bd.get('isin', '')
            key = cusip or isin
            if key and key not in seen:
                seen.add(key)
                bonds.append({
                    'cusip': cusip,
                    'isin': isin,
                    'name': bd.get('name', ''),
                })
        print(f'從 Gist 載入 {len(bonds)} 筆債券')
        return bonds
    except Exception as e:
        print(f'載入失敗：{e}')
        return []


def fetch_finra_price(cusip):
    """從 FINRA Bond Page 抓取最新價格"""
    if not cusip:
        return None
    try:
        from playwright.sync_api import sync_playwright
        url = f'https://www.finra.org/finra-data/fixed-income/bond?cusip={cusip}&bondType=CA'

        price_data = {}

        def handle_response(response):
            # FINRA 內部 API 通常是 ramp.finra.org
            if 'finra' in response.url and ('bond' in response.url.lower() or 'trace' in response.url.lower()):
                if 'json' in response.headers.get('content-type', ''):
                    try:
                        j = response.json()
                        # 嘗試找 last price / current price
                        if isinstance(j, dict):
                            for key in ['lastSalePrice', 'price', 'lastTradePrice', 'currentPrice']:
                                if j.get(key):
                                    price_data['price'] = float(j[key])
                                    return
                            # 也找 returnedRows 結構
                            rows = j.get('returnedRows') or j.get('data', [])
                            if rows and isinstance(rows, list):
                                for r in rows[:3]:
                                    for key in ['lastSalePrice','price','reportedPrice']:
                                        if r.get(key):
                                            price_data['price'] = float(r[key])
                                            return
                    except:
                        pass

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
            )
            page = context.new_page()
            page.on('response', handle_response)
            page.goto(url, wait_until='networkidle', timeout=45000)
            time.sleep(3)

            # 如果攔截不到 API，從 DOM 抓
            if not price_data.get('price'):
                try:
                    text = page.content()
                    # FINRA 頁面格式：找價格欄位
                    patterns = [
                        r'"lastSalePrice"\s*:\s*([\d.]+)',
                        r'"price"\s*:\s*([\d.]+)',
                        r'Last\s*Trade\s*Price[^\d]*\$?([\d.]+)',
                        r'Last\s*Price[^\d]*\$?([\d.]+)',
                    ]
                    for p_re in patterns:
                        m = re.search(p_re, text)
                        if m:
                            v = float(m.group(1))
                            if 1 < v < 200:
                                price_data['price'] = v
                                break
                except Exception as e:
                    print(f'  DOM parse error: {e}')

            browser.close()

        if price_data.get('price'):
            print(f'  ✓ FINRA: {price_data["price"]}')
            return price_data['price']
    except Exception as e:
        print(f'  Playwright error: {e}')
    return None


def save_to_gist(data):
    # 讀現有 nav，合併 bonds
    existing = {}
    if GIST_NAV_ID:
        try:
            r = requests.get(f'https://api.github.com/gists/{GIST_NAV_ID}', headers=GH_HEADERS, timeout=10)
            content = r.json().get('files', {}).get('fund_nav.json', {}).get('content', '')
            if content:
                existing = json.loads(content)
        except:
            pass

    # 合併
    output = {**existing}
    output['updated_at'] = datetime.now().isoformat()
    output['bonds'] = data

    content = json.dumps(output, ensure_ascii=False, indent=2)
    payload = {
        'description': 'FCN投資系統 - 基金與債券價格',
        'public': False,
        'files': {'fund_nav.json': {'content': content}},
    }
    if GIST_NAV_ID:
        r = requests.patch(f'https://api.github.com/gists/{GIST_NAV_ID}', headers=GH_HEADERS, json=payload, timeout=15)
    else:
        r = requests.post('https://api.github.com/gists', headers=GH_HEADERS, json=payload, timeout=15)
    r.raise_for_status()
    print(f'Gist 已更新')


def main():
    print(f'=== 債券價格更新 {datetime.now().strftime("%Y-%m-%d %H:%M")} ===\n')
    bonds = load_bonds_from_gist()
    if not bonds:
        print('沒有債券資料，跳過')
        return

    today = datetime.now().strftime('%Y-%m-%d')
    results = {}

    for b in bonds:
        cusip = b['cusip']
        print(f'抓取：{b["name"]} (CUSIP: {cusip})')
        price = fetch_finra_price(cusip)
        if price:
            for key in [b['cusip'], b['isin']]:
                if key:
                    results[key] = {
                        'price': price,
                        'date': today,
                        'name': b['name'],
                    }
        else:
            print('  ✗ 失敗')
        print()

    save_to_gist(results)
    print(f'完成：{len([r for r in results.values()])}/{len(bonds)} 筆')


main()
