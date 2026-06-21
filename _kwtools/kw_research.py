"""
kw_research_v2.py
GAS側の themes シートから未リサーチテーマを読み、
SerpAPI + Google Ads API でKWデータを取得して、
kw_research シートに書き込む。
"""
import sys
import os
import logging
import time
import json
import re
from datetime import datetime, timezone, timedelta

# ★ .env 読み込み(ローカル実行用、GitHub Actions では環境変数が直接設定される)
from dotenv import load_dotenv
try:
    from dotenv import load_dotenv
    load_dotenv('miroca_kw_research.env', override=True)
except Exception:
    pass

import requests
import gspread
from google.oauth2.service_account import Credentials
from google.ads.googleads.client import GoogleAdsClient
from google.ads.googleads.errors import GoogleAdsException

# ==========================================
# 設定
# ==========================================
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID', '')  # 対象スプレッドシートID
SERPAPI_KEY = os.environ.get('SERPAPI_KEY', '')
GOOGLE_ADS_CUSTOMER_ID = os.environ.get('GOOGLE_ADS_CUSTOMER_ID', '')

# サービスアカウントJSON のパス or 内容(GitHub Actionsでは内容を環境変数で渡す)
SERVICE_ACCOUNT_FILE = os.environ.get('GOOGLE_SHEETS_CREDENTIALS', 'service_account.json')

# Google Ads YAML のパス or 内容
GOOGLE_ADS_YAML_PATH = os.environ.get('GOOGLE_ADS_YAML_PATH', 'google-ads.yaml')

# 1回の実行で処理する最大テーマ数(API制限回避のため)
MAX_THEMES_PER_RUN = int(os.environ.get('MAX_THEMES_PER_RUN', '10'))

# リフレッシュ判定: 何日前までのデータを「リサーチ済み」と扱うか
RESEARCH_FRESHNESS_DAYS = 7

# ロギング
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)
logger = logging.getLogger(__name__)

# ==========================================
# Google Sheets 接続
# ==========================================
def get_gspread_client():
    """サービスアカウントで gspread クライアントを生成"""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
    ]
    
    # GitHub Actions 環境では内容を直接渡す場合と、ファイルパスの場合がある
    if SERVICE_ACCOUNT_FILE and os.path.exists(SERVICE_ACCOUNT_FILE):
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    elif os.environ.get('GOOGLE_SHEETS_CREDENTIALS_JSON'):
        # 環境変数に JSON 内容が入っている場合(GitHub Actions)
        info = json.loads(os.environ['GOOGLE_SHEETS_CREDENTIALS_JSON'])
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    else:
        raise RuntimeError('サービスアカウント認証情報が見つかりません')
    
    return gspread.authorize(creds)

# ==========================================
# シート読み込み: 未リサーチテーマ抽出
# ==========================================
def extract_seed_keyword(theme):
    """decide_themes 用: target_queries の先頭クエリを seed にする。
    無ければ review_keywords、それも無ければ title の先頭語にフォールバック。"""
    # target_queries は "クエリA,クエリB,クエリC" のカンマ区切り
    tq = theme.get('target_queries', '') or ''
    if tq.strip():
        first = tq.split(',')[0].strip()
        if first:
            return first

    # フォールバック1: review_keywords の先頭
    rk = theme.get('review_keywords', '') or ''
    if rk.strip():
        first = rk.split(',')[0].strip()
        if first:
            # 単語1つだと弱いので "ハウスクリーニング" を補う
            return f'ハウスクリーニング {first}' if len(first) <= 4 else first

    # フォールバック2: title 先頭15文字（記号除去）
    title = theme.get('title', '') or ''
    return re.sub(r'[【】｜\|\[\]()()？?、。｜]', ' ', title)[:15].strip()

def get_themes_needing_research(gc, spreadsheet_id):
    """themes シートから未リサーチのテーマを取得"""
    sh = gc.open_by_key(spreadsheet_id)
    themes_ws = sh.worksheet('decide_themes')
    
    # 全データを辞書リストで取得
    all_themes = themes_ws.get_all_records()
    
    # 既リサーチの seed_keyword を取得
    try:
        kw_ws = sh.worksheet('kw_research')
        kw_records = kw_ws.get_all_records()
    except gspread.WorksheetNotFound:
        # kw_research シートがなければ作る
        kw_ws = sh.add_worksheet(title='kw_research', rows=100, cols=10)
        kw_ws.update('A1:G1', [[
            'seed_keyword', 'fetched_at', 'paa_questions', 'related_searches',
            'ads_keywords', 'status', 'note'
        ]])
        kw_records = []
    
    # 7日以内にリサーチ済みの seed_keyword セット
    researched = set()
    now = datetime.now(timezone.utc)
    for rec in kw_records:
        fetched = rec.get('fetched_at', '')
        if not fetched:
            continue
        try:
            fetched_dt = datetime.fromisoformat(str(fetched).replace('Z', '+00:00'))
            age_days = (now - fetched_dt).total_seconds() / 86400
            if age_days < RESEARCH_FRESHNESS_DAYS:
                researched.add(rec['seed_keyword'].strip())
        except (ValueError, TypeError):
            continue
    
    # 対象テーマ抽出
    needing = []
    for t in all_themes:
        status = t.get('status', '')
        if status not in ('approved', 'ready', 'draft', 'generated'):
            continue
        if str(t.get('reject_reason', '')).strip():   # リジェクト理由があれば除外
            continue
        
        seed_kw = extract_seed_keyword(t)
        if not seed_kw or seed_kw in researched:
            continue
        
        # 重複除外(同じseed_kwが複数テーマに対応する場合は1回だけ処理)
        if any(n['seed_keyword'] == seed_kw for n in needing):
            continue
        
        needing.append({
            'slug': t.get('slug', ''),
            'title': t.get('title', ''),
            'seed_keyword': seed_kw,
            'phase': t.get('phase', ''),
        })
    
    return needing, sh, kw_ws

# ==========================================
# SerpAPI
# ==========================================
def get_serpapi_data(keyword):
    logger.info(f'[SerpAPI] {keyword}')
    url = 'https://serpapi.com/search'
    params = {
        'engine': 'google',
        'q': keyword,
        'hl': 'ja',
        'gl': 'jp',
        'api_key': SERPAPI_KEY,
    }
    
    result = {'paa': [], 'related': []}
    try:
        r = requests.get(url, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        
        if 'related_questions' in data:
            result['paa'] = [q.get('question') for q in data['related_questions'] if 'question' in q]
        if 'related_searches' in data:
            result['related'] = [s.get('query') for s in data['related_searches'] if 'query' in s]
    except Exception as e:
        logger.error(f'[SerpAPI] エラー: {e}')
    
    return result

# ==========================================
# Google Ads API
# ==========================================
def get_ads_api_data(client, customer_id, keyword):
    logger.info(f'[AdsAPI] {keyword}')
    result_list = []
    
    try:
        service = client.get_service('KeywordPlanIdeaService')
        request = client.get_type('GenerateKeywordIdeasRequest')
        request.customer_id = customer_id
        request.language = client.get_service('GoogleAdsService').language_constant_path('1005')
        request.geo_target_constants.append(
            client.get_service('GoogleAdsService').geo_target_constant_path('2392')
        )
        request.keyword_plan_network = client.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH_AND_PARTNERS
        request.keyword_seed.keywords.append(keyword)
        
        ideas = service.generate_keyword_ideas(request=request)
        
        for idea in ideas:
            m = idea.keyword_idea_metrics
            low = m.low_top_of_page_bid_micros / 1_000_000 if m.low_top_of_page_bid_micros else 0
            high = m.high_top_of_page_bid_micros / 1_000_000 if m.high_top_of_page_bid_micros else 0
            cpc = int((low + high) / 2)  # 低・高の平均をCPCとする
            
            result_list.append({
                'keyword': idea.text,
                'volume': m.avg_monthly_searches or 0,
                'competition': m.competition.name,
                'cpc': cpc,
            })
    except GoogleAdsException as ex:
        logger.error(f'[AdsAPI] エラー: {ex.error.code().name} - {ex.failure.errors[0].message}')
    except Exception as e:
        logger.error(f'[AdsAPI] 予期せぬエラー: {e}')
    
    return result_list

# ==========================================
# kw_research シートに書き込み
# ==========================================
def save_to_kw_research_sheet(kw_ws, seed_kw, paa, related, ads):
    """kw_research シートに書き込み(重複時は上書き)"""
    now_iso = datetime.now(timezone.utc).isoformat()
    
    # 既存行を検索
    existing = kw_ws.get_all_records()
    target_row = None
    for i, rec in enumerate(existing, start=2):  # 1-indexed + ヘッダ
        if rec.get('seed_keyword', '').strip() == seed_kw.strip():
            target_row = i
            break
    
    row_values = [
        seed_kw,
        now_iso,
        json.dumps(paa, ensure_ascii=False),
        json.dumps(related, ensure_ascii=False),
        json.dumps(ads, ensure_ascii=False),
        'ok',
        f'ads={len(ads)} paa={len(paa)} related={len(related)}'
    ]
    
    if target_row:
        kw_ws.update(f'A{target_row}:G{target_row}', [row_values])
        logger.info(f'  → 上書き: row {target_row}')
    else:
        kw_ws.append_row(row_values)
        logger.info(f'  → 新規追加')

# ==========================================
# メイン
# ==========================================
def main():
    logger.info('=' * 60)
    logger.info('KWリサーチ開始')
    logger.info('=' * 60)
    
    # 設定確認
    if not SPREADSHEET_ID:
        logger.error('SPREADSHEET_ID が未設定')
        return 1
    if not SERPAPI_KEY:
        logger.error('SERPAPI_KEY が未設定')
        return 1
    if not GOOGLE_ADS_CUSTOMER_ID:
        logger.error('GOOGLE_ADS_CUSTOMER_ID が未設定')
        return 1
    
    # Google Sheets 接続
    gc = get_gspread_client()
    logger.info('Sheets 接続 OK')
    
    # 対象テーマ取得
    themes, sh, kw_ws = get_themes_needing_research(gc, SPREADSHEET_ID)
    logger.info(f'未リサーチテーマ: {len(themes)}件')
    
    if not themes:
        logger.info('処理すべきテーマなし。終了。')
        return 0
    
    # 件数制限
    themes = themes[:MAX_THEMES_PER_RUN]
    logger.info(f'今回処理: {len(themes)}件(上限 {MAX_THEMES_PER_RUN})')
    
    # Google Ads クライアント初期化
    try:
        ads_client = GoogleAdsClient.load_from_storage(
            path=GOOGLE_ADS_YAML_PATH,
            version='v23'
        )
    except Exception as e:
        logger.error(f'Google Ads API 初期化失敗: {e}')
        return 1
    
    # 各テーマ処理
    success = 0
    failed = 0
    
    for i, theme in enumerate(themes, 1):
        seed = theme['seed_keyword']
        logger.info(f'\n[{i}/{len(themes)}] {seed} (slug={theme["slug"]})')
        
        try:
            # SerpAPI
            serp = get_serpapi_data(seed)
            time.sleep(1)  # レート対策
            
            # Google Ads API
            ads = get_ads_api_data(ads_client, GOOGLE_ADS_CUSTOMER_ID, seed)
            time.sleep(1)
            
            # 保存
            save_to_kw_research_sheet(kw_ws, seed, serp['paa'], serp['related'], ads)
            success += 1
            
            logger.info(f'  ✓ paa={len(serp["paa"])} related={len(serp["related"])} ads={len(ads)}')
            
        except Exception as e:
            logger.error(f'  ✗ 失敗: {e}')
            failed += 1
    
    logger.info('\n' + '=' * 60)
    logger.info(f'完了: 成功={success} / 失敗={failed}')
    logger.info('=' * 60)
    
    return 0

if __name__ == '__main__':
    sys.exit(main())
