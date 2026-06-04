"""
Meta (Facebook/Instagram) 광고 성과 데이터 자동 수집 스크립트
Marketing API v25.0 사용
"""

import requests
import pandas as pd
from datetime import datetime
from pathlib import Path


# ── 설정 읽기 ──────────────────────────────────────────────
def load_config():
    config = {}
    env_path = Path(__file__).parent / '.env.meta'
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if '=' in line and not line.startswith('#'):
                key, val = line.split('=', 1)
                config[key.strip()] = val.strip()
    return config


# ── 액세스 토큰 장기화 (단기 → 60일) ──────────────────────
def extend_token(config):
    url = 'https://graph.facebook.com/oauth/access_token'
    params = {
        'grant_type': 'fb_exchange_token',
        'client_id':       config['META_APP_ID'],
        'client_secret':   config['META_APP_SECRET'],
        'fb_exchange_token': config['META_ACCESS_TOKEN'],
    }
    r = requests.get(url, params=params)
    r.raise_for_status()
    new_token = r.json().get('access_token')
    if new_token:
        # .env.meta 갱신
        env_path = Path(__file__).parent / '.env.meta'
        lines = env_path.read_text().splitlines()
        updated = []
        for line in lines:
            if line.startswith('META_ACCESS_TOKEN='):
                updated.append(f'META_ACCESS_TOKEN={new_token}')
            else:
                updated.append(line)
        env_path.write_text('\n'.join(updated) + '\n')
        print('🔑 토큰 장기화 완료 (60일 유효)')
        return new_token
    return config['META_ACCESS_TOKEN']


# ── 광고 인사이트 수집 ──────────────────────────────────────
def fetch_ad_insights(account_id, token, start_date, end_date):
    """일별 광고 성과 데이터 수집 (계정 전체 합산)"""
    url = f'https://graph.facebook.com/v25.0/{account_id}/insights'
    params = {
        'fields': 'impressions,clicks,spend,ctr,cpc,actions',
        'time_range': f'{{"since":"{start_date}","until":"{end_date}"}}',
        'time_increment': '1',   # 일별
        'level': 'account',
        'access_token': token,
    }

    rows = []
    next_url = url
    while next_url:
        r = requests.get(next_url, params=params if next_url == url else {})
        if r.status_code != 200:
            print(f'❌ API 오류: {r.text}')
            break
        data = r.json()
        rows.extend(data.get('data', []))
        next_url = data.get('paging', {}).get('next')

    return rows


# ── 전환 이벤트 파싱 ────────────────────────────────────────
def parse_actions(actions):
    """actions 배열에서 버튼 전환 수 추출"""
    if not actions:
        return 0
    # 링크 클릭, 랜딩페이지뷰 제외한 실질 전환
    exclude = {'link_click', 'landing_page_view', 'post_engagement', 'page_engagement'}
    total = 0
    for a in actions:
        if a.get('action_type') not in exclude:
            total += int(float(a.get('value', 0)))
    return total


# ── DataFrame 변환 ──────────────────────────────────────────
def to_dataframe(rows):
    if not rows:
        return pd.DataFrame()

    records = []
    for row in rows:
        records.append({
            '날짜': pd.to_datetime(row['date_start']),
            '매체': 'Meta',
            '노출': int(row.get('impressions', 0)),
            '클릭': int(row.get('clicks', 0)),
            '비용': float(row.get('spend', 0)),
            '전환': parse_actions(row.get('actions')),
            'CTR':  float(row.get('ctr', 0)),
            'CPC':  float(row.get('cpc', 0)),
        })

    df = pd.DataFrame(records)
    return df.sort_values('날짜').reset_index(drop=True)


# ── 메인 실행 ───────────────────────────────────────────────
def main(start_date=None, end_date=None, extend=True):
    config = load_config()

    # 토큰 장기화
    if extend:
        token = extend_token(config)
    else:
        token = config['META_ACCESS_TOKEN']

    account_id = config['META_AD_ACCOUNT_ID']

    # 날짜 기본값: 이번 달
    today = datetime.today()
    if not start_date:
        start_date = today.replace(day=1).strftime('%Y-%m-%d')
    if not end_date:
        end_date = today.strftime('%Y-%m-%d')

    print(f'📅 수집 기간: {start_date} ~ {end_date}')
    print(f'📦 광고 계정: {account_id}')

    rows = fetch_ad_insights(account_id, token, start_date, end_date)
    df = to_dataframe(rows)

    if df.empty:
        print('⚠️  수집된 데이터가 없습니다.')
        return None

    print(f'\n✅ {len(df)}일치 데이터 수집 완료\n')
    print(df.to_string(index=False))

    # CSV 저장
    fname = f'meta_광고성과_{today.strftime("%Y%m%d")}.csv'
    out = Path(__file__).parent / fname
    df.to_csv(out, index=False, encoding='utf-8-sig')
    print(f'\n💾 저장 완료: {out}')

    return df


if __name__ == '__main__':
    main()
